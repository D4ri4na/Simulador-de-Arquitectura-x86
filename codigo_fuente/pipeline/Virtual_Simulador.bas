Attribute VB_Name = "Virtual_Simulador"
' Simulador de Memoria Virtual - Módulo Único
Type MemoryCell
    address As String
    value As String
    instruction As String
    dataType As String ' "INSTR", "DATA", "STACK", "FREE"
    Accessed As Boolean
    Modified As Boolean
End Type

' Variables globales de memoria
Dim VirtualMemory() As MemoryCell
Dim MemorySize As Long
Dim ProgramCounter As Long
Dim StackPointer As Long

' Registros del sistema
Dim AX As Long, BX As Long, CX As Long, DX As Long
Dim SI As Long, DI As Long, BP As Long, SP As Long
Dim Flags As Long
Dim IsRunning As Boolean

' =============================================
' INICIALIZACIÓN DE MEMORIA
' =============================================

Sub InitializeVirtualMemory(Optional size As Long = 512)
    MemorySize = size
    ReDim VirtualMemory(0 To MemorySize - 1)
    
    ' Inicializar punteros
    ProgramCounter = 0
    StackPointer = MemorySize - 1
    
    ' Inicializar registros
    AX = 0: BX = 0: CX = 0: DX = 0
    SI = 0: DI = 0: BP = 0: SP = MemorySize - 1
    Flags = 0
    IsRunning = False
    
    ' Limpiar memoria
    Dim i As Long
    For i = 0 To MemorySize - 1
        VirtualMemory(i).address = "0x" & Format(Hex(i), "0000")
        VirtualMemory(i).value = "00"
        VirtualMemory(i).instruction = ""
        VirtualMemory(i).dataType = "FREE"
        VirtualMemory(i).Accessed = False
        VirtualMemory(i).Modified = False
    Next i
    
    ' Inicializar área de stack
    For i = MemorySize - 50 To MemorySize - 1
        VirtualMemory(i).dataType = "STACK"
        VirtualMemory(i).value = "00"
    Next i
    
    CreateMemoryDisplay
    UpdateRegisterDisplay
    MsgBox "Memoria virtual inicializada: " & MemorySize & " bytes", vbInformation
End Sub

' =============================================
' OPERACIONES BÁSICAS DE MEMORIA
' =============================================

Function WriteMemory(address As Long, value As String, Optional dataType As String = "DATA") As Boolean
    If address < 0 Or address >= MemorySize Then
        MsgBox "Error: Dirección fuera de rango - " & address, vbCritical
        WriteMemory = False
        Exit Function
    End If
    
    VirtualMemory(address).value = value
    VirtualMemory(address).dataType = dataType
    VirtualMemory(address).Modified = True
    VirtualMemory(address).Accessed = True
    WriteMemory = True
    
    UpdateMemoryDisplay
End Function

Function ReadMemory(address As Long) As String
    If address < 0 Or address >= MemorySize Then
        MsgBox "Error: Dirección fuera de rango - " & address, vbCritical
        ReadMemory = "ERROR"
        Exit Function
    End If
    
    VirtualMemory(address).Accessed = True
    ReadMemory = VirtualMemory(address).value
    UpdateMemoryDisplay
End Function

' =============================================
' CARGA DE PROGRAMAS
' =============================================

Sub LoadProgramFromCodeSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CodigoFuente")
    
    ' Encontrar última fila con código
    Dim lastRow As Long
    lastRow = 0
    Dim i As Long
    For i = 1 To 1000
        If Trim(ws.Cells(i, 1).value) = "" Then
            lastRow = i - 1
            Exit For
        End If
    Next i
    
    If lastRow = 0 Then lastRow = ws.Cells(ws.ROWS.count, "A").End(xlUp).row
    
    ' Limpiar área de instrucciones anterior
    For i = 0 To MemorySize - 1
        If VirtualMemory(i).dataType = "INSTR" Then
            VirtualMemory(i).instruction = ""
            VirtualMemory(i).value = "00"
            VirtualMemory(i).dataType = "FREE"
        End If
    Next i
    
    ' Cargar nuevas instrucciones
    Dim instructionCount As Long
    instructionCount = 0
    
    For i = 1 To lastRow
        Dim instruction As String
        instruction = Trim(ws.Cells(i, 1).value)
        
        If instruction <> "" Then
            If WriteMemory(i - 1, instruction, "INSTR") Then
                VirtualMemory(i - 1).instruction = instruction
                instructionCount = instructionCount + 1
            End If
        End If
    Next i
    
    ProgramCounter = 0
    UpdateMemoryDisplay
    MsgBox "Programa cargado: " & instructionCount & " instrucciones", vbInformation
End Sub

' =============================================
' EJECUCIÓN DE INSTRUCCIONES
' =============================================

Sub ExecuteFullProgram()
    If Not IsRunning Then
        InitializeRegisters
        ProgramCounter = 0
        IsRunning = True
    End If
    
    CreateExecutionTrace
    Dim stepCount As Long
    stepCount = 0
    Dim maxSteps As Long
    maxSteps = 1000
    
    While IsRunning And ProgramCounter < MemorySize And stepCount < maxSteps
        stepCount = stepCount + 1
        If Not ExecuteNextStep(stepCount) Then
            Exit While
        End If
        DoEvents
    Wend
    
    If stepCount >= maxSteps Then
        AddToExecutionTrace "EJECUCIÓN DETENIDA: Límite de pasos excedido", stepCount + 2
    Else
        AddToExecutionTrace "EJECUCIÓN COMPLETADA", stepCount + 2
    End If
    
    IsRunning = False
    UpdateRegisterDisplay
End Sub

Function ExecuteNextStep(stepNumber As Long) As Boolean
    If ProgramCounter >= MemorySize Or VirtualMemory(ProgramCounter).instruction = "" Then
        IsRunning = False
        ExecuteNextStep = False
        Exit Function
    End If
    
    Dim instruction As String
    instruction = VirtualMemory(ProgramCounter).instruction
    
    ' Mostrar en traza de ejecución
    AddToExecutionTrace "PC=" & ProgramCounter & ": " & instruction, stepNumber + 1
    
    ' Parsear y ejecutar
    If ParseAndExecuteInstruction(instruction) Then
        ExecuteNextStep = True
    Else
        AddToExecutionTrace "ERROR en instrucción", stepNumber + 1
        IsRunning = False
        ExecuteNextStep = False
    End If
End Function

Function ParseAndExecuteInstruction(instruction As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Limpiar instrucción
    Dim cleanInstruction As String
    cleanInstruction = Split(instruction, ";")(0)
    
    Dim parts() As String
    parts = Split(Trim(cleanInstruction), " ")
    
    If UBound(parts) < 0 Then
        ProgramCounter = ProgramCounter + 1
        ParseAndExecuteInstruction = True
        Exit Function
    End If
    
    Dim opcode As String
    opcode = UCase(Trim(parts(0)))
    
    Dim operand1 As String, operand2 As String
    operand1 = ""
    operand2 = ""
    
    If UBound(parts) >= 1 Then operand1 = Trim(parts(1))
    If UBound(parts) >= 2 Then operand2 = Trim(parts(2))
    
    ' Ejecutar instrucción
    Select Case opcode
        Case "MOV"
            ParseAndExecuteInstruction = ExecuteMOV(operand1, operand2)
        Case "ADD"
            ParseAndExecuteInstruction = ExecuteADD(operand1, operand2)
        Case "SUB"
            ParseAndExecuteInstruction = ExecuteSUB(operand1, operand2)
        Case "MUL"
            ParseAndExecuteInstruction = ExecuteMUL(operand1)
        Case "DIV"
            ParseAndExecuteInstruction = ExecuteDIV(operand1)
        Case "INC"
            ParseAndExecuteInstruction = ExecuteINC(operand1)
        Case "DEC"
            ParseAndExecuteInstruction = ExecuteDEC(operand1)
        Case "JMP"
            ParseAndExecuteInstruction = ExecuteJMP(operand1)
        Case "JZ", "JE"
            ParseAndExecuteInstruction = ExecuteJZ(operand1)
        Case "JNZ", "JNE"
            ParseAndExecuteInstruction = ExecuteJNZ(operand1)
        Case "CALL"
            ParseAndExecuteInstruction = ExecuteCALL(operand1)
        Case "RET"
            ParseAndExecuteInstruction = ExecuteRET()
        Case "PUSH"
            ParseAndExecuteInstruction = ExecutePUSH(operand1)
        Case "POP"
            ParseAndExecuteInstruction = ExecutePOP(operand1)
        Case "CMP"
            ParseAndExecuteInstruction = ExecuteCMP(operand1, operand2)
        Case "HLT"
            IsRunning = False
            ParseAndExecuteInstruction = True
        Case "NOP"
            ProgramCounter = ProgramCounter + 1
            ParseAndExecuteInstruction = True
        Case Else
            MsgBox "Instrucción no reconocida: " & opcode, vbExclamation
            ParseAndExecuteInstruction = False
    End Select
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error ejecutando: " & instruction, vbCritical
    ParseAndExecuteInstruction = False
End Function

' =============================================
' IMPLEMENTACIÓN DE INSTRUCCIONES
' =============================================

Function ExecuteMOV(dest As String, src As String) As Boolean
    Dim value As Long
    value = GetOperandValue(src)
    
    If SetRegisterValue(dest, value) Then
        ProgramCounter = ProgramCounter + 1
        ExecuteMOV = True
    Else
        ExecuteMOV = False
    End If
    UpdateRegisterDisplay
End Function

Function ExecuteADD(dest As String, src As String) As Boolean
    Dim regValue As Long, srcValue As Long
    regValue = GetRegisterValue(dest)
    srcValue = GetOperandValue(src)
    
    If SetRegisterValue(dest, regValue + srcValue) Then
        UpdateFlags (regValue + srcValue)
        ProgramCounter = ProgramCounter + 1
        ExecuteADD = True
    Else
        ExecuteADD = False
    End If
    UpdateRegisterDisplay
End Function

Function ExecuteSUB(dest As String, src As String) As Boolean
    Dim regValue As Long, srcValue As Long
    regValue = GetRegisterValue(dest)
    srcValue = GetOperandValue(src)
    
    If SetRegisterValue(dest, regValue - srcValue) Then
        UpdateFlags (regValue - srcValue)
        ProgramCounter = ProgramCounter + 1
        ExecuteSUB = True
    Else
        ExecuteSUB = False
    End If
    UpdateRegisterDisplay
End Function

Function ExecuteJMP(address As String) As Boolean
    Dim jumpAddr As Long
    jumpAddr = Val(address)
    
    If jumpAddr >= 0 And jumpAddr < MemorySize Then
        ProgramCounter = jumpAddr
        ExecuteJMP = True
    Else
        ExecuteJMP = False
    End If
End Function

Function ExecutePUSH(value As String) As Boolean
    Dim pushValue As Long
    pushValue = GetOperandValue(value)
    
    If StackPointer > 0 Then
        StackPointer = StackPointer - 1
        WriteMemory StackPointer, CStr(pushValue), "STACK"
        SP = StackPointer
        ProgramCounter = ProgramCounter + 1
        ExecutePUSH = True
    Else
        ExecutePUSH = False
    End If
    UpdateRegisterDisplay
End Function

Function ExecutePOP(dest As String) As Boolean
    If StackPointer < MemorySize - 1 Then
        Dim popValue As String
        popValue = ReadMemory(StackPointer)
        StackPointer = StackPointer + 1
        SP = StackPointer
        
        If SetRegisterValue(dest, Val(popValue)) Then
            ProgramCounter = ProgramCounter + 1
            ExecutePOP = True
        Else
            ExecutePOP = False
        End If
    Else
        ExecutePOP = False
    End If
    UpdateRegisterDisplay
End Function

Function ExecuteCALL(address As String) As Boolean
    Dim callAddr As Long
    callAddr = Val(address)
    
    If StackPointer > 0 And callAddr >= 0 And callAddr < MemorySize Then
        ' Guardar dirección de retorno
        StackPointer = StackPointer - 1
        WriteMemory StackPointer, CStr(ProgramCounter + 1), "STACK"
        SP = StackPointer
        
        ' Saltar a la dirección
        ProgramCounter = callAddr
        ExecuteCALL = True
    Else
        ExecuteCALL = False
    End If
    UpdateRegisterDisplay
End Function

Function ExecuteRET() As Boolean
    If StackPointer < MemorySize - 1 Then
        Dim returnAddr As String
        returnAddr = ReadMemory(StackPointer)
        StackPointer = StackPointer + 1
        SP = StackPointer
        
        ProgramCounter = Val(returnAddr)
        ExecuteRET = True
    Else
        ExecuteRET = False
    End If
    UpdateRegisterDisplay
End Function

' =============================================
' FUNCIONES AUXILIARES
' =============================================

Function GetRegisterValue(regName As String) As Long
    Select Case UCase(regName)
        Case "AX": GetRegisterValue = AX
        Case "BX": GetRegisterValue = BX
        Case "CX": GetRegisterValue = CX
        Case "DX": GetRegisterValue = DX
        Case "SI": GetRegisterValue = SI
        Case "DI": GetRegisterValue = DI
        Case "BP": GetRegisterValue = BP
        Case "SP": GetRegisterValue = SP
        Case Else: GetRegisterValue = 0
    End Select
End Function

Function SetRegisterValue(regName As String, value As Long) As Boolean
    Select Case UCase(regName)
        Case "AX": AX = value
        Case "BX": BX = value
        Case "CX": CX = value
        Case "DX": DX = value
        Case "SI": SI = value
        Case "DI": DI = value
        Case "BP": BP = value
        Case "SP": SP = value
        Case Else: SetRegisterValue = False: Exit Function
    End Select
    SetRegisterValue = True
End Function

Function GetOperandValue(operand As String) As Long
    If Left(operand, 1) = "[" And Right(operand, 1) = "]" Then
        ' Es una referencia a memoria
        Dim memAddr As Long
        memAddr = Val(Mid(operand, 2, Len(operand) - 2))
        GetOperandValue = Val(ReadMemory(memAddr))
    Else
        ' Es un valor inmediato o registro
        GetOperandValue = Val(operand)
    End If
End Function

Sub UpdateFlags(value As Long)
    ' Bandera Zero
    If value = 0 Then
        Flags = Flags Or 1
    Else
        Flags = Flags And Not 1
    End If
    
    ' Bandera Sign (negativo)
    If value < 0 Then
        Flags = Flags Or 2
    Else
        Flags = Flags And Not 2
    End If
End Sub

Sub InitializeRegisters()
    AX = 0: BX = 0: CX = 0: DX = 0
    SI = 0: DI = 0: BP = 0: SP = MemorySize - 1
    Flags = 0
    StackPointer = MemorySize - 1
    UpdateRegisterDisplay
End Sub

' =============================================
' INTERFAZ VISUAL
' =============================================

Sub CreateMemoryDisplay()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("MemoriaVirtual")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "MemoriaVirtual"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Cells(1, 1).value = "DIRECCIÓN"
    ws.Cells(1, 2).value = "VALOR"
    ws.Cells(1, 3).value = "INSTRUCCIÓN"
    ws.Cells(1, 4).value = "TIPO"
    ws.Cells(1, 5).value = "ACCEDIDO"
    ws.Cells(1, 6).value = "MODIFICADO"
    
    UpdateMemoryDisplay
End Sub

Sub UpdateMemoryDisplay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MemoriaVirtual")
    
    Dim row As Long
    row = 2
    Dim i As Long
    
    For i = 0 To MemorySize - 1
        If VirtualMemory(i).dataType <> "FREE" Or VirtualMemory(i).Accessed Then
            ws.Cells(row, 1).value = VirtualMemory(i).address
            ws.Cells(row, 2).value = VirtualMemory(i).value
            ws.Cells(row, 3).value = VirtualMemory(i).instruction
            ws.Cells(row, 4).value = VirtualMemory(i).dataType
            ws.Cells(row, 5).value = IIf(VirtualMemory(i).Accessed, "?", "")
            ws.Cells(row, 6).value = IIf(VirtualMemory(i).Modified, "?", "")
            row = row + 1
        End If
    Next i
    
    ws.Columns.AutoFit
End Sub

Sub UpdateRegisterDisplay()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Registros")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Registros"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Cells(1, 1).value = "REGISTRO"
    ws.Cells(1, 2).value = "VALOR"
    ws.Cells(1, 3).value = "HEX"
    
    ws.Cells(2, 1).value = "AX": ws.Cells(2, 2).value = AX: ws.Cells(2, 3).value = "0x" & Hex(AX)
    ws.Cells(3, 1).value = "BX": ws.Cells(3, 2).value = BX: ws.Cells(3, 3).value = "0x" & Hex(BX)
    ws.Cells(4, 1).value = "CX": ws.Cells(4, 2).value = CX: ws.Cells(4, 3).value = "0x" & Hex(CX)
    ws.Cells(5, 1).value = "DX": ws.Cells(5, 2).value = DX: ws.Cells(5, 3).value = "0x" & Hex(DX)
    ws.Cells(6, 1).value = "SI": ws.Cells(6, 2).value = SI: ws.Cells(6, 3).value = "0x" & Hex(SI)
    ws.Cells(7, 1).value = "DI": ws.Cells(7, 2).value = DI: ws.Cells(7, 3).value = "0x" & Hex(DI)
    ws.Cells(8, 1).value = "BP": ws.Cells(8, 2).value = BP: ws.Cells(8, 3).value = "0x" & Hex(BP)
    ws.Cells(9, 1).value = "SP": ws.Cells(9, 2).value = SP: ws.Cells(9, 3).value = "0x" & Hex(SP)
    
    ws.Cells(11, 1).value = "Program Counter"
    ws.Cells(11, 2).value = ProgramCounter
    ws.Cells(12, 1).value = "Stack Pointer"
    ws.Cells(12, 2).value = StackPointer
    ws.Cells(13, 1).value = "Flags"
    ws.Cells(13, 2).value = Flags
    ws.Cells(13, 3).value = "0x" & Hex(Flags)
    
    ws.Columns.AutoFit
End Sub

Sub CreateExecutionTrace()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TrazaEjecucion")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "TrazaEjecucion"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Cells(1, 1).value = "PASO"
    ws.Cells(1, 2).value = "INSTRUCCIÓN"
    ws.Cells(1, 3).value = "AX"
    ws.Cells(1, 4).value = "BX"
    ws.Cells(1, 5).value = "CX"
    ws.Cells(1, 6).value = "DX"
    ws.Cells(1, 7).value = "ESTADO"
End Sub

Sub AddToExecutionTrace(instruction As String, stepNumber As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TrazaEjecucion")
    
    ws.Cells(stepNumber, 1).value = stepNumber - 1
    ws.Cells(stepNumber, 2).value = instruction
    ws.Cells(stepNumber, 3).value = AX
    ws.Cells(stepNumber, 4).value = BX
    ws.Cells(stepNumber, 5).value = CX
    ws.Cells(stepNumber, 6).value = DX
    ws.Cells(stepNumber, 7).value = "OK"
    
    ws.Columns.AutoFit
End Sub

' =============================================
' FUNCIONES DE CONTROL PRINCIPALES
' =============================================

Sub IniciarSimulador()
    InitializeVirtualMemory 256
    LoadProgramFromCodeSheet
    MsgBox "Simulador listo. Use 'EjecutarPrograma' para comenzar.", vbInformation
End Sub

Sub EjecutarPrograma()
    ExecuteFullProgram
End Sub

Sub PasoAPaso()
    If Not IsRunning Then
        InitializeRegisters
        ProgramCounter = 0
        IsRunning = True
        CreateExecutionTrace
    End If
    
    If ExecuteNextStep(GetLastExecutionStep() + 1) Then
        UpdateRegisterDisplay
    Else
        IsRunning = False
        MsgBox "Ejecución terminada", vbInformation
    End If
End Sub

Function GetLastExecutionStep() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TrazaEjecucion")
    GetLastExecutionStep = ws.Cells(ws.ROWS.count, 1).End(xlUp).row - 1
End Function

' Instrucciones adicionales (simplificadas)
Function ExecuteMUL(operand As String) As Boolean
    Dim value As Long
    value = GetOperandValue(operand)
    AX = AX * value
    UpdateFlags (AX)
    ProgramCounter = ProgramCounter + 1
    ExecuteMUL = True
End Function

Function ExecuteINC(operand As String) As Boolean
    If SetRegisterValue(operand, GetRegisterValue(operand) + 1) Then
        UpdateFlags (GetRegisterValue(operand))
        ProgramCounter = ProgramCounter + 1
        ExecuteINC = True
    Else
        ExecuteINC = False
    End If
End Function

Function ExecuteJZ(operand As String) As Boolean
    If (Flags And 1) = 1 Then ' Zero flag set
        ExecuteJMP operand
    Else
        ProgramCounter = ProgramCounter + 1
        ExecuteJZ = True
    End If
End Function

Function ExecuteCMP(op1 As String, op2 As String) As Boolean
    Dim val1 As Long, val2 As Long
    val1 = GetOperandValue(op1)
    val2 = GetOperandValue(op2)
    UpdateFlags (val1 - val2)
    ProgramCounter = ProgramCounter + 1
    ExecuteCMP = True
End Function
