Attribute VB_Name = "RAM_Simulator"
' Módulo: RAM_Simulator_NASM (Versión Optimizada con Carga Dinámica)

Option Explicit

Private Const RAM_SIZE As Integer = 256
Private Const ROWS As Integer = 16
Private Const COLS As Integer = 16

Private RAM(0 To RAM_SIZE - 1) As Byte
Private RAM_Display(0 To RAM_SIZE - 1) As String
Private RAM_Loaded(0 To RAM_SIZE - 1) As Boolean ' Nueva: Rastrear qué se ha cargado
Private Program() As AssemblyInstruction
Private DataSectionStart As Integer
Private TextSectionStart As Integer
Private CurrentInstruction As Integer

Type AssemblyInstruction
    address As Integer
    OriginalLine As String
    opcode As String
    Operand1 As String
    Operand2 As String
    Operand3 As String
    bytes As String
    Length As Integer
    section As String
End Type

' ============= INICIALIZACIÓN =============
Sub InitializeRAMSimulatorNASM()
    Application.ScreenUpdating = False ' OPTIMIZACIÓN
    Application.Calculation = xlCalculationManual
    
    ClearRAM
    ReadNASMProgramFromCells
    DrawRAMGridNASM
    
    ' NO cargar el programa en RAM todavía
    ' Solo mostrar la información del programa
    DisplayNASMProgramInfo
    
    UpdateRAMDisplayNASM_Improved
    UpdateNASMStatus "Listo", "---", "---", "Presione 'Siguiente' para comenzar"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Simulador inicializado. La memoria se llenará al ejecutar el programa."
End Sub

Sub ClearRAM()
    Dim i As Integer
    For i = 0 To RAM_SIZE - 1
        RAM(i) = 0
        RAM_Display(i) = "00"
        RAM_Loaded(i) = False ' Nueva: Marcar como no cargado
    Next i
End Sub

' ============= LECTURA Y PARSEO (Sin cambios significativos) =============
Sub ReadNASMProgramFromCells()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ProgramaNASM")
    
    Dim lastRow As Integer
    lastRow = ws.Cells(ws.ROWS.count, "A").End(xlUp).row
    
    If lastRow < 2 Then
        MsgBox "No se encontró programa en las celdas."
        Exit Sub
    End If
    
    Dim validLineCount As Integer
    validLineCount = CountValidNASMLines(ws, lastRow)
    
    If validLineCount = 0 Then
        MsgBox "No se encontraron líneas válidas de código NASM"
        Exit Sub
    End If
    
    ReDim Program(0 To validLineCount - 1)
    ParseNASMLines ws, lastRow, validLineCount
    
    CurrentInstruction = 0
    DataSectionStart = &H0
    TextSectionStart = &H80
End Sub

Function CountValidNASMLines(ws As Worksheet, lastRow As Integer) As Integer
    Dim i As Integer, count As Integer, line As String
    count = 0
    For i = 2 To lastRow
        line = Trim(ws.Cells(i, 1).value)
        If IsValidNASMLine(line) Then count = count + 1
    Next i
    CountValidNASMLines = count
End Function

Function IsValidNASMLine(line As String) As Boolean
    If line = "" Or Left(Trim(line), 1) = ";" Then Exit Function
    Dim cleanLine As String
    cleanLine = LCase(Trim(line))
    
    IsValidNASMLine = (cleanLine = "section .data" Or cleanLine = "section .text" Or _
                       cleanLine = "global _start" Or Right(cleanLine, 1) = ":" Or _
                       InStr(cleanLine, "db ") > 0 Or InStr(cleanLine, "dw ") > 0 Or _
                       InStr(cleanLine, "dd ") > 0 Or InStr(cleanLine, "equ ") > 0 Or _
                       InStr(cleanLine, "mov ") > 0 Or InStr(cleanLine, "add ") > 0 Or _
                       InStr(cleanLine, "sub ") > 0 Or InStr(cleanLine, "int ") > 0 Or _
                       InStr(cleanLine, "xor ") > 0 Or cleanLine = "nop")
End Function

Sub ParseNASMLines(ws As Worksheet, lastRow As Integer, validLineCount As Integer)
    Dim i As Integer, programIndex As Integer
    Dim currentSection As String, currentAddress As Integer, line As String
    
    programIndex = 0
    currentSection = ".data"
    currentAddress = DataSectionStart
    
    For i = 2 To lastRow
        line = Trim(ws.Cells(i, 1).value)
        
        If IsValidNASMLine(line) Then
            With Program(programIndex)
                .OriginalLine = line
                
                If LCase(Trim(line)) = "section .data" Then
                    currentSection = ".data"
                    currentAddress = DataSectionStart
                ElseIf LCase(Trim(line)) = "section .text" Then
                    currentSection = ".text"
                    currentAddress = TextSectionStart
                End If
                
                .section = currentSection
                .address = currentAddress
                
                If LCase(Trim(line)) = "section .data" Or LCase(Trim(line)) = "section .text" Then
                    .opcode = "section"
                    .Operand1 = currentSection
                ElseIf LCase(Trim(line)) = "global _start" Then
                    .opcode = "global"
                    .Operand1 = "_start"
                ElseIf Right(line, 1) = ":" Then
                    .opcode = "label"
                    .Operand1 = Replace(line, ":", "")
                Else
                    ParseNASMInstruction line, programIndex
                    .Length = GetVisualLength(Program(programIndex))
                    currentAddress = currentAddress + .Length
                End If
            End With
            programIndex = programIndex + 1
        End If
    Next i
End Sub

Function GetVisualLength(inst As AssemblyInstruction) As Integer
    Dim Length As Integer
    Length = 0
    If inst.section = ".text" Then
        If inst.opcode <> "" Then Length = Length + 1
        If inst.Operand1 <> "" Then Length = Length + 1
        If inst.Operand2 <> "" Then Length = Length + 1
    Else
        Length = Len(Replace(inst.bytes, " ", "")) / 2
    End If
    If Length = 0 And inst.opcode <> "label" And inst.opcode <> "global" And inst.opcode <> "section" Then Length = 1
    GetVisualLength = Length
End Function

Sub ParseNASMInstruction(line As String, index As Integer)
    If InStr(line, ";") > 0 Then line = Trim(Left(line, InStr(line, ";") - 1))
    
    If InStr(line, "db ") > 0 Or InStr(line, "dw ") > 0 Or InStr(line, "dd ") > 0 Then
        ParseDataDefinition line, index
    Else
        ParseCPUInstruction line, index
    End If
End Sub

Sub ParseDataDefinition(line As String, index As Integer)
    Dim parts() As String, varName As String, dataType As String, value As String, i As Integer
    parts = Split(line, " ")
    varName = parts(0)
    dataType = parts(1)
    
    value = ""
    For i = 2 To UBound(parts)
        value = value & parts(i) & " "
    Next i
    value = Trim(value)
    
    With Program(index)
        .opcode = dataType
        .Operand1 = varName
        .Operand2 = value
        
        Select Case LCase(dataType)
            Case "db"
                .bytes = IIf(Left(value, 1) = """", StringToHex(Replace(value, """", "")), Format(Hex(Val(value)), "00"))
            Case "dw"
                .bytes = GetWordBytes(value)
            Case "dd"
                .bytes = GetDoubleWordBytes(value)
            Case Else
                .bytes = "00"
        End Select
        .Length = Len(Replace(.bytes, " ", "")) \ 2
    End With
End Sub

Sub ParseCPUInstruction(line As String, index As Integer)
    Dim parts() As String, mainParts() As String
    mainParts = Split(line, ",")
    parts = Split(Trim(mainParts(0)), " ")
    
    With Program(index)
        .opcode = parts(0)
        If UBound(parts) >= 1 Then .Operand1 = Trim(parts(1))
        If UBound(mainParts) >= 1 Then .Operand2 = Trim(mainParts(1))
        .bytes = "SIM"
    End With
End Sub

Function StringToHex(s As String) As String
    Dim i As Integer, result As String
    For i = 1 To Len(s)
        result = result & Format(Hex(Asc(Mid(s, i, 1))), "00") & " "
    Next i
    StringToHex = Trim(result)
End Function

Function GetWordBytes(value As String) As String
    Dim num As Integer
    If IsNumeric(value) Then
        num = Val(value)
        GetWordBytes = Format(Hex(num Mod 256), "00") & " " & Format(Hex(num \ 256), "00")
    Else
        GetWordBytes = "00 00"
    End If
End Function

Function GetDoubleWordBytes(value As String) As String
    Dim num As Long
    If IsNumeric(value) Then
        num = Val(value)
        GetDoubleWordBytes = Format(Hex(num Mod 256), "00") & " " & _
                             Format(Hex((num \ 256) Mod 256), "00") & " " & _
                             Format(Hex((num \ 65536) Mod 256), "00") & " " & _
                             Format(Hex(num \ 16777216), "00")
    Else
        GetDoubleWordBytes = "00 00 00 00"
    End If
End Function

' ============= NUEVA FUNCIÓN: CARGA DINÁMICA DE INSTRUCCIÓN =============
Sub LoadCurrentInstructionToRAM()
    If CurrentInstruction > UBound(Program) Then Exit Sub
    
    Dim instruction As AssemblyInstruction
    instruction = Program(CurrentInstruction)
    
    Dim i As Integer, j As Integer
    Dim bytes() As String
    Dim byteValue As Byte
    Dim currentAddr As Integer
    
    currentAddr = instruction.address
    
    If instruction.section = ".text" Then
        ' Cargar mnemónicos/operandos en RAM_Display
        If instruction.opcode <> "label" And instruction.opcode <> "global" And instruction.opcode <> "section" Then
            If currentAddr < RAM_SIZE And Not RAM_Loaded(currentAddr) Then
                RAM_Display(currentAddr) = instruction.opcode
                RAM_Loaded(currentAddr) = True
            End If
            
            If instruction.Operand1 <> "" Then
                If currentAddr + 1 < RAM_SIZE And Not RAM_Loaded(currentAddr + 1) Then
                    RAM_Display(currentAddr + 1) = instruction.Operand1
                    RAM_Loaded(currentAddr + 1) = True
                End If
            End If
            
            If instruction.Operand2 <> "" Then
                If currentAddr + 2 < RAM_SIZE And Not RAM_Loaded(currentAddr + 2) Then
                    RAM_Display(currentAddr + 2) = instruction.Operand2
                    RAM_Loaded(currentAddr + 2) = True
                End If
            End If
        End If
        
    ElseIf instruction.section = ".data" Then
        ' Cargar bytes de datos
        If instruction.bytes <> "" Then
            bytes = Split(instruction.bytes, " ")
            For j = 0 To UBound(bytes)
                If bytes(j) <> "" And currentAddr + j < RAM_SIZE And Not RAM_Loaded(currentAddr + j) Then
                    byteValue = CInt("&H" & bytes(j))
                    RAM(currentAddr + j) = byteValue
                    RAM_Display(currentAddr + j) = Format(Hex(byteValue), "00")
                    RAM_Loaded(currentAddr + j) = True
                End If
            Next j
        End If
    End If
End Sub

' ============= INTERFAZ GRÁFICA =============
Sub DrawRAMGridNASM()
    Dim i As Integer, j As Integer, ws As Worksheet
    
    On Error Resume Next
    Set ws = Worksheets("RAM")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "RAM"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    ' Encabezados
    With ws
        .Range("B1").value = "SIMULADOR DE MEMORIA RAM - NASM (CARGA DINÁMICA)"
        .Range("B1:J1").Merge
        .Range("B1").Font.Bold = True
        .Range("B1").Font.size = 14
        
        .Range("A3").value = "Dir"
        .Range("A3").Font.Bold = True
        
        For i = 0 To COLS - 1
            .Cells(3, i + 2).value = Hex(i)
            .Cells(3, i + 2).Font.Bold = True
            .Cells(3, i + 2).HorizontalAlignment = xlCenter
            .Cells(3, i + 2).Interior.color = RGB(200, 200, 200)
        Next i
        
        ' Crear celdas de RAM
        For i = 0 To ROWS - 1
            .Cells(4 + i, 1).value = "0x" & Format(Hex(i * COLS), "00")
            .Cells(4 + i, 1).Font.Bold = True
            .Cells(4 + i, 1).Interior.color = RGB(200, 200, 200)
            
            For j = 0 To COLS - 1
                With .Cells(4 + i, j + 2)
                    .value = "00"
                    .HorizontalAlignment = xlCenter
                    .Borders.LineStyle = xlContinuous
                    .Interior.color = RGB(240, 240, 240)
                    .Font.Name = "Courier New"
                    .Font.size = 9
                End With
            Next j
        Next i
        
        .Columns("A").ColumnWidth = 6
        .Columns("B:Q").ColumnWidth = 4.5
        
        ' Panel de programa
        .Range("S3").value = "PROGRAMA"
        .Range("S3:U3").Merge
        .Range("S3").Font.Bold = True
        .Range("S4").value = "Addr"
        .Range("T4").value = "Sec"
        .Range("U4").value = "Código"
        .Range("S4:U4").Font.Bold = True
        
        ' Panel de estado
        .Range("S18").value = "ESTADO"
        .Range("S18:U18").Merge
        .Range("S18").Font.Bold = True
        .Range("S19").value = "Instr:"
        .Range("S20").value = "Dir:"
        .Range("S21").value = "Sec:"
        .Range("S22").value = "Estado:"
    End With
    
    CreateControlButtonsNASM
End Sub

Sub DisplayNASMProgramInfo()
    Dim i As Integer, ws As Worksheet
    Set ws = Worksheets("RAM")
    
    ws.Range("S5:U17").ClearContents
    
    For i = 0 To UBound(Program)
        If 5 + i <= 17 Then
            ws.Cells(5 + i, 19).value = "0x" & Format(Hex(Program(i).address), "00")
            ws.Cells(5 + i, 20).value = Left(Program(i).section, 5)
            ws.Cells(5 + i, 21).value = Left(Program(i).OriginalLine, 25)
        End If
    Next i
End Sub

' OPTIMIZADA: Actualización más rápida
Sub UpdateRAMDisplayNASM_Improved()
    Dim i As Integer, j As Integer, addr As Integer
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    
    ' Actualizar solo el rango necesario de una vez
    Dim displayData() As Variant
    ReDim displayData(1 To ROWS, 1 To COLS)
    
    For i = 0 To ROWS - 1
        For j = 0 To COLS - 1
            addr = i * COLS + j
            displayData(i + 1, j + 1) = RAM_Display(addr)
        Next j
    Next i
    
    ' Una sola escritura al rango completo (MUY RÁPIDO)
    ws.Range("B4").Resize(ROWS, COLS).value = displayData
    
    ' Aplicar colores por sección
    For i = 0 To ROWS - 1
        For j = 0 To COLS - 1
            addr = i * COLS + j
            With ws.Cells(4 + i, j + 2)
                If addr >= TextSectionStart Then
                    .Interior.color = RGB(200, 255, 200)
                Else
                    .Interior.color = RGB(200, 220, 255)
                End If
            End With
        Next j
    Next i
End Sub

' ============= EJECUCIÓN =============
Sub ExecuteNextInstructionNASM()
    If CurrentInstruction > UBound(Program) Then
        MsgBox "Programa completado"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False ' OPTIMIZACIÓN
    
    ' Saltar directivas
    While Program(CurrentInstruction).section <> ".text" Or _
          Program(CurrentInstruction).opcode = "label" Or _
          Program(CurrentInstruction).opcode = "global"
        CurrentInstruction = CurrentInstruction + 1
        If CurrentInstruction > UBound(Program) Then
            UpdateNASMStatus "COMPLETADO", "---", "---", "Fin del programa"
            Application.ScreenUpdating = True
            Exit Sub
        End If
    Wend
    
    ' NUEVA: Cargar instrucción actual en RAM dinámicamente
    LoadCurrentInstructionToRAM
    
    UpdateRAMDisplayNASM_Improved
    HighlightCurrentInstructionNASM
    SimulateNASMMemoryAccess
    
    CurrentInstruction = CurrentInstruction + 1
    
    Application.ScreenUpdating = True ' OPTIMIZACIÓN
End Sub

Sub HighlightCurrentInstructionNASM()
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    
    ws.Range("S5:U17").Interior.color = RGB(255, 255, 255)
    
    If CurrentInstruction <= UBound(Program) Then
        Dim displayRow As Integer
        displayRow = 5 + CurrentInstruction
        If displayRow <= 17 Then
            ws.Range("S" & displayRow & ":U" & displayRow).Interior.color = RGB(255, 255, 0)
            ws.Range("S" & displayRow & ":U" & displayRow).Font.Bold = True
        End If
    End If
End Sub

Sub SimulateNASMMemoryAccess()
    If CurrentInstruction > UBound(Program) Then Exit Sub
    
    Dim instruction As AssemblyInstruction
    instruction = Program(CurrentInstruction)
    
    Dim i As Integer, addr As Integer
    For i = 0 To instruction.Length - 1
        addr = instruction.address + i
        If addr < RAM_SIZE Then HighlightMemoryCellNASM addr, RGB(255, 165, 0)
    Next i
    
    UpdateNASMStatus instruction.OriginalLine, "0x" & Format(Hex(instruction.address), "00"), _
                      instruction.section, "Ejecutando"
End Sub

Sub HighlightMemoryCellNASM(addr As Integer, color As Long)
    Dim row As Integer, col As Integer, ws As Worksheet
    Set ws = Worksheets("RAM")
    
    row = 4 + (addr \ COLS)
    col = 2 + (addr Mod COLS)
    
    With ws.Cells(row, col)
        .Interior.color = color
        .Font.Bold = True
    End With
End Sub

Sub UpdateNASMStatus(instruction As String, address As String, section As String, accessType As String)
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    
    ws.Range("T19").value = Left(instruction, 40)
    ws.Range("T20").value = address
    ws.Range("T21").value = section
    ws.Range("T22").value = accessType
End Sub

' ============= CONTROLES =============
Sub CreateControlButtonsNASM()
    Dim btn As Button, ws As Worksheet
    Set ws = Worksheets("RAM")
    
    On Error Resume Next: ws.Buttons.Delete: On Error GoTo 0
    
    Set btn = ws.Buttons.Add(300, 370, 100, 25)
    btn.OnAction = "ExecuteNextInstructionNASM"
    btn.Characters.Text = "Siguiente"
    
    Set btn = ws.Buttons.Add(410, 370, 100, 25)
    btn.OnAction = "ExecuteFullProgramNASM"
    btn.Characters.Text = "Ejecutar Todo"
    
    Set btn = ws.Buttons.Add(520, 370, 80, 25)
    btn.OnAction = "ResetSimulatorNASM"
    btn.Characters.Text = "Reiniciar"
End Sub

Sub ExecuteFullProgramNASM()
    Application.ScreenUpdating = False ' OPTIMIZACIÓN
    
    Do While CurrentInstruction <= UBound(Program)
        ExecuteNextInstructionNASM
        If CurrentInstruction > UBound(Program) Then Exit Do
    Loop
    
    Application.ScreenUpdating = True
    MsgBox "Ejecución completada"
End Sub

Sub ResetSimulatorNASM()
    CurrentInstruction = 0
    InitializeRAMSimulatorNASM
End Sub

