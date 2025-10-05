Attribute VB_Name = "RAM_Simulator"
' Módulo: RAM_Simulator_NASM
' Descripción: Simulador de memoria RAM que lee programas NASM desde celdas Excel
' Versión: 2.0 (Visualización Mejorada)

Option Explicit

' Constantes para la configuración
Private Const RAM_SIZE As Integer = 256
Private Const ROWS As Integer = 16
Private Const COLS As Integer = 16

' Variables globales
Private RAM(0 To RAM_SIZE - 1) As Byte ' Para la lógica interna y datos
Private RAM_Display(0 To RAM_SIZE - 1) As String ' --- NUEVO: Para la visualización en la cuadrícula
Private Program() As AssemblyInstruction
Private DataSectionStart As Integer
Private TextSectionStart As Integer
Private CurrentInstruction As Integer

' Tipo para instrucciones
Type AssemblyInstruction
    address As Integer
    OriginalLine As String
    opcode As String
    Operand1 As String
    Operand2 As String
    Operand3 As String
    bytes As String
    Length As Integer
    section As String ' .data, .text, etc.
End Type

' Inicializar el simulador para NASM
Sub InitializeRAMSimulatorNASM()
    ' Limpiar RAM
    ClearRAM
    
    ' Leer programa NASM desde celdas
    ReadNASMProgramFromCells
    
    ' Dibujar interfaz
    DrawRAMGridNASM
    
    ' Cargar programa en RAM (con la nueva lógica de visualización)
    LoadNASMProgramIntoRAM_Improved ' --- MODIFICADO: Llama a la nueva función de carga
    
    ' Actualizar visualización
    UpdateRAMDisplayNASM_Improved ' --- MODIFICADO: Llama a la nueva función de actualización
    
    UpdateNASMStatus "Listo", "---", "---", "Programa cargado"
    MsgBox "Simulador de RAM para NASM inicializado. Programa cargado desde celdas."
End Sub

' Limpiar la RAM y el array de visualización
Sub ClearRAM() ' --- MODIFICADO ---
    Dim i As Integer
    For i = 0 To RAM_SIZE - 1
        RAM(i) = 0
        RAM_Display(i) = "00" ' Inicializa la visualización con "00"
    Next i
End Sub

' Leer programa NASM desde celdas de Excel (Sin cambios)
Sub ReadNASMProgramFromCells()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ProgramaNASM")
    
    Dim lastRow As Integer
    lastRow = ws.Cells(ws.ROWS.count, "A").End(xlUp).row
    
    If lastRow < 2 Then
        MsgBox "No se encontró programa en las celdas. Por favor, ingrese el programa en la hoja 'ProgramaNASM'"
        Exit Sub
    End If
    
    ' Primera pasada: contar líneas válidas
    Dim validLineCount As Integer
    validLineCount = CountValidNASMLines(ws, lastRow)
    
    If validLineCount = 0 Then
        MsgBox "No se encontraron líneas válidas de código NASM"
        Exit Sub
    End If
    
    ReDim Program(0 To validLineCount - 1)
    
    ' Segunda pasada: procesar líneas
    ParseNASMLines ws, lastRow, validLineCount
    
    CurrentInstruction = 0
    DataSectionStart = &H0   ' Sección .data empieza en 0x00
    TextSectionStart = &H80 ' Sección .text empieza en 0x80
End Sub

' --- INICIO DE SECCIÓN DE CÓDIGO SIN CAMBIOS (Parseo) ---
' Estas funciones de parseo no necesitan cambios ya que la lógica
' de interpretación del código NASM es la misma.

' Contar líneas válidas de código NASM
Function CountValidNASMLines(ws As Worksheet, lastRow As Integer) As Integer
    Dim i As Integer
    Dim count As Integer
    Dim line As String
    
    count = 0
    For i = 2 To lastRow
        line = Trim(ws.Cells(i, 1).value)
        If IsValidNASMLine(line) Then
            count = count + 1
        End If
    Next i
    
    CountValidNASMLines = count
End Function

' Verificar si una línea es válida de código NASM
Function IsValidNASMLine(line As String) As Boolean
    If line = "" Then Exit Function
    If Left(Trim(line), 1) = ";" Then Exit Function ' Comentarios
    
    Dim cleanLine As String
    cleanLine = LCase(Trim(line))
    
    ' Secciones y directivas
    If cleanLine = "section .data" Then IsValidNASMLine = True
    If cleanLine = "section .text" Then IsValidNASMLine = True
    If cleanLine = "global _start" Then IsValidNASMLine = True
    If Right(cleanLine, 1) = ":" Then IsValidNASMLine = True ' Etiquetas
    
    ' Directivas de datos
    If InStr(cleanLine, "db ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "dw ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "dd ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "equ ") > 0 Then IsValidNASMLine = True
    
    ' Instrucciones de CPU
    If InStr(cleanLine, "mov ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "add ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "sub ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "cmp ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "inc ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "dec ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "and ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "or ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "xor ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "jmp ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "je ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "jne ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "jl ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "jle ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "jg ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "jge ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "call ") > 0 Then IsValidNASMLine = True
    If cleanLine = "ret" Then IsValidNASMLine = True
    If InStr(cleanLine, "push ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "pop ") > 0 Then IsValidNASMLine = True
    If InStr(cleanLine, "int ") > 0 Then IsValidNASMLine = True
    If cleanLine = "nop" Then IsValidNASMLine = True
    If cleanLine = "hlt" Then IsValidNASMLine = True
End Function

' Parsear líneas NASM
Sub ParseNASMLines(ws As Worksheet, lastRow As Integer, validLineCount As Integer)
    Dim i As Integer
    Dim programIndex As Integer
    Dim currentSection As String
    Dim currentAddress As Integer
    Dim line As String
    
    programIndex = 0
    currentSection = ".data" ' Empezar en .data por defecto
    currentAddress = DataSectionStart
    
    For i = 2 To lastRow
        line = Trim(ws.Cells(i, 1).value)
        
        If IsValidNASMLine(line) Then
            With Program(programIndex)
                .OriginalLine = line
                
                ' Actualizar sección si la línea es una directiva de sección
                If LCase(Trim(line)) = "section .data" Then
                    currentSection = ".data"
                    currentAddress = DataSectionStart
                ElseIf LCase(Trim(line)) = "section .text" Then
                    currentSection = ".text"
                    currentAddress = TextSectionStart
                End If
                
                .section = currentSection
                
                ' Determinar tipo de línea y asignar dirección
                If LCase(Trim(line)) = "section .data" Or LCase(Trim(line)) = "section .text" Then
                    .address = currentAddress
                    .opcode = "section"
                    .Operand1 = currentSection
                ElseIf LCase(Trim(line)) = "global _start" Then
                    .address = currentAddress
                    .opcode = "global"
                    .Operand1 = "_start"
                ElseIf Right(line, 1) = ":" Then
                    .address = currentAddress
                    .opcode = "label"
                    .Operand1 = Replace(line, ":", "")
                Else
                    .address = currentAddress
                    ParseNASMInstruction line, programIndex
                    ' El avance de la dirección se basa ahora en la longitud visual
                    .Length = GetVisualLength(Program(programIndex))
                    currentAddress = currentAddress + .Length
                End If
            End With
            
            programIndex = programIndex + 1
        End If
    Next i
End Sub

' --- NUEVA FUNCIÓN: Calcular la longitud visual de una instrucción ---
Function GetVisualLength(inst As AssemblyInstruction) As Integer
    Dim Length As Integer
    Length = 0
    If inst.section = ".text" Then
        If inst.opcode <> "" Then Length = Length + 1
        If inst.Operand1 <> "" Then Length = Length + 1
        If inst.Operand2 <> "" Then Length = Length + 1
        If inst.Operand3 <> "" Then Length = Length + 1
    Else ' Para .data, la longitud es la de los bytes
        Length = Len(Replace(inst.bytes, " ", "")) / 2
    End If
    
    If Length = 0 And inst.opcode <> "label" And inst.opcode <> "global" And inst.opcode <> "section" Then Length = 1
    GetVisualLength = Length
End Function

' Parsear instrucción NASM individual
Sub ParseNASMInstruction(line As String, index As Integer)
    If InStr(line, ";") > 0 Then
        line = Trim(Left(line, InStr(line, ";") - 1))
    End If
    
    If InStr(line, "db ") > 0 Or InStr(line, "dw ") > 0 Or InStr(line, "dd ") > 0 Then
        ParseDataDefinition line, index
    ElseIf InStr(line, "equ ") > 0 Then
        ParseEquDefinition line, index
    Else
        ParseCPUInstruction line, index
    End If
End Sub

' Parsear definición de datos (db, dw, dd)
Sub ParseDataDefinition(line As String, index As Integer)
    Dim parts() As String
    Dim varName As String
    Dim dataType As String
    Dim value As String
    Dim i As Integer
    
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
                If Left(value, 1) = """" Then ' String
                    .bytes = StringToHex(Replace(value, """", ""))
                Else ' Valor numérico
                    .bytes = Format(Hex(Val(value)), "00")
                End If
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

' Parsear definición EQU
Sub ParseEquDefinition(line As String, index As Integer)
    Dim parts() As String
    parts = Split(line, " ")
    
    With Program(index)
        .opcode = "equ"
        .Operand1 = parts(0)
        .Operand2 = parts(2)
        .bytes = ""
        .Length = 0
    End With
End Sub

' Parsear instrucción de CPU
Sub ParseCPUInstruction(line As String, index As Integer)
    Dim parts() As String
    Dim mainParts() As String
    
    mainParts = Split(line, ",")
    parts = Split(Trim(mainParts(0)), " ")
    
    With Program(index)
        .opcode = parts(0)
        
        If UBound(parts) >= 1 Then
            .Operand1 = Trim(parts(1))
        End If
        
        If UBound(mainParts) >= 1 Then
            .Operand2 = Trim(mainParts(1))
        End If
        
        ' La generación de bytes ya no es necesaria para la visualización,
        ' pero se puede mantener para una simulación más profunda.
        .bytes = "SIM" ' Marcador para instrucción simulada
    End With
End Sub

' Convertir string a hexadecimal
Function StringToHex(s As String) As String
    Dim i As Integer
    Dim result As String
    
    result = ""
    For i = 1 To Len(s)
        result = result & Format(Hex(Asc(Mid(s, i, 1))), "00") & " "
    Next i
    
    StringToHex = Trim(result)
End Function

' Función para bytes de palabra (16 bits)
Function GetWordBytes(value As String) As String
    Dim num As Integer
    If IsNumeric(value) Then
        num = Val(value)
        GetWordBytes = Format(Hex(num Mod 256), "00") & " " & Format(Hex(num \ 256), "00")
    Else
        GetWordBytes = "00 00"
    End If
End Function

' Función para bytes de doble palabra (32 bits)
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

' --- FIN DE SECCIÓN DE CÓDIGO SIN CAMBIOS ---


' Dibujar la cuadrícula de RAM para NASM (Sin cambios)
Sub DrawRAMGridNASM()
    Dim i As Integer, j As Integer
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Worksheets("RAM")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "RAM"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    With ws
        .Range("B1").value = "SIMULADOR DE MEMORIA RAM - NASM"
        .Range("B1").Font.Bold = True
        .Range("B1").Font.size = 16
        .Range("B1:J1").Merge
        
        .Range("A3").value = "Dirección"
        .Range("A3").Font.Bold = True
        
        For i = 0 To COLS - 1
            .Cells(3, i + 2).value = Hex(i)
            .Cells(3, i + 2).Font.Bold = True
            .Cells(3, i + 2).HorizontalAlignment = xlCenter
            .Cells(3, i + 2).Interior.color = RGB(200, 200, 200)
        Next i
        
        Dim row As Integer
        row = 4
        For i = 0 To ROWS - 1
            Dim addr As Integer
            addr = i * COLS
            
            .Cells(row + i, 1).value = "0x" & Format(Hex(addr), "000")
            .Cells(row + i, 1).Font.Bold = True
            .Cells(row + i, 1).Interior.color = RGB(200, 200, 200)
            
            For j = 0 To COLS - 1
                With .Cells(row + i, j + 2)
                    .value = "00"
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Borders.LineStyle = xlContinuous
                    .Interior.color = RGB(240, 240, 240)
                    .Font.Name = "Courier New"
                End With
            Next j
        Next i
        
        .Columns("A").ColumnWidth = 8
        .Columns("B:Q").ColumnWidth = 5 ' Un poco más ancho para el texto
        
        .Range("S3").value = "PROGRAMA NASM CARGADO"
        .Range("S3").Font.Bold = True
        .Range("S3:U3").Merge
        .Range("S3:U3").HorizontalAlignment = xlCenter
        .Range("S3:U3").Interior.color = RGB(200, 200, 200)
        
        .Range("S4").value = "Addr"
        .Range("T4").value = "Sección"
        .Range("U4").value = "Código"
        .Range("S4:U4").Font.Bold = True
        .Range("S4:U4").Interior.color = RGB(220, 220, 220)
        
        .Range("S18").value = "ESTADO DE EJECUCIÓN"
        .Range("S18:U18").Merge
        .Range("S18:U18").Font.Bold = True
        .Range("S18:U18").HorizontalAlignment = xlCenter
        .Range("S18:U18").Interior.color = RGB(200, 200, 200)
        
        .Range("S19").value = "Instrucción:"
        .Range("T19:U19").Merge
        .Range("S20").value = "Dirección:"
        .Range("T20:U20").Merge
        .Range("S21").value = "Sección:"
        .Range("T21:U21").Merge
        .Range("S22").value = "Acceso:"
        .Range("T22:U22").Merge
    End With
    
    CreateControlButtonsNASM
End Sub

' --- NUEVA FUNCIÓN MEJORADA para cargar el programa en la RAM ---
Sub LoadNASMProgramIntoRAM_Improved()
    Dim i As Integer, j As Integer
    Dim bytes() As String
    Dim byteValue As Byte
    Dim currentAddr As Integer

    For i = 0 To UBound(Program)
        currentAddr = Program(i).address
        
        If Program(i).section = ".text" Then
            ' Para la sección de código, colocamos los mnemónicos/operandos en RAM_Display
            If Program(i).opcode <> "label" And Program(i).opcode <> "global" And Program(i).opcode <> "section" Then
                If currentAddr < RAM_SIZE Then RAM_Display(currentAddr) = Program(i).opcode
                
                If Program(i).Operand1 <> "" Then
                    If currentAddr + 1 < RAM_SIZE Then RAM_Display(currentAddr + 1) = Program(i).Operand1
                End If
                If Program(i).Operand2 <> "" Then
                    If currentAddr + 2 < RAM_SIZE Then RAM_Display(currentAddr + 2) = Program(i).Operand2
                End If
                ' Nota: Se asume que las instrucciones tienen máx 3 componentes (opcode, op1, op2)
            End If

        ElseIf Program(i).section = ".data" Then
            ' Para la sección de datos, cargamos los bytes hexadecimales como antes
            If Program(i).bytes <> "" Then
                bytes = Split(Program(i).bytes, " ")
                
                For j = 0 To UBound(bytes)
                    If bytes(j) <> "" Then
                        byteValue = CInt("&H" & bytes(j))
                        If currentAddr + j < RAM_SIZE Then
                            RAM(currentAddr + j) = byteValue
                            RAM_Display(currentAddr + j) = Format(Hex(byteValue), "00")
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    
    DisplayNASMProgramInfo
End Sub


' Mostrar información del programa NASM (Sin cambios)
Sub DisplayNASMProgramInfo()
    Dim i As Integer
    Dim startRow As Integer
    Dim ws As Worksheet
    
    Set ws = Worksheets("RAM")
    startRow = 5
    
    ws.Range("S5:U17").ClearContents
    ws.Range("S5:U17").Interior.color = RGB(255, 255, 255)
    
    For i = 0 To UBound(Program)
        If startRow + i <= 17 Then
            ws.Cells(startRow + i, 19).value = "0x" & Format(Hex(Program(i).address), "000")
            ws.Cells(startRow + i, 20).value = Program(i).section
            ws.Cells(startRow + i, 21).value = Program(i).OriginalLine
        End If
    Next i
End Sub

' --- NUEVA FUNCIÓN MEJORADA para actualizar la visualización de la RAM ---
Sub UpdateRAMDisplayNASM_Improved()
    Dim i As Integer, j As Integer
    Dim addr As Integer
    Dim displayRow As Integer
    Dim ws As Worksheet
    
    Set ws = Worksheets("RAM")
    displayRow = 4
    
    For i = 0 To ROWS - 1
        For j = 0 To COLS - 1
            addr = i * COLS + j
            
            ' Usamos el nuevo array RAM_Display para poblar las celdas
            ws.Cells(displayRow + i, j + 2).value = RAM_Display(addr)
            
            ' Color por sección
            If addr >= TextSectionStart And addr < RAM_SIZE Then
                ' Sección .text - color verde claro
                ws.Cells(displayRow + i, j + 2).Interior.color = RGB(200, 255, 200)
            Else
                ' Sección .data - color azul claro
                ws.Cells(displayRow + i, j + 2).Interior.color = RGB(200, 220, 255)
            End If
            
            ws.Cells(displayRow + i, j + 2).Font.color = RGB(0, 0, 0)
            ws.Cells(displayRow + i, j + 2).Font.Bold = False
        Next j
    Next i
End Sub

' Ejecutar siguiente instrucción NASM (Sin cambios en la lógica principal)
Sub ExecuteNextInstructionNASM()
    If CurrentInstruction > UBound(Program) Then
        MsgBox "Programa completado"
        Exit Sub
    End If
    
    ' Saltar directivas que no se ejecutan
    While Program(CurrentInstruction).section <> ".text" Or Program(CurrentInstruction).opcode = "label" Or Program(CurrentInstruction).opcode = "global"
        CurrentInstruction = CurrentInstruction + 1
        If CurrentInstruction > UBound(Program) Then
            UpdateNASMStatus "COMPLETADO", "---", "---", "Programa finalizado"
            MsgBox "Fin del programa."
            Exit Sub
        End If
    Wend
    
    UpdateRAMDisplayNASM_Improved ' Refrescar por si el resaltado anterior debe limpiarse
    HighlightCurrentInstructionNASM
    SimulateNASMMemoryAccess
    
    CurrentInstruction = CurrentInstruction + 1
    
    If CurrentInstruction > UBound(Program) Then
        UpdateNASMStatus "COMPLETADO", "---", "---", "Programa finalizado"
    End If
End Sub

' Resaltar instrucción actual NASM (Sin cambios)
Sub HighlightCurrentInstructionNASM()
    Dim i As Integer
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    
    For i = 5 To 17
        ws.Range("S" & i & ":U" & i).Interior.color = RGB(255, 255, 255)
        ws.Range("S" & i & ":U" & i).Font.Bold = False
    Next i
    
    If CurrentInstruction <= UBound(Program) Then
        Dim displayIndex As Integer
        displayIndex = FindDisplayIndexForInstruction(CurrentInstruction)
        
        If 5 + displayIndex <= 17 Then
            ws.Range("S" & 5 + displayIndex & ":U" & 5 + displayIndex).Interior.color = RGB(255, 255, 0) ' Amarillo
            ws.Range("S" & 5 + displayIndex & ":U" & 5 + displayIndex).Font.Bold = True
        End If
    End If
End Sub

Function FindDisplayIndexForInstruction(programIdx As Integer) As Integer
    Dim i As Integer
    Dim count As Integer
    count = -1
    For i = 0 To programIdx
         count = count + 1
    Next i
    FindDisplayIndexForInstruction = count
End Function


' Simular acceso a memoria para NASM
Sub SimulateNASMMemoryAccess()
    Dim addr As Integer
    Dim i As Integer
    Dim instruction As AssemblyInstruction
    
    If CurrentInstruction > UBound(Program) Then Exit Sub
    
    instruction = Program(CurrentInstruction)
    
    ' Resaltar los bytes/componentes de la instrucción
    If instruction.Length > 0 Then
        For i = 0 To instruction.Length - 1
            addr = instruction.address + i
            If addr < RAM_SIZE Then
                HighlightMemoryCellNASM addr, RGB(255, 165, 0), "Ejecución" ' Naranja para la instrucción actual
            End If
        Next i
    End If
    
    ' Actualizar panel de estado (simplificado)
    UpdateNASMStatus instruction.OriginalLine, "0x" & Format(Hex(instruction.address), "000"), instruction.section, "Ejecutando"
End Sub

' Resaltar celda de memoria NASM (Sin cambios)
Sub HighlightMemoryCellNASM(addr As Integer, color As Long, accessType As String)
    Dim row As Integer, col As Integer
    Dim displayRow As Integer
    Dim ws As Worksheet
    
    Set ws = Worksheets("RAM")
    displayRow = 4
    
    row = displayRow + (addr \ COLS)
    col = 2 + (addr Mod COLS)
    
    If row >= displayRow And row < displayRow + ROWS And col >= 2 And col < 2 + COLS Then
        With ws.Cells(row, col)
            .Interior.color = color
            .Font.Bold = True
        End With
    End If
End Sub

' Actualizar panel de estado NASM (Sin cambios)
Sub UpdateNASMStatus(instruction As String, address As String, section As String, accessType As String)
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    
    ws.Range("T19").value = instruction
    ws.Range("T20").value = address
    ws.Range("T21").value = section
    ws.Range("T22").value = accessType
End Sub

' Crear botones de control NASM (Sin cambios)
Sub CreateControlButtonsNASM()
    Dim btn As Button
    Dim ws As Worksheet
    
    Set ws = Worksheets("RAM")
    
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    Set btn = ws.Buttons.Add(300, 400, 100, 25)
    btn.OnAction = "ExecuteNextInstructionNASM"
    btn.Characters.Text = "Siguiente"
    btn.Name = "BtnNextNASM"
    
    Set btn = ws.Buttons.Add(410, 400, 100, 25)
    btn.OnAction = "ExecuteFullProgramNASM"
    btn.Characters.Text = "Ejecutar Todo"
    btn.Name = "BtnRunAllNASM"
    
    Set btn = ws.Buttons.Add(520, 400, 80, 25)
    btn.OnAction = "ResetSimulatorNASM"
    btn.Characters.Text = "Reiniciar"
    btn.Name = "BtnResetNASM"
End Sub

' Ejecutar programa completo NASM (Sin cambios)
Sub ExecuteFullProgramNASM()
    Do While CurrentInstruction <= UBound(Program)
        ExecuteNextInstructionNASM
        If CurrentInstruction > UBound(Program) Then Exit Do
        DoEvents
        Wait 0.2 ' Reducido para una ejecución más rápida
    Loop
End Sub

' Reiniciar simulador NASM
Sub ResetSimulatorNASM()
    CurrentInstruction = 0
    InitializeRAMSimulatorNASM
End Sub

' Función de espera (Sin cambios)
Sub Wait(seconds As Double)
    Dim startTime As Double
    startTime = Timer
    Do While Timer < startTime + seconds
        DoEvents
    Loop
End Sub

' Crear hoja de programa NASM de ejemplo (Sin cambios)
Sub CreateSampleNASMProgramSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Worksheets("ProgramaNASM")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "ProgramaNASM"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    ws.Range("A1").value = "Código NASM (Ingrese su programa aquí)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.size = 12
    
    ' Programa de ejemplo
    ws.Range("A2").value = "section .text"
    ws.Range("A3").value = "global _start"
    ws.Range("A4").value = "_start:"
    ws.Range("A5").value = "mov eax, 10"
    ws.Range("A6").value = "add eax, 20"
    ws.Range("A7").value = "mov ebx, 5"
    ws.Range("A8").value = "sub eax, ebx"

End Sub

