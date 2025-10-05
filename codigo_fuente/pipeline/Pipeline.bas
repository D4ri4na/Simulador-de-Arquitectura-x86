Attribute VB_Name = "Pipeline"
' =============================================
' SIMULADOR DE PIPELINE MEJORADO
' Etapas: IF, ID, EX, MEM, WB
' Visualización clara del flujo de instrucciones
' =============================================

Type PipelineInstruction
    instruction As String
    Opcode As String
    Operand1 As String
    Operand2 As String
    stage As String
    CycleEntered As Long
    CurrentStageCycle As Long
    Color As Long
    Result As String
    Stalled As Boolean
    InstructionNumber As Long
End Type

' Variables globales
Dim Pipeline() As PipelineInstruction
Dim ClockCycle As Long
Dim Instructions() As String
Dim CurrentInstructionIndex As Long
Dim PipelineStages(4) As String
Dim IsPipelineRunning As Boolean
Dim InstructionColors As Collection

' =============================================
' INICIALIZACIÓN
' =============================================

Sub IniciarSimuladorPipeline()
    InitializePipeline
    CreatePipelineDisplay
    UpdatePipelineDisplay
    CreatePipelineLog
    
    MsgBox "? Simulador de Pipeline Inicializado" & vbCrLf & vbCrLf & _
           "Instrucciones cargadas: " & GetInstructionCount() & vbCrLf & vbCrLf & _
           "Controles disponibles:" & vbCrLf & _
           "• EjecutarPipelineCompleto - Simulación automática" & vbCrLf & _
           "• AvanzarCiclo - Paso a paso" & vbCrLf & _
           "• PausarPipeline - Pausar simulación" & vbCrLf & _
           "• ReiniciarPipeline - Reiniciar desde el inicio", vbInformation, "Pipeline Simulator"
End Sub

Sub InitializePipeline()
    ' Configurar etapas
    PipelineStages(0) = "IF"
    PipelineStages(1) = "ID"
    PipelineStages(2) = "EX"
    PipelineStages(3) = "MEM"
    PipelineStages(4) = "WB"
    
    ' Inicializar pipeline vacío
    ReDim Pipeline(0 To 4)
    Dim i As Integer
    For i = 0 To 4
        ClearPipelineSlot i
    Next i
    
    ClockCycle = 0
    CurrentInstructionIndex = 0
    Set InstructionColors = New Collection
    IsPipelineRunning = False
    
    ' Cargar instrucciones
    LoadInstructionsFromSheet
End Sub

Sub ClearPipelineSlot(index As Integer)
    Pipeline(index).instruction = ""
    Pipeline(index).Opcode = ""
    Pipeline(index).Operand1 = ""
    Pipeline(index).Operand2 = ""
    Pipeline(index).stage = ""
    Pipeline(index).CycleEntered = 0
    Pipeline(index).CurrentStageCycle = 0
    Pipeline(index).Color = RGB(255, 255, 255)
    Pipeline(index).Result = ""
    Pipeline(index).Stalled = False
    Pipeline(index).InstructionNumber = 0
End Sub

Sub LoadInstructionsFromSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CodigoPipeline")
    On Error GoTo 0
    
    If ws Is Nothing Then
        CrearEjemploPipeline
        Set ws = ThisWorkbook.Sheets("CodigoPipeline")
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.ROWS.count, "A").End(xlUp).row
    
    ' Buscar primera fila con instrucciones
    Dim firstRow As Long
    firstRow = 1
    Do While firstRow <= lastRow
        If Trim(ws.Cells(firstRow, 1).value) <> "" And _
           Not InStr(1, ws.Cells(firstRow, 1).value, "Ejemplo", vbTextCompare) > 0 Then
            Exit Do
        End If
        firstRow = firstRow + 1
    Loop
    
    If firstRow > lastRow Then
        MsgBox "No se encontraron instrucciones. Creando ejemplo...", vbInformation
        CrearEjemploPipeline
        LoadInstructionsFromSheet
        Exit Sub
    End If
    
    ' Cargar instrucciones
    Dim count As Long
    count = 0
    Dim i As Long
    ReDim Instructions(0 To 100) ' Temporal
    
    For i = firstRow To lastRow
        Dim inst As String
        inst = Trim(ws.Cells(i, 1).value)
        If inst <> "" And Left(inst, 1) <> "'" And Left(inst, 1) <> ";" Then
            Instructions(count) = inst
            count = count + 1
        End If
    Next i
    
    If count > 0 Then
        ReDim Preserve Instructions(0 To count - 1)
    Else
        ReDim Instructions(0 To 0)
        Instructions(0) = "NOP"
    End If
End Sub

Function GetInstructionCount() As Long
    On Error GoTo ErrorHandler
    GetInstructionCount = UBound(Instructions) + 1
    Exit Function
ErrorHandler:
    GetInstructionCount = 0
End Function

' =============================================
' SIMULACIÓN
' =============================================

Sub EjecutarPipelineCompleto()
    IsPipelineRunning = True
    Dim initialCycle As Long
    initialCycle = ClockCycle
    
    Do While IsPipelineRunning And (ClockCycle - initialCycle) < 100
        ClockCycle = ClockCycle + 1
        AdvancePipeline
        UpdatePipelineDisplay
        LogPipelineState
        DoEvents
        
        If AllInstructionsCompleted() Then
            IsPipelineRunning = False
            LogMessage "? SIMULACIÓN COMPLETADA - Todas las instrucciones finalizadas"
            MsgBox "Simulación completada en " & ClockCycle & " ciclos", vbInformation
        End If
        
        Application.Wait (Now + TimeValue("0:00:00.5"))
    Loop
    
    IsPipelineRunning = False
End Sub

Sub AvanzarCiclo()
    If IsPipelineRunning Then
        MsgBox "Pause la simulación automática primero", vbExclamation
        Exit Sub
    End If
    
    ClockCycle = ClockCycle + 1
    AdvancePipeline
    UpdatePipelineDisplay
    LogPipelineState
    
    If AllInstructionsCompleted() Then
        MsgBox "? Todas las instrucciones han sido procesadas en " & ClockCycle & " ciclos", vbInformation
    End If
End Sub

Sub AdvancePipeline()
    ' Mover instrucciones de derecha a izquierda (WB -> IF)
    ' WB: completar y liberar
    If Pipeline(4).stage = "WB" Then
        Pipeline(4).CurrentStageCycle = Pipeline(4).CurrentStageCycle + 1
        If Pipeline(4).CurrentStageCycle >= 1 Then
            LogMessage "? Instrucción completada: " & Pipeline(4).instruction
            ClearPipelineSlot 4
        End If
    End If
    
    ' MEM -> WB
    If Pipeline(3).stage = "MEM" And Pipeline(4).stage = "" Then
        Pipeline(3).CurrentStageCycle = Pipeline(3).CurrentStageCycle + 1
        If Pipeline(3).CurrentStageCycle >= 1 Then
            Pipeline(4) = Pipeline(3)
            Pipeline(4).stage = "WB"
            Pipeline(4).CurrentStageCycle = 0
            ProcessInstructionInStage 4
            ClearPipelineSlot 3
        End If
    ElseIf Pipeline(3).stage = "MEM" Then
        Pipeline(3).Stalled = True
    End If
    
    ' EX -> MEM
    If Pipeline(2).stage = "EX" And Pipeline(3).stage = "" Then
        Pipeline(2).CurrentStageCycle = Pipeline(2).CurrentStageCycle + 1
        If Pipeline(2).CurrentStageCycle >= 1 Then
            Pipeline(3) = Pipeline(2)
            Pipeline(3).stage = "MEM"
            Pipeline(3).CurrentStageCycle = 0
            Pipeline(3).Stalled = False
            ProcessInstructionInStage 3
            ClearPipelineSlot 2
        End If
    ElseIf Pipeline(2).stage = "EX" Then
        Pipeline(2).Stalled = True
    End If
    
    ' ID -> EX
    If Pipeline(1).stage = "ID" And Pipeline(2).stage = "" Then
        Pipeline(1).CurrentStageCycle = Pipeline(1).CurrentStageCycle + 1
        If Pipeline(1).CurrentStageCycle >= 1 Then
            Pipeline(2) = Pipeline(1)
            Pipeline(2).stage = "EX"
            Pipeline(2).CurrentStageCycle = 0
            Pipeline(2).Stalled = False
            ProcessInstructionInStage 2
            ClearPipelineSlot 1
        End If
    ElseIf Pipeline(1).stage = "ID" Then
        Pipeline(1).Stalled = True
    End If
    
    ' IF -> ID
    If Pipeline(0).stage = "IF" And Pipeline(1).stage = "" Then
        Pipeline(0).CurrentStageCycle = Pipeline(0).CurrentStageCycle + 1
        If Pipeline(0).CurrentStageCycle >= 1 Then
            Pipeline(1) = Pipeline(0)
            Pipeline(1).stage = "ID"
            Pipeline(1).CurrentStageCycle = 0
            Pipeline(1).Stalled = False
            ProcessInstructionInStage 1
            ClearPipelineSlot 0
        End If
    ElseIf Pipeline(0).stage = "IF" Then
        Pipeline(0).Stalled = True
    End If
    
    ' Nueva instrucción -> IF
    If Pipeline(0).stage = "" And CurrentInstructionIndex <= UBound(Instructions) Then
        InsertNewInstruction Instructions(CurrentInstructionIndex), CurrentInstructionIndex + 1
        CurrentInstructionIndex = CurrentInstructionIndex + 1
    End If
End Sub

Sub InsertNewInstruction(instruction As String, instNum As Long)
    Pipeline(0).instruction = instruction
    Pipeline(0).stage = "IF"
    Pipeline(0).CycleEntered = ClockCycle
    Pipeline(0).CurrentStageCycle = 0
    Pipeline(0).Stalled = False
    Pipeline(0).InstructionNumber = instNum
    
    ' Asignar color único
    Pipeline(0).Color = GetInstructionColor(instNum)
    
    ProcessInstructionInStage 0
End Sub

Sub ProcessInstructionInStage(stageIndex As Integer)
    Select Case Pipeline(stageIndex).stage
        Case "IF"
            Pipeline(stageIndex).Result = "Fetching..."
            
        Case "ID"
            ParseInstruction stageIndex
            Pipeline(stageIndex).Result = "Decoding: " & Pipeline(stageIndex).Opcode
            
        Case "EX"
            ExecuteInstruction stageIndex
            
        Case "MEM"
            Pipeline(stageIndex).Result = "Memory Access"
            
        Case "WB"
            Pipeline(stageIndex).Result = "Writing Back"
    End Select
End Sub

Sub ParseInstruction(stageIndex As Integer)
    Dim instruction As String
    instruction = Pipeline(stageIndex).instruction
    
    instruction = Split(instruction, ";")(0)
    instruction = Trim(instruction)
    
    Dim parts() As String
    parts = Split(instruction, " ")
    
    If UBound(parts) >= 0 Then
        Pipeline(stageIndex).Opcode = UCase(Replace(Trim(parts(0)), ",", ""))
    End If
    
    Dim operands As String
    If UBound(parts) >= 1 Then
        operands = Join(Array(parts(1), parts(2), parts(3)), " ")
        operands = Replace(operands, ",", "")
        Dim opParts() As String
        opParts = Split(Trim(operands), " ")
        
        If UBound(opParts) >= 0 Then Pipeline(stageIndex).Operand1 = Trim(opParts(0))
        If UBound(opParts) >= 1 Then Pipeline(stageIndex).Operand2 = Trim(opParts(1))
    End If
End Sub

Sub ExecuteInstruction(stageIndex As Integer)
    Dim op As String
    op = Pipeline(stageIndex).Opcode
    
    Select Case op
        Case "MOV", "LOAD"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " ? " & Pipeline(stageIndex).Operand2
        Case "ADD"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " + " & Pipeline(stageIndex).Operand2
        Case "SUB"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " - " & Pipeline(stageIndex).Operand2
        Case "MUL"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " × " & Pipeline(stageIndex).Operand2
        Case "DIV"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " ÷ " & Pipeline(stageIndex).Operand2
        Case "NOP"
            Pipeline(stageIndex).Result = "No Operation"
        Case Else
            Pipeline(stageIndex).Result = "Execute: " & op
    End Select
End Sub

Function AllInstructionsCompleted() As Boolean
    If CurrentInstructionIndex <= UBound(Instructions) Then
        AllInstructionsCompleted = False
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            AllInstructionsCompleted = False
            Exit Function
        End If
    Next i
    
    AllInstructionsCompleted = True
End Function

' =============================================
' VISUALIZACIÓN MEJORADA
' =============================================

Sub CreatePipelineDisplay()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Pipeline")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Pipeline"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Tab.Color = RGB(100, 150, 200)
    
    ' Encabezado principal
    With ws.Range("A1:H1")
        .Merge
        .value = "SIMULADOR DE PIPELINE - ARQUITECTURA DE 5 ETAPAS"
        .Font.Bold = True
        .Font.size = 18
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(50, 80, 120)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
    End With
    
    ' Información del ciclo
    ws.Range("A3").value = "Ciclo de Reloj Actual:"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").value = 0
    ws.Range("B3").Font.size = 16
    ws.Range("B3").Font.Bold = True
    ws.Range("B3").Font.Color = RGB(200, 0, 0)
    
    ws.Range("D3").value = "Instrucciones Totales:"
    ws.Range("D3").Font.Bold = True
    ws.Range("E3").value = GetInstructionCount()
    ws.Range("E3").Font.size = 14
    ws.Range("E3").Font.Bold = True
    
    ' Diagrama de pipeline
    ws.Range("A5").value = "DIAGRAMA DEL PIPELINE - FLUJO DINÁMICO"
    ws.Range("A5").Font.Bold = True
    ws.Range("A5").Font.size = 14
    ws.Range("A5").Font.Color = RGB(50, 80, 120)
    
    ' Encabezados de etapas
    Dim stages As Variant
    stages = Array("Nº", "IF", "ID", "EX", "MEM", "WB", "Estado", "Ciclos")
    Dim col As Integer
    For col = 0 To 7
        With ws.Cells(7, col + 1)
            .value = stages(col)
            .Font.Bold = True
            .Font.size = 12
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(70, 100, 150)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.Weight = xlMedium
        End With
    Next col
    
    ' Descripciones de etapas
    With ws.Cells(8, 2)
        .value = "Instruction Fetch"
        .Font.Italic = True
        .Font.size = 9
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(8, 3)
        .value = "Instruction Decode"
        .Font.Italic = True
        .Font.size = 9
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(8, 4)
        .value = "Execute"
        .Font.Italic = True
        .Font.size = 9
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(8, 5)
        .value = "Memory Access"
        .Font.Italic = True
        .Font.size = 9
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(8, 6)
        .value = "Write Back"
        .Font.Italic = True
        .Font.size = 9
        .HorizontalAlignment = xlCenter
    End With
    
    ' Filas para instrucciones (con espacio visual)
    Dim row As Integer
    For row = 9 To 18
        ws.ROWS(row).RowHeight = 40
        For col = 1 To 8
            With ws.Cells(row, col)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
        Next col
    Next row
    
    ' Leyenda
    ws.Range("A20").value = "LEYENDA"
    ws.Range("A20").Font.Bold = True
    ws.Range("A20").Font.size = 12
    
    ws.Range("A21").value = "Activo"
    ws.Range("B21").Interior.Color = RGB(144, 238, 144)
    ws.Range("A22").value = "Stalled"
    ws.Range("B22").Interior.Color = RGB(255, 100, 100)
    ws.Range("A23").value = "Completado"
    ws.Range("B23").Interior.Color = RGB(200, 255, 200)
    ws.Range("A24").value = "Vacío"
    ws.Range("B24").Interior.Color = RGB(240, 240, 240)
    
    ' Próximas instrucciones
    ws.Range("D20").value = "PRÓXIMAS INSTRUCCIONES"
    ws.Range("D20").Font.Bold = True
    ws.Range("D20").Font.size = 12
    
    ' Ajustar columnas
    ws.Columns("A:A").ColumnWidth = 8
    ws.Columns("B:F").ColumnWidth = 18
    ws.Columns("G:G").ColumnWidth = 25
    ws.Columns("H:H").ColumnWidth = 12
    
    ' Crear botones de control
    CreatePipelineButtons ws
    
    ' Proteger formato (opcional)
    ws.Cells.Locked = False
End Sub

Sub CreatePipelineButtons(ws As Worksheet)
    ' Limpiar botones existentes
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    ' Botón Ejecutar Todo
    Dim btn As Button
    Set btn = ws.Buttons.Add(50, 450, 100, 25)
    btn.OnAction = "EjecutarPipelineCompleto"
    btn.Characters.Text = "Ejecutar Todo"
    
    ' Botón Avanzar Ciclo
    Set btn = ws.Buttons.Add(160, 450, 100, 25)
    btn.OnAction = "AvanzarCiclo"
    btn.Characters.Text = "Avanzar Ciclo"
    
    ' Botón Pausar
    Set btn = ws.Buttons.Add(270, 450, 80, 25)
    btn.OnAction = "PausarPipeline"
    btn.Characters.Text = "Pausar"
    
    ' Botón Reiniciar
    Set btn = ws.Buttons.Add(360, 450, 80, 25)
    btn.OnAction = "ReiniciarPipeline"
    btn.Characters.Text = "Reiniciar"
End Sub

Sub UpdatePipelineDisplay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Pipeline")
    
    ' Actualizar ciclo
    ws.Range("B3").value = ClockCycle
    
    ' Limpiar área de pipeline
    Dim row As Integer, col As Integer
    For row = 9 To 18
        For col = 1 To 8
            ws.Cells(row, col).value = ""
            ws.Cells(row, col).Interior.Color = RGB(240, 240, 240)
            ws.Cells(row, col).Font.Bold = False
            ws.Cells(row, col).Font.Color = RGB(0, 0, 0)
        Next col
    Next row
    
    ' Mostrar instrucciones en pipeline con su historial
    Dim instructionRows As Object
    Set instructionRows = CreateObject("Scripting.Dictionary")
    
    ' Agrupar instrucciones por número
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            Dim instNum As Long
            instNum = Pipeline(i).InstructionNumber
            
            If Not instructionRows.Exists(instNum) Then
                instructionRows.Add instNum, instructionRows.count + 9
            End If
            
            Dim currentRow As Long
            currentRow = instructionRows(instNum)
            
            ' Número de instrucción
            ws.Cells(currentRow, 1).value = "I" & instNum
            ws.Cells(currentRow, 1).Font.Bold = True
            ws.Cells(currentRow, 1).Interior.Color = Pipeline(i).Color
            
            ' Mostrar etapa actual y etapas completadas
            For col = 2 To 6
                Dim stageCol As String
                stageCol = ws.Cells(7, col).value
                
                If stageCol = Pipeline(i).stage Then
                    ' Etapa actual
                    ws.Cells(currentRow, col).value = Pipeline(i).instruction
                    ws.Cells(currentRow, col).Font.Bold = True
                    
                    If Pipeline(i).Stalled Then
                        ws.Cells(currentRow, col).Interior.Color = RGB(255, 100, 100)
                        ws.Cells(currentRow, col).value = Pipeline(i).instruction & vbLf & "[STALL]"
                        ws.Cells(currentRow, col).Font.Color = RGB(255, 255, 255)
                    Else
                        ws.Cells(currentRow, col).Interior.Color = Pipeline(i).Color
                    End If
                ElseIf GetStageOrder(stageCol) < GetStageOrder(Pipeline(i).stage) Then
                    ' Etapa completada
                    ws.Cells(currentRow, col).value = "?"
                    ws.Cells(currentRow, col).Interior.Color = RGB(200, 255, 200)
                    ws.Cells(currentRow, col).Font.size = 14
                    ws.Cells(currentRow, col).Font.Color = RGB(0, 100, 0)
                End If
            Next col
            
            ' Estado/Resultado
            ws.Cells(currentRow, 7).value = Pipeline(i).Result
            ws.Cells(currentRow, 7).Font.size = 9
            ws.Cells(currentRow, 7).Interior.Color = Pipeline(i).Color
            
            ' Ciclos en etapa
            ws.Cells(currentRow, 8).value = Pipeline(i).CurrentStageCycle + 1 & "/1"
            ws.Cells(currentRow, 8).Font.Bold = True
        End If
    Next i
    
    ' Mostrar próximas instrucciones
    Dim nextRow As Integer
    nextRow = 21
    For i = CurrentInstructionIndex To UBound(Instructions)
        If i <= CurrentInstructionIndex + 5 Then
            ws.Cells(nextRow, 4).value = "I" & (i + 1) & ": " & Instructions(i)
            ws.Cells(nextRow, 4).Font.size = 10
            ws.Cells(nextRow, 4).Interior.Color = GetInstructionColor(i + 1)
            nextRow = nextRow + 1
        End If
    Next i
    
    ' Mostrar estadísticas
    ws.Range("G20").value = "ESTADÍSTICAS"
    ws.Range("G20").Font.Bold = True
    ws.Range("G21").value = "Ciclos: " & ClockCycle
    ws.Range("G22").value = "Instrucciones completadas: " & (CurrentInstructionIndex - CountInstructionsInPipeline())
    ws.Range("G23").value = "Instrucciones en pipeline: " & CountInstructionsInPipeline()
    
    ' Mostrar flujo gráfico
    ShowPipelineFlow ws, instructionRows
End Sub

Function CountInstructionsInPipeline() As Integer
    Dim count As Integer
    count = 0
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            count = count + 1
        End If
    Next i
    CountInstructionsInPipeline = count
End Function

Sub ShowPipelineFlow(ws As Worksheet, instructionRows As Object)
    ' Mostrar líneas de flujo entre etapas
    Dim row As Integer
    For row = 9 To 18
        If ws.Cells(row, 1).value <> "" Then
            ' Dibujar flechas entre etapas completadas
            For col = 2 To 5
                If ws.Cells(row, col).value = "?" And ws.Cells(row, col + 1).value <> "" Then
                    ws.Cells(row, col).value = "? ?"
                    ws.Cells(row, col).Font.size = 10
                End If
            Next col
        End If
    Next row
End Sub

Function GetStageOrder(stage As String) As Integer
    Select Case stage
        Case "IF": GetStageOrder = 0
        Case "ID": GetStageOrder = 1
        Case "EX": GetStageOrder = 2
        Case "MEM": GetStageOrder = 3
        Case "WB": GetStageOrder = 4
        Case Else: GetStageOrder = -1
    End Select
End Function

Function GetInstructionColor(instNum As Long) As Long
    Dim colors As Variant
    colors = Array( _
        RGB(173, 216, 230), _    ' Light Blue
        RGB(255, 182, 193), _    ' Light Pink
        RGB(221, 160, 221), _    ' Plum
        RGB(255, 218, 185), _    ' Peach
        RGB(176, 224, 230), _    ' Powder Blue
        RGB(240, 230, 140), _    ' Khaki
        RGB(152, 251, 152), _    ' Pale Green
        RGB(255, 228, 196), _    ' Bisque
        RGB(230, 230, 250), _    ' Lavender
        RGB(245, 222, 179) _     ' Wheat
    )
    
    GetInstructionColor = colors((instNum - 1) Mod 10)
End Function

' =============================================
' LOGGING
' =============================================

Sub CreatePipelineLog()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LogPipeline")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Pipeline"))
        ws.Name = "LogPipeline"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Tab.Color = RGB(150, 150, 100)
    
    ' Encabezados
    Dim headers As Variant
    headers = Array("Ciclo", "IF", "ID", "EX", "MEM", "WB", "Eventos")
    Dim col As Integer
    For col = 0 To 6
        With ws.Cells(1, col + 1)
            .value = headers(col)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlCenter
        End With
    Next col
    
    ws.Columns.AutoFit
End Sub

Sub LogPipelineState()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LogPipeline")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = ClockCycle
    ws.Cells(nextRow, 1).Font.Bold = True
    
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            ws.Cells(nextRow, 2 + i).value = "I" & Pipeline(i).InstructionNumber & ": " & Pipeline(i).instruction
            ws.Cells(nextRow, 2 + i).Interior.Color = Pipeline(i).Color
            
            If Pipeline(i).Stalled Then
                ws.Cells(nextRow, 2 + i).value = ws.Cells(nextRow, 2 + i).value & " [STALL]"
                ws.Cells(nextRow, 2 + i).Font.Color = RGB(200, 0, 0)
            End If
        End If
    Next i
    
    ' Eventos
    Dim events As String
    events = ""
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            If events <> "" Then events = events & " | "
            events = events & Pipeline(i).stage & ":I" & Pipeline(i).InstructionNumber
            If Pipeline(i).Stalled Then events = events & "(STALL)"
        End If
    Next i
    ws.Cells(nextRow, 7).value = events
End Sub

Sub LogMessage(message As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LogPipeline")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = ClockCycle
    ws.Cells(nextRow, 7).value = message
    ws.Cells(nextRow, 7).Font.Bold = True
    ws.Cells(nextRow, 7).Interior.Color = RGB(255, 255, 150)
End Sub

' =============================================
' CONTROLES
' =============================================

Sub PausarPipeline()
    IsPipelineRunning = False
    LogMessage "?? SIMULACIÓN PAUSADA"
    MsgBox "Simulación pausada en el ciclo " & ClockCycle, vbInformation
End Sub

Sub ReiniciarPipeline()
    InitializePipeline
    CreatePipelineDisplay
    UpdatePipelineDisplay
    CreatePipelineLog
    LogMessage "?? PIPELINE REINICIADO"
    MsgBox "Pipeline reiniciado", vbInformation
End Sub

' =============================================
' EJEMPLO
' =============================================

Sub CrearEjemploPipeline()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CodigoPipeline")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "CodigoPipeline"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Tab.Color = RGB(150, 200, 150)
    
    ws.Range("A1").value = "PROGRAMA DE EJEMPLO - INSTRUCCIONES PARA PIPELINE"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.size = 14
    
    ' Instrucciones de ejemplo
    ws.Range("A3").value = "MOV R1, 10"
    ws.Range("A4").value = "ADD R2, R1, 5"
    ws.Range("A5").value = "SUB R3, R2, 3"
    ws.Range("A6").value = "MUL R4, R1, R2"
    ws.Range("A7").value = "DIV R5, R4, 2"
    ws.Range("A8").value = "ADD R6, R3, R5"
    ws.Range("A9").value = "MOV R7, 100"
    ws.Range("A10").value = "SUB R8, R7, R6"
    ws.Range("A11").value = "NOP"
    ws.Range("A12").value = "ADD R9, R8, 1"
    
    ws.Columns("A:A").ColumnWidth = 30
    
    MsgBox "? Programa de ejemplo creado en 'CodigoPipeline'", vbInformation
End Sub




