Attribute VB_Name = "Pipeline"
' Simulador de Pipeline - IF, ID, EX, MEM, WB
Type PipelineInstruction
    instruction As String
    Opcode As String
    Operand1 As String
    Operand2 As String
    stage As String ' "IF", "ID", "EX", "MEM", "WB", "DONE"
    CycleEntered As Long
    CurrentStageCycle As Long
    Color As Long
    Result As String
    Stalled As Boolean
End Type

' Variables globales del pipeline
Dim Pipeline() As PipelineInstruction
Dim ClockCycle As Long
Dim Instructions() As String
Dim CurrentInstructionIndex As Long
Dim PipelineStages(4) As String
Dim Hazards As Collection
Dim IsPipelineRunning As Boolean

' =============================================
' INICIALIZACIÓN DEL PIPELINE
' =============================================

Sub InitializePipeline()
    ' Configurar etapas del pipeline
    PipelineStages(0) = "IF"   ' Instruction Fetch
    PipelineStages(1) = "ID"   ' Instruction Decode
    PipelineStages(2) = "EX"   ' Execute
    PipelineStages(3) = "MEM"  ' Memory Access
    PipelineStages(4) = "WB"   ' Write Back
    
    ' Inicializar pipeline vacío
    ReDim Pipeline(0 To 4)
    Dim i As Integer
    For i = 0 To 4
        Pipeline(i).instruction = ""
        Pipeline(i).Opcode = ""
        Pipeline(i).Operand1 = ""
        Pipeline(i).Operand2 = ""
        Pipeline(i).stage = ""
        Pipeline(i).CycleEntered = 0
        Pipeline(i).CurrentStageCycle = 0
        Pipeline(i).Color = RGB(255, 255, 255)
        Pipeline(i).Result = ""
        Pipeline(i).Stalled = False
    Next i
    
    ClockCycle = 0
    CurrentInstructionIndex = 0
    Set Hazards = New Collection
    
    ' Cargar instrucciones desde hoja
    LoadInstructionsFromSheet
    
    CreatePipelineDisplay
    UpdatePipelineDisplay
    CreatePipelineLog
    
    MsgBox "Pipeline inicializado con " & GetInstructionCount() & " instrucciones", vbInformation
End Sub

Sub LoadInstructionsFromSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CodigoPipeline")
    
    ' Contar instrucciones
    Dim lastRow As Long
    lastRow = ws.Cells(ws.ROWS.count, "A").End(xlUp).row
    
    If lastRow = 0 Then
        ' Crear instrucciones de ejemplo
        ReDim Instructions(0 To 4)
        Instructions(0) = "MOV R1, 10"
        Instructions(1) = "ADD R2, R1, 5"
        Instructions(2) = "SUB R3, R2, 3"
        Instructions(3) = "MUL R4, R1, R2"
        Instructions(4) = "DIV R5, R4, 2"
    Else
        ReDim Instructions(0 To lastRow - 1)
        Dim i As Long
        For i = 1 To lastRow
            Instructions(i - 1) = Trim(ws.Cells(i, 1).value)
        Next i
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
' SIMULACIÓN DEL PIPELINE
' =============================================

Sub RunPipelineComplete()
    IsPipelineRunning = True
    CreatePipelineLog
    
    While IsPipelineRunning And ClockCycle < 100 ' Límite de seguridad
        ClockCycle = ClockCycle + 1
        AdvancePipeline
        UpdatePipelineDisplay
        LogPipelineState
        DoEvents
        
        ' Verificar si todas las instrucciones han terminado
        If AllInstructionsCompleted() Then
            IsPipelineRunning = False
            LogMessage "SIMULACIÓN COMPLETADA - Todas las instrucciones finalizadas"
        End If
        
        ' Pequeña pausa para visualización
        Application.Wait (Now + TimeValue("0:00:01"))
    Wend
    
    If ClockCycle >= 100 Then
        LogMessage "SIMULACIÓN DETENIDA - Límite de ciclos alcanzado"
    End If
End Sub

Sub AdvancePipeline()
    ' Avanzar las etapas en orden inverso (WB -> MEM -> EX -> ID -> IF)
    AdvanceStage "WB"
    AdvanceStage "MEM"
    AdvanceStage "EX"
    AdvanceStage "ID"
    AdvanceStage "IF"
    
    ' Insertar nueva instrucción en IF si hay disponible
    If CurrentInstructionIndex <= UBound(Instructions) Then
        If Pipeline(0).stage = "" Then ' IF está vacío
            InsertNewInstruction Instructions(CurrentInstructionIndex)
            CurrentInstructionIndex = CurrentInstructionIndex + 1
        End If
    End If
End Sub

Sub AdvanceStage(stage As String)
    Dim stageIndex As Integer
    stageIndex = GetStageIndex(stage)
    
    If Pipeline(stageIndex).stage = stage Then
        ' Incrementar contador de ciclos en esta etapa
        Pipeline(stageIndex).CurrentStageCycle = Pipeline(stageIndex).CurrentStageCycle + 1
        
        ' Verificar si la instrucción puede avanzar a la siguiente etapa
        If CanAdvanceToNextStage(stageIndex) Then
            MoveToNextStage stageIndex
        End If
    End If
End Sub

Function CanAdvanceToNextStage(currentStageIndex As Integer) As Boolean
    Dim nextStageIndex As Integer
    nextStageIndex = currentStageIndex + 1
    
    If nextStageIndex > 4 Then ' WB no tiene siguiente etapa
        CanAdvanceToNextStage = True
        Exit Function
    End If
    
    ' Verificar si la siguiente etapa está ocupada
    If Pipeline(nextStageIndex).stage = "" Or Pipeline(nextStageIndex).stage = "DONE" Then
        CanAdvanceToNextStage = True
    Else
        ' La siguiente etapa está ocupada - crear stall
        Pipeline(currentStageIndex).Stalled = True
        CanAdvanceToNextStage = False
    End If
End Function

Sub MoveToNextStage(currentStageIndex As Integer)
    Dim nextStageIndex As Integer
    nextStageIndex = currentStageIndex + 1
    
    If nextStageIndex <= 4 Then
        ' Mover a la siguiente etapa
        Pipeline(nextStageIndex) = Pipeline(currentStageIndex)
        Pipeline(nextStageIndex).stage = PipelineStages(nextStageIndex)
        Pipeline(nextStageIndex).CurrentStageCycle = 0
        Pipeline(nextStageIndex).Stalled = False
        
        ' Procesar la instrucción en la nueva etapa
        ProcessInstructionInStage nextStageIndex
    Else
        ' Instrucción completada (salida de WB)
        Pipeline(currentStageIndex).stage = "DONE"
    End If
    
    ' Limpiar etapa anterior
    Pipeline(currentStageIndex).instruction = ""
    Pipeline(currentStageIndex).stage = ""
    Pipeline(currentStageIndex).CurrentStageCycle = 0
    Pipeline(currentStageIndex).Stalled = False
End Sub

Sub InsertNewInstruction(instruction As String)
    Pipeline(0).instruction = instruction
    Pipeline(0).Opcode = ""
    Pipeline(0).Operand1 = ""
    Pipeline(0).Operand2 = ""
    Pipeline(0).stage = "IF"
    Pipeline(0).CycleEntered = ClockCycle
    Pipeline(0).CurrentStageCycle = 0
    Pipeline(0).Color = GetRandomColor()
    Pipeline(0).Result = ""
    Pipeline(0).Stalled = False
    
    ' Procesar en etapa IF
    ProcessInstructionInStage 0
End Sub

Sub ProcessInstructionInStage(stageIndex As Integer)
    Dim instruction As String
    instruction = Pipeline(stageIndex).instruction
    
    Select Case PipelineStages(stageIndex)
        Case "IF"
            ' Instruction Fetch - simplemente capturar la instrucción
            Pipeline(stageIndex).Result = "Instrucción capturada"
            
        Case "ID"
            ' Instruction Decode - parsear la instrucción
            ParseInstruction stageIndex
            Pipeline(stageIndex).Result = "Decodificada: " & Pipeline(stageIndex).Opcode
            
        Case "EX"
            ' Execute - ejecutar operación
            ExecuteInstruction stageIndex
            Pipeline(stageIndex).Result = "Ejecutado: " & Pipeline(stageIndex).Result
            
        Case "MEM"
            ' Memory Access - acceso a memoria (simulado)
            Pipeline(stageIndex).Result = "Acceso MEM completado"
            
        Case "WB"
            ' Write Back - escribir resultados
            Pipeline(stageIndex).Result = "Write Back completado"
    End Select
End Sub

Sub ParseInstruction(stageIndex As Integer)
    Dim instruction As String
    instruction = Pipeline(stageIndex).instruction
    
    ' Limpiar comentarios
    instruction = Split(instruction, ";")(0)
    
    Dim parts() As String
    parts = Split(Trim(instruction), " ")
    
    If UBound(parts) >= 0 Then
        Pipeline(stageIndex).Opcode = UCase(Trim(parts(0)))
    End If
    
    If UBound(parts) >= 1 Then
        Pipeline(stageIndex).Operand1 = Trim(parts(1))
    End If
    
    If UBound(parts) >= 2 Then
        Pipeline(stageIndex).Operand2 = Trim(parts(2))
    End If
End Sub

Sub ExecuteInstruction(stageIndex As Integer)
    Dim Opcode As String
    Opcode = Pipeline(stageIndex).Opcode
    
    Select Case Opcode
        Case "MOV"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " = " & Pipeline(stageIndex).Operand2
        Case "ADD"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " + " & Pipeline(stageIndex).Operand2
        Case "SUB"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " - " & Pipeline(stageIndex).Operand2
        Case "MUL"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " * " & Pipeline(stageIndex).Operand2
        Case "DIV"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).Operand1 & " / " & Pipeline(stageIndex).Operand2
        Case Else
            Pipeline(stageIndex).Result = "Operación: " & Opcode
    End Select
End Sub

Function GetStageIndex(stage As String) As Integer
    Dim i As Integer
    For i = 0 To 4
        If PipelineStages(i) = stage Then
            GetStageIndex = i
            Exit Function
        End If
    Next i
    GetStageIndex = -1
End Function

Function AllInstructionsCompleted() As Boolean
    ' Verificar si hay instrucciones pendientes por cargar
    If CurrentInstructionIndex <= UBound(Instructions) Then
        AllInstructionsCompleted = False
        Exit Function
    End If
    
    ' Verificar si todas las etapas del pipeline están vacías
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" And Pipeline(i).stage <> "DONE" Then
            AllInstructionsCompleted = False
            Exit Function
        End If
    Next i
    
    AllInstructionsCompleted = True
End Function

' =============================================
' VISUALIZACIÓN DEL PIPELINE
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
    CreatePipelineDiagram ws
End Sub

Sub CreatePipelineDiagram(ws As Worksheet)
    ' Título
    ws.Cells(1, 1).value = "SIMULADOR DE PIPELINE - 5 ETAPAS"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.size = 16
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Range("A1:F1").Merge
    
    ' Ciclo actual
    ws.Cells(2, 1).value = "Ciclo de Reloj:"
    ws.Cells(2, 2).value = ClockCycle
    ws.Cells(2, 2).Font.Bold = True
    ws.Cells(2, 2).Font.size = 14
    
    ' Instrucciones cargadas
    ws.Cells(3, 1).value = "Instrucciones:"
    ws.Cells(3, 2).value = GetInstructionCount()
    
    ' Encabezados de etapas
    ws.Cells(5, 1).value = "ETAPA"
    ws.Cells(5, 2).value = "IF"
    ws.Cells(5, 3).value = "ID"
    ws.Cells(5, 4).value = "EX"
    ws.Cells(5, 5).value = "MEM"
    ws.Cells(5, 6).value = "WB"
    
    ' Formato de encabezados
    Dim headerRange As Range
    Set headerRange = ws.Range("A5:F5")
    headerRange.Font.Bold = True
    headerRange.Interior.Color = RGB(150, 150, 150)
    headerRange.HorizontalAlignment = xlCenter
    
    ' Descripciones de etapas
    ws.Cells(6, 1).value = "Descripción"
    ws.Cells(6, 2).value = "Instruction Fetch"
    ws.Cells(6, 3).value = "Instruction Decode"
    ws.Cells(6, 4).value = "Execute"
    ws.Cells(6, 5).value = "Memory Access"
    ws.Cells(6, 6).value = "Write Back"
    
    ' Borde para el diagrama
    Set headerRange = ws.Range("A5:F11")
    headerRange.Borders.LineStyle = xlContinuous
    
    ws.Columns.AutoFit
End Sub

Sub UpdatePipelineDisplay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Pipeline")
    
    ' Actualizar ciclo actual
    ws.Cells(2, 2).value = ClockCycle
    
    ' Mostrar instrucciones en cada etapa
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            ' Instrucción
            ws.Cells(7 + i, 2 + i).value = Pipeline(i).instruction
            ws.Cells(7 + i, 2 + i).Interior.Color = Pipeline(i).Color
            
            ' Información adicional
            ws.Cells(7 + i, 1).value = "Inst " & (Pipeline(i).CycleEntered + 1)
            ws.Cells(8 + i, 2 + i).value = Pipeline(i).Result
            
            ' Resaltar stalls
            If Pipeline(i).Stalled Then
                ws.Cells(7 + i, 2 + i).Interior.Color = RGB(255, 100, 100) ' Rojo para stall
                ws.Cells(7 + i, 2 + i).value = Pipeline(i).instruction & " [STALL]"
            End If
        Else
            ' Limpiar celda si no hay instrucción
            ws.Cells(7 + i, 2 + i).value = ""
            ws.Cells(7 + i, 2 + i).Interior.ColorIndex = 0
            ws.Cells(8 + i, 2 + i).value = ""
        End If
    Next i
    
    ' Mostrar próximas instrucciones
    ws.Cells(12, 1).value = "Próximas Instrucciones:"
    Dim nextRow As Long
    nextRow = 13
    Dim j As Long
    For j = CurrentInstructionIndex To UBound(Instructions)
        If j <= CurrentInstructionIndex + 5 Then ' Mostrar solo las próximas 5
            ws.Cells(nextRow, 1).value = Instructions(j)
            nextRow = nextRow + 1
        End If
    Next j
    
    ws.Columns.AutoFit
End Sub

Function GetRandomColor() As Long
    ' Generar color pastel aleatorio
    Dim r As Integer, g As Integer, b As Integer
    r = Int((200 - 150 + 1) * Rnd + 150)
    g = Int((200 - 150 + 1) * Rnd + 150)
    b = Int((200 - 150 + 1) * Rnd + 150)
    GetRandomColor = RGB(r, g, b)
End Function

' =============================================
' LOGGING Y CONTROL
' =============================================

Sub CreatePipelineLog()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LogPipeline")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "LogPipeline"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ' Encabezados del log
    ws.Cells(1, 1).value = "Ciclo"
    ws.Cells(1, 2).value = "IF"
    ws.Cells(1, 3).value = "ID"
    ws.Cells(1, 4).value = "EX"
    ws.Cells(1, 5).value = "MEM"
    ws.Cells(1, 6).value = "WB"
    ws.Cells(1, 7).value = "Eventos"
    
    ' Formato de encabezados
    Dim headerRange As Range
    Set headerRange = ws.Range("A1:G1")
    headerRange.Font.Bold = True
    headerRange.Interior.Color = RGB(200, 200, 200)
    
    ws.Columns.AutoFit
End Sub

Sub LogPipelineState()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LogPipeline")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = ClockCycle
    
    ' Log de cada etapa
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            ws.Cells(nextRow, 2 + i).value = Pipeline(i).instruction
            ws.Cells(nextRow, 2 + i).Interior.Color = Pipeline(i).Color
            
            If Pipeline(i).Stalled Then
                ws.Cells(nextRow, 2 + i).value = Pipeline(i).instruction & " [STALL]"
                ws.Cells(nextRow, 2 + i).Interior.Color = RGB(255, 100, 100)
            End If
        End If
    Next i
    
    ' Log de eventos
    Dim eventText As String
    eventText = ""
    For i = 0 To 4
        If Pipeline(i).stage <> "" And Not Pipeline(i).Stalled Then
            If eventText <> "" Then eventText = eventText & ", "
            eventText = eventText & Pipeline(i).stage & ": " & Pipeline(i).instruction
        End If
    Next i
    
    ws.Cells(nextRow, 7).value = eventText
    
    ws.Columns.AutoFit
End Sub

Sub LogMessage(message As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LogPipeline")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = ClockCycle
    ws.Cells(nextRow, 7).value = message
    ws.Cells(nextRow, 7).Font.Bold = True
    ws.Cells(nextRow, 7).Interior.Color = RGB(255, 255, 0)
    
    ws.Columns.AutoFit
End Sub

' =============================================
' INTERFAZ DE CONTROL
' =============================================

Sub IniciarSimuladorPipeline()
    InitializePipeline
    MsgBox "Simulador de Pipeline inicializado." & vbCrLf & _
           "Use 'EjecutarPipelineCompleto' para simulación automática o " & _
           "'AvanzarCiclo' para paso a paso.", vbInformation
End Sub

Sub EjecutarPipelineCompleto()
    RunPipelineComplete
End Sub

Sub AvanzarCiclo()
    If Not IsPipelineRunning Then
        ClockCycle = ClockCycle + 1
        AdvancePipeline
        UpdatePipelineDisplay
        LogPipelineState
        
        If AllInstructionsCompleted() Then
            MsgBox "Todas las instrucciones han sido procesadas.", vbInformation
        End If
    Else
        MsgBox "Detenga la simulación automática primero.", vbExclamation
    End If
End Sub

Sub PausarPipeline()
    IsPipelineRunning = False
    LogMessage "SIMULACIÓN PAUSADA"
End Sub

Sub ReiniciarPipeline()
    InitializePipeline
    LogMessage "PIPELINE REINICIADO"
End Sub

' =============================================
' EJEMPLO DE USO
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
    ws.Cells(1, 1).value = "Ejemplo de Programa para Pipeline"
    ws.Cells(1, 1).Font.Bold = True
    
    ' Instrucciones de ejemplo
    ws.Cells(3, 1).value = "MOV R1, 10"
    ws.Cells(4, 1).value = "ADD R2, R1, 5"
    ws.Cells(5, 1).value = "SUB R3, R2, 3"
    ws.Cells(6, 1).value = "MUL R4, R1, R2"
    ws.Cells(7, 1).value = "DIV R5, R4, 2"
    ws.Cells(8, 1).value = "ADD R6, R3, R5"
    ws.Cells(9, 1).value = "MOV R7, 100"
    
    ws.Columns.AutoFit
    MsgBox "Ejemplo de programa creado en la pestaña 'CodigoPipeline'", vbInformation
End Sub

