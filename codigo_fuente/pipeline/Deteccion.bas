Attribute VB_Name = "Deteccion"
' =============================================
' SIMULADOR DE PIPELINE CON DETECCIÓN DE RIESGOS DE DATOS
' =============================================

Type PipelineInstruction
    instruction As String
    Opcode As String
    Operand1 As String
    Operand2 As String
    Operand3 As String
    stage As String ' "IF", "ID", "EX", "MEM", "WB", "DONE"
    CycleEntered As Long
    CurrentStageCycle As Long
    Color As Long
    Result As String
    Stalled As Boolean
    StallCycles As Long
    DestinationReg As String
    SourceReg1 As String
    SourceReg2 As String
    InstructionNumber As Long
End Type

' Variables globales del pipeline
Dim Pipeline() As PipelineInstruction
Dim ClockCycle As Long
Dim Instructions() As String
Dim CurrentInstructionIndex As Long
Dim PipelineStages(4) As String
Dim IsPipelineRunning As Boolean
Dim RegisterStatus(15) As String ' Estado de registros R0-R15
Dim TotalStallCycles As Long

' =============================================
' INICIALIZACIÓN DEL PIPELINE
' =============================================

Sub IniciarSimuladorConRiesgos()
    InitializePipeline
    MsgBox "?? Simulador de Pipeline con Detección de Riesgos inicializado." & vbCrLf & _
           "Se detectarán automáticamente riesgos RAW y se insertarán burbujas.", vbInformation
End Sub

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
        ClearPipelineStage i
    Next i
    
    ' Inicializar estado de registros
    For i = 0 To 15
        RegisterStatus(i) = "READY"
    Next i
    
    ClockCycle = 0
    CurrentInstructionIndex = 0
    TotalStallCycles = 0
    
    ' Cargar instrucciones desde hoja
    LoadInstructionsFromSheet
    
    CreateUnifiedPipelineDisplay
    UpdatePipelineDisplay
    
    LogMessage "?? Pipeline inicializado con " & GetInstructionCount() & " instrucciones"
End Sub

Sub ClearPipelineStage(stageIndex As Integer)
    Pipeline(stageIndex).instruction = ""
    Pipeline(stageIndex).Opcode = ""
    Pipeline(stageIndex).Operand1 = ""
    Pipeline(stageIndex).Operand2 = ""
    Pipeline(stageIndex).Operand3 = ""
    Pipeline(stageIndex).stage = ""
    Pipeline(stageIndex).CycleEntered = 0
    Pipeline(stageIndex).CurrentStageCycle = 0
    Pipeline(stageIndex).Color = RGB(255, 255, 255)
    Pipeline(stageIndex).Result = ""
    Pipeline(stageIndex).Stalled = False
    Pipeline(stageIndex).StallCycles = 0
    Pipeline(stageIndex).DestinationReg = ""
    Pipeline(stageIndex).SourceReg1 = ""
    Pipeline(stageIndex).SourceReg2 = ""
    Pipeline(stageIndex).InstructionNumber = 0
End Sub

Sub LoadInstructionsFromSheet()
    ' Instrucciones de ejemplo que causan riesgos
    ReDim Instructions(0 To 5)
    Instructions(0) = "ADD R1, R2, R3"    ' Escribe R1
    Instructions(1) = "SUB R4, R1, R5"    ' Lee R1 - RAW hazard!
    Instructions(2) = "MUL R6, R7, R8"    ' Independiente
    Instructions(3) = "DIV R9, R1, R10"   ' Lee R1 - Otro RAW hazard!
    Instructions(4) = "MOV R11, R12"      ' Independiente
    Instructions(5) = "ADD R13, R4, R1"   ' Lee R1 y R4 - Múltiples hazards!
End Sub

Function GetInstructionCount() As Long
    On Error GoTo ErrorHandler
    GetInstructionCount = UBound(Instructions) + 1
    Exit Function
ErrorHandler:
    GetInstructionCount = 0
End Function

' =============================================
' DETECCIÓN DE RIESGOS DE DATOS
' =============================================

Function CheckForDataHazards(currentStageIndex As Integer) As String
    ' Verificar riesgos de datos solo en etapa ID
    If currentStageIndex <> 1 Then ' Solo verificamos en etapa ID
        CheckForDataHazards = ""
        Exit Function
    End If
    
    Dim currentInstr As PipelineInstruction
    currentInstr = Pipeline(currentStageIndex)
    
    ' Verificar dependencias con instrucciones en etapas posteriores
    Dim hazardType As String
    
    ' Verificar dependencias con instrucción en EX (RAW hazard)
    If Pipeline(2).stage = "EX" Then
        hazardType = CheckRAWHazard(currentInstr, Pipeline(2), "EX")
        If hazardType <> "" Then
            CheckForDataHazards = "RAW-EX: " & hazardType
            Exit Function
        End If
    End If
    
    ' Verificar dependencias con instrucción en MEM (RAW hazard)
    If Pipeline(3).stage = "MEM" Then
        hazardType = CheckRAWHazard(currentInstr, Pipeline(3), "MEM")
        If hazardType <> "" Then
            CheckForDataHazards = "RAW-MEM: " & hazardType
            Exit Function
        End If
    End If
    
    CheckForDataHazards = ""
End Function

Function CheckRAWHazard(currentInstr As PipelineInstruction, previousInstr As PipelineInstruction, stage As String) As String
    ' Verificar dependencias RAW (Read After Write)
    Dim dependency As String
    
    ' Verificar si la instrucción actual lee un registro que la anterior escribe
    If previousInstr.DestinationReg <> "" Then
        ' Verificar primer operando fuente
        If currentInstr.SourceReg1 <> "" And currentInstr.SourceReg1 = previousInstr.DestinationReg Then
            dependency = currentInstr.SourceReg1 & " (Operando 1)"
        End If
        
        ' Verificar segundo operando fuente
        If currentInstr.SourceReg2 <> "" And currentInstr.SourceReg2 = previousInstr.DestinationReg Then
            If dependency <> "" Then dependency = dependency & ", "
            dependency = dependency & currentInstr.SourceReg2 & " (Operando 2)"
        End If
    End If
    
    CheckRAWHazard = dependency
End Function

Sub HandleDataHazard(stageIndex As Integer, hazardMessage As String)
    ' Insertar burbuja (stall) en el pipeline
    Pipeline(stageIndex).Stalled = True
    Pipeline(stageIndex).StallCycles = Pipeline(stageIndex).StallCycles + 1
    TotalStallCycles = TotalStallCycles + 1
    
    ' Crear burbuja en IF también para detener el avance
    If stageIndex = 1 Then ' Hazard detectado en ID
        Pipeline(0).Stalled = True ' Stall IF también
    End If
    
    ' Log del hazard
    LogHazard hazardMessage, Pipeline(stageIndex).instruction, GetAffectedInstruction(hazardMessage)
    
    ' Actualizar visualización
    UpdateHazardDisplay hazardMessage
End Sub

Function GetAffectedInstruction(hazardMessage As String) As String
    ' Extraer la instrucción afectada del mensaje de hazard
    If InStr(hazardMessage, "EX") > 0 And Pipeline(2).stage = "EX" Then
        GetAffectedInstruction = Pipeline(2).instruction
    ElseIf InStr(hazardMessage, "MEM") > 0 And Pipeline(3).stage = "MEM" Then
        GetAffectedInstruction = Pipeline(3).instruction
    Else
        GetAffectedInstruction = "Desconocida"
    End If
End Function

Sub ResolveHazards()
    ' Verificar y resolver hazards en cada etapa
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).Stalled Then
            Pipeline(i).StallCycles = Pipeline(i).StallCycles + 1
            Pipeline(i).Result = "BURBUJA (Ciclo " & Pipeline(i).StallCycles & ")"
            
            ' Si el stall ha durado lo suficiente, resolverlo
            If Pipeline(i).StallCycles >= GetRequiredStallCycles(i) Then
                Pipeline(i).Stalled = False
                Pipeline(i).StallCycles = 0
                Pipeline(i).Result = "Burbuja resuelta"
                LogMessage "? Burbuja resuelta para: " & Pipeline(i).instruction
            End If
        End If
    Next i
End Sub

Function GetRequiredStallCycles(stageIndex As Integer) As Integer
    ' Determinar cuántos ciclos de stall se necesitan para diferentes tipos de hazards
    Select Case stageIndex
        Case 0: ' IF stage
            GetRequiredStallCycles = 1
        Case 1: ' ID stage - RAW hazard más común
            GetRequiredStallCycles = 1
        Case 2: ' EX stage
            GetRequiredStallCycles = 1
        Case 3: ' MEM stage
            GetRequiredStallCycles = 1
        Case 4: ' WB stage
            GetRequiredStallCycles = 1
        Case Else
            GetRequiredStallCycles = 1
    End Select
End Function

' =============================================
' SIMULACIÓN DEL PIPELINE CON DETECCIÓN DE HAZARDS
' =============================================

Sub EjecutarPipelineCompleto()
    IsPipelineRunning = True
    Dim maxCycles As Integer
    maxCycles = 30
    
    Do While IsPipelineRunning And ClockCycle < maxCycles
        ClockCycle = ClockCycle + 1
        AdvancePipeline
        UpdatePipelineDisplay
        DoEvents
        
        If AllInstructionsCompleted() Then
            IsPipelineRunning = False
            LogMessage "? SIMULACIÓN COMPLETADA - Todas las instrucciones finalizadas"
            MsgBox "Simulación completada en " & ClockCycle & " ciclos" & vbCrLf & _
                   "Ciclos de stall: " & TotalStallCycles & vbCrLf & _
                   "Eficiencia: " & Format((ClockCycle - TotalStallCycles) / ClockCycle, "0.0%"), vbInformation
        End If
        
        ' Pequeña pausa para visualización
        Application.Wait (Now + TimeValue("0:00:00.5"))
    Loop
    
    If ClockCycle >= maxCycles Then
        MsgBox "Límite de ciclos alcanzado", vbInformation
    End If
    
    IsPipelineRunning = False
End Sub

Sub AvanzarCiclo()
    If IsPipelineRunning Then
        MsgBox "Detenga la simulación automática primero", vbExclamation
        Exit Sub
    End If
    
    ClockCycle = ClockCycle + 1
    AdvancePipeline
    UpdatePipelineDisplay
    
    If AllInstructionsCompleted() Then
        MsgBox "? Todas las instrucciones completadas en ciclo " & ClockCycle, vbInformation
    End If
End Sub

Sub AdvancePipeline()
    ' Primero resolver hazards existentes
    ResolveHazards
    
    ' Solo avanzar si no hay stalls activos en etapas críticas
    If Not HasActiveStalls() Then
        ' Avanzar las etapas en orden inverso (WB -> MEM -> EX -> ID -> IF)
        AdvanceStage "WB"
        AdvanceStage "MEM"
        AdvanceStage "EX"
        AdvanceStage "ID"
        AdvanceStage "IF"
        
        ' Insertar nueva instrucción en IF si hay disponible
        If CurrentInstructionIndex <= UBound(Instructions) And Not Pipeline(0).Stalled Then
            If Pipeline(0).stage = "" Then ' IF está vacío
                InsertNewInstruction Instructions(CurrentInstructionIndex)
                CurrentInstructionIndex = CurrentInstructionIndex + 1
            End If
        End If
    Else
        ' Log de ciclo de stall
        LogMessage "?? CICLO DE BURBUJA - Pipeline detenido por riesgos de datos"
    End If
    
    ' Verificar hazards después del avance
    CheckHazardsAfterAdvance
End Sub

Sub CheckHazardsAfterAdvance()
    ' Verificar hazards después de que las instrucciones han avanzado
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage = "ID" And Not Pipeline(i).Stalled Then
            Dim hazardMessage As String
            hazardMessage = CheckForDataHazards(i)
            If hazardMessage <> "" Then
                HandleDataHazard i, hazardMessage
            End If
        End If
    Next i
End Sub

Function HasActiveStalls() As Boolean
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).Stalled Then
            HasActiveStalls = True
            Exit Function
        End If
    Next i
    HasActiveStalls = False
End Function

Sub AdvanceStage(stage As String)
    Dim stageIndex As Integer
    stageIndex = GetStageIndex(stage)
    
    If stageIndex = -1 Then Exit Sub
    
    If Pipeline(stageIndex).stage = stage And Not Pipeline(stageIndex).Stalled Then
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
    
    ' Verificar si la siguiente etapa está vacía y no está en stall
    If Pipeline(nextStageIndex).stage = "" And Not Pipeline(nextStageIndex).Stalled Then
        CanAdvanceToNextStage = True
    Else
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
        
        ' Procesar la instrucción en la nueva etapa
        ProcessInstructionInStage nextStageIndex
    Else
        ' Instrucción completada (salida de WB)
        Pipeline(currentStageIndex).stage = "DONE"
        LogMessage "? Instrucción completada: " & Pipeline(currentStageIndex).instruction
    End If
    
    ' Limpiar etapa anterior
    ClearPipelineStage currentStageIndex
End Sub

Sub InsertNewInstruction(instruction As String)
    Pipeline(0).instruction = instruction
    Pipeline(0).stage = "IF"
    Pipeline(0).CycleEntered = ClockCycle
    Pipeline(0).CurrentStageCycle = 0
    Pipeline(0).Color = GetInstructionColor(CurrentInstructionIndex + 1)
    Pipeline(0).Stalled = False
    Pipeline(0).InstructionNumber = CurrentInstructionIndex + 1
    
    ' Procesar en etapa IF
    ProcessInstructionInStage 0
    LogMessage "?? Nueva instrucción: " & instruction
End Sub

Sub ProcessInstructionInStage(stageIndex As Integer)
    Select Case PipelineStages(stageIndex)
        Case "IF"
            Pipeline(stageIndex).Result = "Capturando instrucción"
            
        Case "ID"
            ParseInstruction stageIndex
            Pipeline(stageIndex).Result = "Decodificando: " & Pipeline(stageIndex).Opcode
            
        Case "EX"
            ExecuteInstruction stageIndex
            Pipeline(stageIndex).Result = "Ejecutando: " & Pipeline(stageIndex).Result
            
        Case "MEM"
            Pipeline(stageIndex).Result = "Acceso a memoria"
            
        Case "WB"
            Pipeline(stageIndex).Result = "Escritura de resultado"
    End Select
End Sub

Sub ParseInstruction(stageIndex As Integer)
    Dim instruction As String
    instruction = Pipeline(stageIndex).instruction
    
    ' Limpiar comentarios
    If InStr(instruction, ";") > 0 Then
        instruction = Trim(Left(instruction, InStr(instruction, ";") - 1))
    End If
    
    Dim parts() As String
    parts = Split(Trim(instruction), " ")
    
    If UBound(parts) >= 0 Then
        Pipeline(stageIndex).Opcode = UCase(Trim(parts(0)))
    End If
    
    ' Extraer operandos basado en el tipo de instrucción
    ExtractOperands stageIndex, parts
End Sub

Sub ExtractOperands(stageIndex As Integer, parts() As String)
    ' Reiniciar operandos
    Pipeline(stageIndex).DestinationReg = ""
    Pipeline(stageIndex).SourceReg1 = ""
    Pipeline(stageIndex).SourceReg2 = ""
    
    Select Case Pipeline(stageIndex).Opcode
        Case "MOV", "LOAD"
            If UBound(parts) >= 2 Then
                Pipeline(stageIndex).DestinationReg = ExtractRegister(parts(1))
                Pipeline(stageIndex).Operand1 = parts(2)
            End If
            
        Case "ADD", "SUB", "MUL", "DIV", "AND", "OR"
            If UBound(parts) >= 3 Then
                Pipeline(stageIndex).DestinationReg = ExtractRegister(parts(1))
                Pipeline(stageIndex).SourceReg1 = ExtractRegister(parts(2))
                Pipeline(stageIndex).SourceReg2 = ExtractRegister(parts(3))
            End If
            
        Case "STORE"
            If UBound(parts) >= 2 Then
                Pipeline(stageIndex).SourceReg1 = ExtractRegister(parts(1))
                Pipeline(stageIndex).Operand1 = parts(2)
            End If
    End Select
End Sub

Function ExtractRegister(operand As String) As String
    ' Extraer nombre de registro de un operando
    operand = Replace(operand, ",", "") ' Remover comas
    If UCase(Left(operand, 1)) = "R" Then
        ' Es un registro, extraer R0, R1, etc.
        ExtractRegister = UCase(Trim(operand))
    Else
        ExtractRegister = ""
    End If
End Function

Sub ExecuteInstruction(stageIndex As Integer)
    Dim op As String
    op = Pipeline(stageIndex).Opcode
    
    Select Case op
        Case "MOV", "LOAD"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).DestinationReg & " ? " & Pipeline(stageIndex).Operand1
        Case "ADD"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).SourceReg1 & " + " & Pipeline(stageIndex).SourceReg2
        Case "SUB"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).SourceReg1 & " - " & Pipeline(stageIndex).SourceReg2
        Case "MUL"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).SourceReg1 & " × " & Pipeline(stageIndex).SourceReg2
        Case "DIV"
            Pipeline(stageIndex).Result = Pipeline(stageIndex).SourceReg1 & " ÷ " & Pipeline(stageIndex).SourceReg2
        Case Else
            Pipeline(stageIndex).Result = "Operación: " & op
    End Select
End Sub

Function AllInstructionsCompleted() As Boolean
    ' Verificar si hay más instrucciones por cargar
    If CurrentInstructionIndex <= UBound(Instructions) Then
        AllInstructionsCompleted = False
        Exit Function
    End If
    
    ' Verificar si hay instrucciones en el pipeline
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" And Pipeline(i).stage <> "DONE" Then
            AllInstructionsCompleted = False
            Exit Function
        End If
    Next i
    
    AllInstructionsCompleted = True
End Function

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

' =============================================
' INTERFAZ UNIFICADA CON VISUALIZACIÓN DE RIESGOS
' =============================================

Sub CreateUnifiedPipelineDisplay()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PipelineRiesgos")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "PipelineRiesgos"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Tab.Color = RGB(70, 130, 180) ' Steel blue
    
    ' Título principal
    With ws.Range("A1:H1")
        .Merge
        .value = "?? PIPELINE CON DETECCIÓN DE RIESGOS DE DATOS"
        .Font.Bold = True
        .Font.size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(70, 130, 180)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
    End With
    
    ' Información del ciclo
    ws.Range("A3").value = "Ciclo Actual:"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").value = ClockCycle
    ws.Range("B3").Font.size = 14
    ws.Range("B3").Font.Bold = True
    
    ws.Range("D3").value = "Instrucciones:"
    ws.Range("D3").Font.Bold = True
    ws.Range("E3").value = GetInstructionCount()
    
    ws.Range("G3").value = "Burbujas:"
    ws.Range("G3").Font.Bold = True
    ws.Range("H3").value = TotalStallCycles
    ws.Range("H3").Interior.Color = RGB(255, 200, 200)
    
    ' Diagrama del pipeline
    ws.Range("A5").value = "?? DIAGRAMA DEL PIPELINE"
    ws.Range("A5").Font.Bold = True
    ws.Range("A5").Font.size = 12
    
    ' Encabezados de etapas
    Dim stages As Variant
    stages = Array("Inst#", "IF", "ID", "EX", "MEM", "WB", "Estado", "Riesgos")
    
    Dim col As Integer
    For col = 0 To 7
        With ws.Cells(7, col + 1)
            .value = stages(col)
            .Font.Bold = True
            .Font.size = 10
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(70, 130, 180)
            .HorizontalAlignment = xlCenter
            .Borders.Weight = xlThin
        End With
    Next col
    
    ' Área de visualización del pipeline
    Dim row As Integer
    For row = 8 To 12
        ws.ROWS(row).RowHeight = 35
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
    
    ' Área de riesgos detectados
    ws.Range("A14").value = "?? RIESGOS DETECTADOS"
    ws.Range("A14").Font.Bold = True
    ws.Range("A14").Font.size = 12
    
    ' Encabezados de riesgos
    Dim hazardHeaders As Variant
    hazardHeaders = Array("Ciclo", "Tipo", "Instrucción Afectada", "Instrucción en Conflicto", "Registros", "Acción")
    
    For col = 0 To 5
        With ws.Cells(15, col + 1)
            .value = hazardHeaders(col)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(178, 34, 34) ' Firebrick red
            .HorizontalAlignment = xlCenter
        End With
    Next col
    
    ' Área de log
    ws.Range("A25").value = "?? LOG DE EVENTOS"
    ws.Range("A25").Font.Bold = True
    ws.Range("A25").Font.size = 12
    
    ' Botones de control
    CreatePipelineButtons ws
    
    ' Ajustar columnas
    ws.Columns("A:A").ColumnWidth = 8
    ws.Columns("B:F").ColumnWidth = 15
    ws.Columns("G:G").ColumnWidth = 20
    ws.Columns("H:H").ColumnWidth = 15
End Sub

Sub CreatePipelineButtons(ws As Worksheet)
    ' Limpiar botones existentes
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    ' Botón Ejecutar Todo
    Dim btn As Button
    Set btn = ws.Buttons.Add(50, 400, 100, 30)
    btn.OnAction = "EjecutarPipelineCompleto"
    btn.Characters.Text = "?? Ejecutar"
    
    ' Botón Avanzar Ciclo
    Set btn = ws.Buttons.Add(160, 400, 100, 30)
    btn.OnAction = "AvanzarCiclo"
    btn.Characters.Text = "?? Avanzar"
    
    ' Botón Reiniciar
    Set btn = ws.Buttons.Add(270, 400, 80, 30)
    btn.OnAction = "ReiniciarPipeline"
    btn.Characters.Text = "?? Reiniciar"
End Sub

Sub UpdatePipelineDisplay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PipelineRiesgos")
    
    ' Actualizar información
    ws.Range("B3").value = ClockCycle
    ws.Range("E3").value = GetInstructionCount()
    ws.Range("H3").value = TotalStallCycles
    
    ' Limpiar área de visualización
    Dim row As Integer, col As Integer
    For row = 8 To 12
        For col = 1 To 8
            ws.Cells(row, col).value = ""
            ws.Cells(row, col).Interior.Color = RGB(240, 240, 240)
            ws.Cells(row, col).Font.Bold = False
        Next col
    Next row
    
    ' Mostrar instrucciones en cada etapa
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            ' Número de instrucción
            ws.Cells(8 + i, 1).value = "I" & Pipeline(i).InstructionNumber
            ws.Cells(8 + i, 1).Font.Bold = True
            
            ' Instrucción en etapa correspondiente
            Dim stageCol As Integer
            stageCol = GetStageColumn(Pipeline(i).stage)
            
            If stageCol > 0 Then
                ws.Cells(8 + i, stageCol).value = Pipeline(i).instruction
                ws.Cells(8 + i, stageCol).Interior.Color = Pipeline(i).Color
                ws.Cells(8 + i, stageCol).Font.Bold = True
                
                ' Estado
                ws.Cells(8 + i, 7).value = Pipeline(i).Result
                
                ' Mostrar riesgos
                If Pipeline(i).Stalled Then
                    ws.Cells(8 + i, 8).value = "?? BURBUJA ACTIVA"
                    ws.Cells(8 + i, 8).Interior.Color = RGB(255, 100, 100)
                    ws.Cells(8 + i, 8).Font.Bold = True
                    ' Resaltar toda la fila en rojo para stalls
                    For col = 1 To 8
                        ws.Cells(8 + i, col).Interior.Color = RGB(255, 200, 200)
                    Next col
                Else
                    ws.Cells(8 + i, 8).value = "? Sin riesgos"
                    ws.Cells(8 + i, 8).Interior.Color = RGB(200, 255, 200)
                End If
            End If
        End If
    Next i
End Sub

Function GetStageColumn(stage As String) As Integer
    Select Case stage
        Case "IF": GetStageColumn = 2
        Case "ID": GetStageColumn = 3
        Case "EX": GetStageColumn = 4
        Case "MEM": GetStageColumn = 5
        Case "WB": GetStageColumn = 6
        Case Else: GetStageColumn = 0
    End Select
End Function

Sub UpdateHazardDisplay(hazardMessage As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PipelineRiesgos")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    If nextRow < 16 Then nextRow = 16
    
    ws.Cells(nextRow, 1).value = ClockCycle
    ws.Cells(nextRow, 2).value = hazardMessage
    ws.Cells(nextRow, 3).value = Pipeline(1).instruction ' Instrucción en ID
    ws.Cells(nextRow, 4).value = GetAffectedInstruction(hazardMessage)
    ws.Cells(nextRow, 5).value = ExtractRegistersFromHazard(hazardMessage)
    ws.Cells(nextRow, 6).value = "INSERTAR BURBUJA"
    
    ' Resaltar fila en rojo para RAW hazards
    ws.ROWS(nextRow).Interior.Color = RGB(255, 200, 200)
    ws.ROWS(nextRow).Font.Bold = True
End Sub

Function ExtractRegistersFromHazard(hazardMessage As String) As String
    ' Extraer nombres de registros del mensaje de hazard
    If InStr(hazardMessage, "(") > 0 Then
        Dim regPart As String
        regPart = Mid(hazardMessage, InStr(hazardMessage, "(") + 1)
        regPart = Left(regPart, InStr(regPart, ")") - 1)
        ExtractRegistersFromHazard = regPart
    Else
        ExtractRegistersFromHazard = "N/A"
    End If
End Function

Function GetInstructionColor(instNum As Long) As Long
    Dim colors As Variant
    colors = Array( _
        RGB(173, 216, 230), _ ' Light Blue
        RGB(255, 182, 193), _ ' Light Pink
        RGB(221, 160, 221), _ ' Plum
        RGB(255, 218, 185), _ ' Peach
        RGB(176, 224, 230), _ ' Powder Blue
        RGB(240, 230, 140), _ ' Khaki
        RGB(152, 251, 152), _ ' Pale Green
        RGB(255, 228, 196)   ' Bisque
    )
    GetInstructionColor = colors((instNum - 1) Mod 8)
End Function

' =============================================
' LOGGING
' =============================================

Sub LogMessage(message As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PipelineRiesgos")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    If nextRow < 26 Then nextRow = 26
    
    ws.Cells(nextRow, 1).value = ClockCycle
    ws.Cells(nextRow, 2).value = message
    
    ' Color según tipo de mensaje
    If InStr(message, "?") > 0 Then
        ws.ROWS(nextRow).Interior.Color = RGB(200, 255, 200)
    ElseIf InStr(message, "??") > 0 Or InStr(message, "BURBUJA") > 0 Then
        ws.ROWS(nextRow).Interior.Color = RGB(255, 200, 200)
    ElseIf InStr(message, "??") > 0 Then
        ws.ROWS(nextRow).Interior.Color = RGB(200, 220, 255)
    Else
        ws.ROWS(nextRow).Interior.Color = RGB(240, 240, 240)
    End If
End Sub

Sub LogHazard(hazardMessage As String, currentInstr As String, conflictInstr As String)
    LogMessage "?? RIESGO: " & hazardMessage & " | " & currentInstr & " espera por " & conflictInstr
End Sub


