Attribute VB_Name = "Deteccion"
' Simulador de Pipeline con Detección de Riesgos de Datos
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
End Type

' Variables globales del pipeline
Dim Pipeline() As PipelineInstruction
Dim ClockCycle As Long
Dim Instructions() As String
Dim CurrentInstructionIndex As Long
Dim PipelineStages(4) As String
Dim IsPipelineRunning As Boolean
Dim RegisterStatus(15) As String ' Estado de registros R0-R15
Dim StallBuffer As Collection

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
        ClearPipelineStage i
    Next i
    
    ' Inicializar estado de registros
    For i = 0 To 15
        RegisterStatus(i) = "READY"
    Next i
    
    ClockCycle = 0
    CurrentInstructionIndex = 0
    Set StallBuffer = New Collection
    
    ' Cargar instrucciones desde hoja
    LoadInstructionsFromSheet
    
    CreatePipelineDisplay
    UpdatePipelineDisplay
    CreatePipelineLog
    CreateHazardDisplay
    
    MsgBox "Pipeline con detección de riesgos inicializado con " & GetInstructionCount() & " instrucciones", vbInformation
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
End Sub

' =============================================
' DETECCIÓN DE RIESGOS DE DATOS
' =============================================

Function CheckForDataHazards(currentStageIndex As Integer) As String
    ' Esta función verifica riesgos de datos y devuelve el tipo de hazard detectado
    If currentStageIndex <> 1 Then ' Solo verificamos en etapa ID
        CheckForDataHazards = ""
        Exit Function
    End If
    
    Dim currentInstr As PipelineInstruction
    currentInstr = Pipeline(currentStageIndex)
    
    ' Verificar contra instrucciones en etapas EX, MEM, WB
    Dim hazardType As String
    hazardType = ""
    
    ' Verificar dependencias con instrucción en EX
    If Pipeline(2).stage = "EX" Then
        hazardType = CheckDependency(currentInstr, Pipeline(2), "EX")
        If hazardType <> "" Then
            CheckForDataHazards = "RAW-EX: " & hazardType
            Exit Function
        End If
    End If
    
    ' Verificar dependencias con instrucción en MEM
    If Pipeline(3).stage = "MEM" Then
        hazardType = CheckDependency(currentInstr, Pipeline(3), "MEM")
        If hazardType <> "" Then
            CheckForDataHazards = "RAW-MEM: " & hazardType
            Exit Function
        End If
    End If
    
    ' Verificar dependencias con instrucción en WB
    If Pipeline(4).stage = "WB" Then
        hazardType = CheckDependency(currentInstr, Pipeline(4), "WB")
        If hazardType <> "" Then
            CheckForDataHazards = "RAW-WB: " & hazardType
            Exit Function
        End If
    End If
    
    CheckForDataHazards = ""
End Function

Function CheckDependency(currentInstr As PipelineInstruction, previousInstr As PipelineInstruction, stage As String) As String
    ' Verificar dependencias RAW (Read After Write)
    Dim dependency As String
    
    ' Verificar si la instrucción actual lee un registro que la anterior escribe
    If previousInstr.DestinationReg <> "" Then
        ' Verificar primer operando fuente
        If currentInstr.SourceReg1 <> "" And currentInstr.SourceReg1 = previousInstr.DestinationReg Then
            dependency = currentInstr.SourceReg1 & " (Op1)"
        End If
        
        ' Verificar segundo operando fuente
        If currentInstr.SourceReg2 <> "" And currentInstr.SourceReg2 = previousInstr.DestinationReg Then
            If dependency <> "" Then dependency = dependency & ", "
            dependency = dependency & currentInstr.SourceReg2 & " (Op2)"
        End If
    End If
    
    CheckDependency = dependency
End Function

Sub HandleDataHazard(stageIndex As Integer, hazardMessage As String)
    ' Insertar burbuja (stall) en el pipeline
    Pipeline(stageIndex).Stalled = True
    Pipeline(stageIndex).StallCycles = Pipeline(stageIndex).StallCycles + 1
    
    ' Crear burbuja en IF si es necesario
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
    ElseIf InStr(hazardMessage, "WB") > 0 And Pipeline(4).stage = "WB" Then
        GetAffectedInstruction = Pipeline(4).instruction
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
            Pipeline(i).Result = "STALL (Ciclo " & Pipeline(i).StallCycles & ")"
            
            ' Si el stall ha durado lo suficiente, resolverlo
            If Pipeline(i).StallCycles >= GetRequiredStallCycles(i) Then
                Pipeline(i).Stalled = False
                Pipeline(i).StallCycles = 0
                Pipeline(i).Result = "Stall resuelto"
                LogMessage "Stall resuelto para: " & Pipeline(i).instruction
            End If
        End If
    Next i
End Sub

Function GetRequiredStallCycles(stageIndex As Integer) As Integer
    ' Determinar cuántos ciclos de stall se necesitan
    Select Case stageIndex
        Case 1: ' ID stage - RAW hazard
            GetRequiredStallCycles = 1
        Case Else
            GetRequiredStallCycles = 1
    End Select
End Function

' =============================================
' SIMULACIÓN DEL PIPELINE CON DETECCIÓN DE HAZARDS
' =============================================

Sub AdvancePipeline()
    ' Primero resolver hazards existentes
    ResolveHazards
    
    ' Solo avanzar si no hay stalls activos
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
        LogMessage "CICLO DE STALL - Pipeline detenido por hazards"
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
    
    ' Verificar si la siguiente etapa está ocupada o en stall
    If (Pipeline(nextStageIndex).stage = "" Or Pipeline(nextStageIndex).stage = "DONE") And _
       Not Pipeline(nextStageIndex).Stalled Then
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
        ' Liberar registro de destino
        If Pipeline(currentStageIndex).DestinationReg <> "" Then
            FreeRegister Pipeline(currentStageIndex).DestinationReg
        End If
    End If
    
    ' Limpiar etapa anterior
    ClearPipelineStage currentStageIndex
End Sub

Sub InsertNewInstruction(instruction As String)
    Pipeline(0).instruction = instruction
    Pipeline(0).stage = "IF"
    Pipeline(0).CycleEntered = ClockCycle
    Pipeline(0).CurrentStageCycle = 0
    Pipeline(0).Color = GetRandomColor()
    Pipeline(0).Stalled = False
    
    ' Procesar en etapa IF
    ProcessInstructionInStage 0
End Sub

Sub ProcessInstructionInStage(stageIndex As Integer)
    Dim instruction As String
    instruction = Pipeline(stageIndex).instruction
    
    Select Case PipelineStages(stageIndex)
        Case "IF"
            Pipeline(stageIndex).Result = "Instrucción capturada"
            
        Case "ID"
            ParseInstruction stageIndex
            Pipeline(stageIndex).Result = "Decodificada: " & Pipeline(stageIndex).Opcode
            
            ' Reservar registro de destino si existe
            If Pipeline(stageIndex).DestinationReg <> "" Then
                ReserveRegister Pipeline(stageIndex).DestinationReg, Pipeline(stageIndex).instruction
            End If
            
        Case "EX"
            ExecuteInstruction stageIndex
            Pipeline(stageIndex).Result = "Ejecutado: " & Pipeline(stageIndex).Result
            
        Case "MEM"
            Pipeline(stageIndex).Result = "Acceso MEM completado"
            
        Case "WB"
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
    
    ' Extraer operandos basado en el tipo de instrucción
    ExtractOperands stageIndex, parts
End Sub

Sub ExtractOperands(stageIndex As Integer, parts() As String)
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
    If Left(operand, 1) = "R" Or Left(operand, 1) = "r" Then
        ' Es un registro, extraer R0, R1, etc.
        Dim regPart As String
        regPart = Split(operand, ",")(0)
        ExtractRegister = UCase(Trim(regPart))
    Else
        ExtractRegister = ""
    End If
End Function

Sub ReserveRegister(reg As String, instruction As String)
    ' Marcar registro como en uso
    If reg <> "" Then
        Dim regIndex As Integer
        regIndex = GetRegisterIndex(reg)
        If regIndex >= 0 Then
            RegisterStatus(regIndex) = instruction
        End If
    End If
End Sub

Sub FreeRegister(reg As String)
    ' Liberar registro
    If reg <> "" Then
        Dim regIndex As Integer
        regIndex = GetRegisterIndex(reg)
        If regIndex >= 0 Then
            RegisterStatus(regIndex) = "READY"
        End If
    End If
End Sub

Function GetRegisterIndex(reg As String) As Integer
    ' Convertir nombre de registro a índice
    If Len(reg) >= 2 Then
        GetRegisterIndex = Val(Mid(reg, 2))
    Else
        GetRegisterIndex = -1
    End If
End Function

' =============================================
' VISUALIZACIÓN MEJORADA CON HAZARDS
' =============================================

Sub CreateHazardDisplay()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Riesgos")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Riesgos"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ' Título
    ws.Cells(1, 1).value = "DETECCIÓN DE RIESGOS DE DATOS"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.size = 14
    ws.Range("A1:D1").Merge
    
    ' Encabezados
    ws.Cells(3, 1).value = "Ciclo"
    ws.Cells(3, 2).value = "Tipo de Riesgo"
    ws.Cells(3, 3).value = "Instrucción Afectada"
    ws.Cells(3, 4).value = "Instrucción en Conflicto"
    ws.Cells(3, 5).value = "Registros Involucrados"
    ws.Cells(3, 6).value = "Acción"
    
    ' Formato de encabezados
    Dim headerRange As Range
    Set headerRange = ws.Range("A3:F3")
    headerRange.Font.Bold = True
    headerRange.Interior.Color = RGB(180, 180, 180)
    
    ws.Columns.AutoFit
End Sub

Sub UpdateHazardDisplay(hazardMessage As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Riesgos")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = ClockCycle
    ws.Cells(nextRow, 2).value = hazardMessage
    ws.Cells(nextRow, 3).value = Pipeline(1).instruction ' Instrucción en ID
    ws.Cells(nextRow, 4).value = GetAffectedInstruction(hazardMessage)
    ws.Cells(nextRow, 5).value = ExtractRegistersFromHazard(hazardMessage)
    ws.Cells(nextRow, 6).value = "INSERTAR BURBUJA"
    
    ' Resaltar fila según tipo de hazard
    If InStr(hazardMessage, "RAW") > 0 Then
        ws.ROWS(nextRow).Interior.Color = RGB(255, 200, 200) ' Rojo para RAW
    ElseIf InStr(hazardMessage, "WAR") > 0 Then
        ws.ROWS(nextRow).Interior.Color = RGB(255, 255, 200) ' Amarillo para WAR
    ElseIf InStr(hazardMessage, "WAW") > 0 Then
        ws.ROWS(nextRow).Interior.Color = RGB(200, 255, 255) ' Azul para WAW
    End If
    
    ws.Columns.AutoFit
End Sub

Function ExtractRegistersFromHazard(hazardMessage As String) As String
    ' Extraer nombres de registros del mensaje de hazard
    Dim regPart As String
    If InStr(hazardMessage, "(") > 0 Then
        regPart = Mid(hazardMessage, InStr(hazardMessage, "("))
        regPart = Replace(regPart, "(", "")
        regPart = Replace(regPart, ")", "")
        ExtractRegistersFromHazard = regPart
    Else
        ExtractRegistersFromHazard = "N/A"
    End If
End Function

Sub UpdatePipelineDisplay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Pipeline")
    
    ' Actualizar ciclo actual
    ws.Cells(2, 2).value = ClockCycle
    
    ' Mostrar instrucciones en cada etapa
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).stage <> "" Then
            ' Instrucción con color
            ws.Cells(7 + i, 2 + i).value = Pipeline(i).instruction
            ws.Cells(7 + i, 2 + i).Interior.Color = Pipeline(i).Color
            
            ' Información adicional
            ws.Cells(7 + i, 1).value = "Inst " & (Pipeline(i).CycleEntered + 1)
            ws.Cells(8 + i, 2 + i).value = Pipeline(i).Result
            
            ' Resaltar stalls y hazards
            If Pipeline(i).Stalled Then
                ws.Cells(7 + i, 2 + i).Interior.Color = RGB(255, 100, 100) ' Rojo para stall
                ws.Cells(7 + i, 2 + i).value = Pipeline(i).instruction & " [BURBUJA]"
                ws.Cells(8 + i, 2 + i).value = "STALL por riesgo de datos"
            End If
        Else
            ' Limpiar celda si no hay instrucción
            ws.Cells(7 + i, 2 + i).value = ""
            ws.Cells(7 + i, 2 + i).Interior.ColorIndex = 0
            ws.Cells(8 + i, 2 + i).value = ""
        End If
    Next i
    
    ' Mostrar burbujas explícitamente
    ShowBubbles ws
    
    ws.Columns.AutoFit
End Sub

Sub ShowBubbles(ws As Worksheet)
    ' Mostrar burbujas en el diagrama del pipeline
    Dim bubbleRow As Long
    bubbleRow = 12
    
    ws.Cells(bubbleRow, 1).value = "BURBUJAS ACTIVAS:"
    ws.Cells(bubbleRow, 1).Font.Bold = True
    bubbleRow = bubbleRow + 1
    
    Dim i As Integer
    For i = 0 To 4
        If Pipeline(i).Stalled Then
            ws.Cells(bubbleRow, 1).value = "Etapa " & PipelineStages(i) & ": " & Pipeline(i).instruction
            ws.Cells(bubbleRow, 2).value = "Ciclos en stall: " & Pipeline(i).StallCycles
            ws.Cells(bubbleRow, 1).Interior.Color = RGB(255, 200, 200)
            bubbleRow = bubbleRow + 1
        End If
    Next i
End Sub

Sub LogHazard(hazardMessage As String, currentInstr As String, conflictInstr As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LogPipeline")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = ClockCycle
    ws.Cells(nextRow, 7).value = "RIESGO: " & hazardMessage & " | " & currentInstr & " espera por " & conflictInstr
    ws.Cells(nextRow, 7).Font.Bold = True
    ws.Cells(nextRow, 7).Interior.Color = RGB(255, 200, 200)
    
    ws.Columns.AutoFit
End Sub

' =============================================
' FUNCIONES AUXILIARES (mantener las del código anterior)
' =============================================

Function GetRandomColor() As Long
    Dim r As Integer, g As Integer, b As Integer
    r = Int((200 - 150 + 1) * Rnd + 150)
    g = Int((200 - 150 + 1) * Rnd + 150)
    b = Int((200 - 150 + 1) * Rnd + 150)
    GetRandomColor = RGB(r, g, b)
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
' EJEMPLO CON RIESGOS DE DATOS
' =============================================

Sub CrearEjemploConRiesgos()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CodigoPipeline")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "CodigoPipeline"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Cells(1, 1).value = "Ejemplo con Riesgos de Datos (RAW)"
    ws.Cells(1, 1).Font.Bold = True
    
    ' Instrucciones que causarán riesgos RAW
    ws.Cells(3, 1).value = "ADD R1, R2, R3"    ' Escribe R1
    ws.Cells(4, 1).value = "SUB R4, R1, R5"    ' Lee R1 - RAW hazard!
    ws.Cells(5, 1).value = "MUL R6, R7, R8"    ' Independiente
    ws.Cells(6, 1).value = "DIV R9, R1, R10"   ' Lee R1 - Otro RAW hazard!
    ws.Cells(7, 1).value = "MOV R11, R12"      ' Independiente
    ws.Cells(8, 1).value = "ADD R13, R4, R1"   ' Lee R1 y R4 - Múltiples hazards!
    
    ws.Columns.AutoFit
    MsgBox "Ejemplo con riesgos de datos creado." & vbCrLf & _
           "Observe cómo se detectan y resuelven los hazards RAW automáticamente.", vbInformation
End Sub

Sub IniciarSimuladorConRiesgos()
    InitializePipeline
    MsgBox "Simulador de Pipeline con Detección de Riesgos inicializado." & vbCrLf & _
           "Los riesgos RAW se detectarán automáticamente y se insertarán burbujas.", vbInformation
End Sub

