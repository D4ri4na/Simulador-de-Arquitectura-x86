Attribute VB_Name = "ModuloPrincipal"
Attribute VB_Name = "ModuloPrincipal"
Option Explicit

Public Clock As Long

Public Sub InicializarSimulador()
    Clock = 0
    InicializarRegistros
    InicializarMemoria
    LimpiarPipeline
    
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Simulador")
    If Not ws Is Nothing Then
        ActualizarVisualizacion
    End If
    On Error GoTo 0
    
    MsgBox "Simulador Inicializado. Escribe tu código y presiona 'Cargar Programa'", vbInformation
End Sub

Public Sub CargarProgramaEnMemoria()
    Dim i As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Simulador")
    
    For i = 0 To 255: RAM(i) = "NOP": Next i
    
    i = 0
    Dim cell As Range
    For Each cell In ws.Range("CodigoFuente").Cells
        If Trim(cell.value) <> "" Then
            RAM(i) = Trim(cell.value)
            i = i + 1
        End If
    Next cell
    
    EIP = 0
    LimpiarPipeline
    ActualizarVisualizacion
    
    MsgBox "Programa cargado: " & i & " instrucciones", vbInformation
End Sub

Public Sub EjecutarUnCiclo()
    Clock = Clock + 1
    AvanzarCicloPipeline
    ActualizarVisualizacion
End Sub

Public Sub EjecutarTodo()
    Dim maxCiclos As Long: maxCiclos = 100
    Do While EIP < 256 And Clock < maxCiclos
        If Trim(RAM(EIP)) = "" Or RAM(EIP) = "NOP" Then Exit Do
        EjecutarUnCiclo
        DoEvents
    Loop
    MsgBox "Ejecución completada. Ciclos: " & Clock, vbInformation
End Sub

