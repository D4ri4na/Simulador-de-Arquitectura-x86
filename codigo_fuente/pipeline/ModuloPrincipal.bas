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
    ThisWorkbook.Sheets("Simulador").Calculate
    On Error GoTo 0
End Sub

Public Sub CargarProgramaEnMemoria()
    Dim i As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Simulador")
    
    InicializarMemoria
    
    i = 0
    Dim cell As Range
    For Each cell In ws.Range("CodigoFuente").Cells
        If Trim(cell.Value) <> "" Then
            If i < MEM_SIZE Then
                RAM(i) = Trim(cell.Value)
                i = i + 1
            End If
        End If
    Next cell
    
    EIP = 0
    LimpiarPipeline
    ws.Calculate
    MsgBox "Programa cargado: " & i & " instrucciones.", vbInformation
End Sub

Public Sub EjecutarUnCiclo()
    Clock = Clock + 1
    AvanzarCicloPipeline
    On Error Resume Next
    ThisWorkbook.Sheets("Simulador").Calculate
    On Error GoTo 0
End Sub
