Attribute VB_Name = "ModuloPipeline"
Attribute VB_Name = "ModuloPipeline"
Option Explicit

Public Type IF_ID_Register
    instruccion As String
    PC As Long
End Type

Public Type ID_EX_Register
    instruccion As String
    opcode As String
    destino As String
    fuente As String
    ControlSignal As String
End Type

Public Type EX_MEM_Register
    instruccion As String
    ResultadoALU As Long
    destino As String
    ControlSignal As String
End Type

Public Type MEM_WB_Register
    instruccion As String
    DatoLeido As String
    ResultadoALU As Long
    destino As String
    ControlSignal As String
End Type

Public IF_ID As IF_ID_Register
Public ID_EX As ID_EX_Register
Public EX_MEM As EX_MEM_Register
Public MEM_WB As MEM_WB_Register
Public stall As Boolean

Public Sub LimpiarPipeline()
    IF_ID.instruccion = "NOP": IF_ID.PC = 0
    ID_EX.instruccion = "NOP": ID_EX.opcode = "NOP"
    EX_MEM.instruccion = "NOP"
    MEM_WB.instruccion = "NOP"
    stall = False
End Sub

Public Sub AvanzarCicloPipeline()
    Etapa_WB
    Etapa_MEM
    Etapa_EX
    Etapa_ID
    Etapa_IF
End Sub

Private Sub Etapa_IF()
    If Not stall Then
        IF_ID.instruccion = LeerDesdeMemoria(EIP)
        IF_ID.PC = EIP
        
        Dim opcode As String
        If Len(IF_ID.instruccion) > 0 Then
            opcode = UCase(Split(IF_ID.instruccion & " ", " ")(0))
            If opcode <> "JMP" Then EIP = EIP + 1
        End If
    End If
End Sub

Private Sub Etapa_ID()
    If Not stall Then
        ID_EX.instruccion = IF_ID.instruccion
        ParsearInstruccionSimple IF_ID.instruccion, ID_EX.opcode, ID_EX.destino, ID_EX.fuente
    Else
        ID_EX.instruccion = "NOP"
        ID_EX.opcode = "NOP"
        ID_EX.destino = ""
        ID_EX.fuente = ""
    End If
    
    ' Detección de riesgos
    Dim fuenteNecesaria As String
    fuenteNecesaria = ID_EX.fuente
    
    If (EX_MEM.ControlSignal = "ALU" Or EX_MEM.ControlSignal = "MEM_READ") And _
       (EX_MEM.destino = fuenteNecesaria) And (fuenteNecesaria <> "") Then
        stall = True
    Else
        stall = False
    End If
End Sub

Private Sub Etapa_EX()
    EX_MEM.instruccion = ID_EX.instruccion
    EX_MEM.destino = ID_EX.destino
    
    Select Case ID_EX.opcode
        Case "MOV"
            EX_MEM.ResultadoALU = ObtenerValorOperando(ID_EX.fuente)
            EX_MEM.ControlSignal = "ALU"
        Case "ADD"
            EX_MEM.ResultadoALU = ObtenerValorRegistro(ID_EX.destino) + ObtenerValorOperando(ID_EX.fuente)
            EX_MEM.ControlSignal = "ALU"
        Case "SUB"
            EX_MEM.ResultadoALU = ObtenerValorRegistro(ID_EX.destino) - ObtenerValorOperando(ID_EX.fuente)
            EX_MEM.ControlSignal = "ALU"
        Case "JMP"
            EIP = CLng(ID_EX.destino)
            IF_ID.instruccion = "NOP"
            ID_EX.instruccion = "NOP": ID_EX.opcode = "NOP"
            EX_MEM.ControlSignal = "JUMP"
        Case Else
            EX_MEM.ControlSignal = "NONE"
    End Select
End Sub

Private Sub Etapa_MEM()
    MEM_WB.instruccion = EX_MEM.instruccion
    MEM_WB.ResultadoALU = EX_MEM.ResultadoALU
    MEM_WB.destino = EX_MEM.destino
    MEM_WB.ControlSignal = EX_MEM.ControlSignal
End Sub

Private Sub Etapa_WB()
    If MEM_WB.destino <> "" And (MEM_WB.ControlSignal = "ALU" Or MEM_WB.ControlSignal = "MEM_READ") Then
        ActualizarRegistro MEM_WB.destino, MEM_WB.ResultadoALU
    End If
End Sub

