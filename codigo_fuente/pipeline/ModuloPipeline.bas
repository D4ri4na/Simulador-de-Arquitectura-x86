Attribute VB_Name = "ModuloPipeline"
Attribute VB_Name = "ModuloPipeline"
Option Explicit

Public Type IF_ID_Register
    Instruccion As String
    PC As Long
End Type
Public Type ID_EX_Register
    Instruccion As String
    opcode As String
    Destino As String
    Fuente As String
    ControlSignal As String
End Type
Public Type EX_MEM_Register
    Instruccion As String
    ResultadoALU As Long
    Destino As String
    ControlSignal As String
End Type
Public Type MEM_WB_Register
    Instruccion As String
    ResultadoALU As Long
    Destino As String
    ControlSignal As String
End Type

Public IF_ID As IF_ID_Register
Public ID_EX As ID_EX_Register
Public EX_MEM As EX_MEM_Register
Public MEM_WB As MEM_WB_Register
Public stall As Boolean

Public Sub LimpiarPipeline()
    IF_ID.Instruccion = "NOP": IF_ID.PC = 0
    ID_EX.Instruccion = "NOP": ID_EX.opcode = "NOP"
    EX_MEM.Instruccion = "NOP"
    MEM_WB.Instruccion = "NOP"
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
        IF_ID.Instruccion = LeerDesdeMemoria(EIP)
        IF_ID.PC = EIP
        Dim opcode As String: Dim d As String: Dim f As String
        ParsearInstruccionSimple IF_ID.Instruccion, opcode, d, f
        If opcode <> "JMP" Then EIP = EIP + 1
    End If
End Sub

Private Sub Etapa_ID()
    If Not stall Then
        ID_EX.Instruccion = IF_ID.Instruccion
        ParsearInstruccionSimple IF_ID.Instruccion, ID_EX.opcode, ID_EX.Destino, ID_EX.Fuente
    Else
        ID_EX.Instruccion = "NOP": ID_EX.opcode = "NOP"
        ID_EX.Destino = "": ID_EX.Fuente = ""
    End If
    
    Dim fuenteNecesaria As String
    ParsearInstruccionSimple ID_EX.Instruccion, ID_EX.opcode, ID_EX.Destino, fuenteNecesaria

    If (ID_EX.opcode = "ADD" Or ID_EX.opcode = "SUB") And _
       (EX_MEM.Destino = ID_EX.Destino Or EX_MEM.Destino = ID_EX.Fuente) And _
       (EX_MEM.ControlSignal = "ALU") Then
        stall = True
    Else
        stall = False
    End If
End Sub

Private Sub Etapa_EX()
    EX_MEM.Instruccion = ID_EX.Instruccion
    EX_MEM.Destino = ID_EX.Destino
    
    Select Case ID_EX.opcode
        Case "MOV"
            EX_MEM.ResultadoALU = ObtenerValorOperando(ID_EX.Fuente)
            EX_MEM.ControlSignal = "ALU"
        Case "ADD"
            EX_MEM.ResultadoALU = ObtenerValorRegistro(ID_EX.Destino) + ObtenerValorOperando(ID_EX.Fuente)
            EX_MEM.ControlSignal = "ALU"
        Case "SUB"
            EX_MEM.ResultadoALU = ObtenerValorRegistro(ID_EX.Destino) - ObtenerValorOperando(ID_EX.Fuente)
            EX_MEM.ControlSignal = "ALU"
        Case "JMP"
            EIP = CLng(ID_EX.Destino)
            IF_ID.Instruccion = "NOP"
            ID_EX.Instruccion = "NOP": ID_EX.opcode = "NOP"
            EX_MEM.ControlSignal = "JUMP"
        Case Else
            EX_MEM.ControlSignal = "NONE"
    End Select
End Sub

Private Sub Etapa_MEM()
    MEM_WB.Instruccion = EX_MEM.Instruccion
    MEM_WB.ResultadoALU = EX_MEM.ResultadoALU
    MEM_WB.Destino = EX_MEM.Destino
    MEM_WB.ControlSignal = EX_MEM.ControlSignal
End Sub

Private Sub Etapa_WB()
    If MEM_WB.Destino <> "" And MEM_WB.ControlSignal = "ALU" Then
        ActualizarRegistro MEM_WB.Destino, MEM_WB.ResultadoALU
    End If
End Sub

