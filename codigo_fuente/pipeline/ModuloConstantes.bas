Attribute VB_Name = "ModuloConstantes"
Attribute VB_Name = "ModuloConstantes"
Option Explicit

' Constantes del sistema
Public Const MEM_SIZE As Long = 256
Public Const CACHE_LINES As Long = 8

' Variables globales de registros
Public EAX As Long
Public EBX As Long
Public ECX As Long
Public EDX As Long
Public EIP As Long

' Flags
Public ZeroFlag As Boolean
Public CarryFlag As Boolean
Public SignFlag As Boolean

Public Sub InicializarRegistros()
    EAX = 0: EBX = 0: ECX = 0: EDX = 0: EIP = 0
    ZeroFlag = False: CarryFlag = False: SignFlag = False
End Sub

Public Function ObtenerValorRegistro(nombre As String) As Long
    Select Case UCase(Trim(nombre))
        Case "EAX": ObtenerValorRegistro = EAX
        Case "EBX": ObtenerValorRegistro = EBX
        Case "ECX": ObtenerValorRegistro = ECX
        Case "EDX": ObtenerValorRegistro = EDX
        Case Else: ObtenerValorRegistro = 0
    End Select
End Function

Public Sub ActualizarRegistro(nombre As String, valor As Long)
    Select Case UCase(Trim(nombre))
        Case "EAX": EAX = valor
        Case "EBX": EBX = valor
        Case "ECX": ECX = valor
        Case "EDX": EDX = valor
    End Select
    ZeroFlag = (valor = 0)
    SignFlag = (valor < 0)
End Sub

Public Function ObtenerValorOperando(operando As String) As Long
    If IsNumeric(operando) Then
        ObtenerValorOperando = CLng(operando)
    Else
        ObtenerValorOperando = ObtenerValorRegistro(operando)
    End If
End Function

Public Sub ParsearInstruccionSimple(instruccion As String, ByRef opcode As String, ByRef destino As String, ByRef fuente As String)
    Dim partes() As String
    
    If Len(Trim(instruccion)) = 0 Or UCase(instruccion) = "NOP" Then
        opcode = "NOP": destino = "": fuente = ""
        Exit Sub
    End If
    
    partes = Split(Trim(instruccion), " ")
    opcode = UCase(Trim(partes(0)))
    
    If UBound(partes) >= 1 Then
        Dim operandos As String
        operandos = Trim(partes(1))
        
        If InStr(operandos, ",") > 0 Then
            Dim ops() As String
            ops = Split(operandos, ",")
            destino = Trim(ops(0))
            If UBound(ops) >= 1 Then fuente = Trim(ops(1))
        Else
            destino = operandos
            fuente = ""
        End If
    End If
End Sub

Public Sub ParsearInstruccion(instruccion As String)
    Dim opcode As String, destino As String, fuente As String
    ParsearInstruccionSimple instruccion, opcode, destino, fuente
    
    Select Case opcode
        Case "MOV": ActualizarRegistro destino, ObtenerValorOperando(fuente)
        Case "ADD": ActualizarRegistro destino, ObtenerValorRegistro(destino) + ObtenerValorOperando(fuente)
        Case "SUB": ActualizarRegistro destino, ObtenerValorRegistro(destino) - ObtenerValorOperando(fuente)
        Case "INC": ActualizarRegistro destino, ObtenerValorRegistro(destino) + 1
        Case "DEC": ActualizarRegistro destino, ObtenerValorRegistro(destino) - 1
        Case "JMP": EIP = CLng(destino) - 1
    End Select
End Sub
4
