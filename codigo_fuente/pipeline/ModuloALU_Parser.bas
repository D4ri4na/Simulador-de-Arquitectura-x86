Attribute VB_Name = "ModuloALU_Parser"
Attribute VB_Name = "ModuloALU_Parser"
Option Explicit

' --- Obtiene el valor de un registro por su nombre ---
Public Function ObtenerValorRegistro(nombre As String) As Long
    Select Case UCase(Trim(nombre))
        Case "EAX": ObtenerValorRegistro = EAX
        Case "EBX": ObtenerValorRegistro = EBX
        Case "ECX": ObtenerValorRegistro = ECX
        Case "EDX": ObtenerValorRegistro = EDX
        Case Else: ObtenerValorRegistro = 0
    End Select
End Function

' --- Actualiza el valor de un registro por su nombre ---
Public Sub ActualizarRegistro(nombre As String, valor As Long)
    Select Case UCase(Trim(nombre))
        Case "EAX": EAX = valor
        Case "EBX": EBX = valor
        Case "ECX": ECX = valor
        Case "EDX": EDX = valor
    End Select
End Sub

' --- Obtiene el valor de un operando (puede ser un registro o un número) ---
Public Function ObtenerValorOperando(operando As String) As Long
    If IsNumeric(operando) Then
        ObtenerValorOperando = CLng(operando)
    Else
        ObtenerValorOperando = ObtenerValorRegistro(operando)
    End If
End Function

' --- Parsea una instrucción en sus componentes: opcode, destino y fuente ---
Public Sub ParsearInstruccionSimple(ByVal Instruccion As String, ByRef opcode As String, ByRef Destino As String, ByRef Fuente As String)
    Dim partes() As String
    opcode = "": Destino = "": Fuente = ""

    If Len(Trim(Instruccion)) = 0 Or UCase(Instruccion) = "NOP" Then
        opcode = "NOP"
        Exit Sub
    End If

    partes = Split(Trim(Instruccion), " ")
    opcode = UCase(Trim(partes(0)))

    If UBound(partes) >= 1 Then
        Dim operandos As String
        operandos = Trim(partes(1))
        
        If InStr(operandos, ",") > 0 Then
            Dim ops() As String
            ops = Split(operandos, ",")
            Destino = Trim(ops(0))
            If UBound(ops) >= 1 Then Fuente = Trim(ops(1))
        Else
            Destino = operandos
        End If
    End If
End Sub
