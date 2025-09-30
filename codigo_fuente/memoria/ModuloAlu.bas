Attribute VB_Name = "ModuloAlu"
' ModuloALU.bas - VERSI�N COMPLETA
Option Explicit


' Funci�n para obtener valor de un registro por nombre
Private Function ObtenerValorRegistro(nombreRegistro As String) As Long
    Select Case UCase(nombreRegistro)
        Case "EAX": ObtenerValorRegistro = EAX
        Case "EBX": ObtenerValorRegistro = EBX
        Case "ECX": ObtenerValorRegistro = ECX
        Case "EDX": ObtenerValorRegistro = EDX
        Case "ESI": ObtenerValorRegistro = ESI
        Case "EDI": ObtenerValorRegistro = EDI
        Case "EBP": ObtenerValorRegistro = EBP
        Case "ESP": ObtenerValorRegistro = ESP
        Case Else: ObtenerValorRegistro = 0
    End Select
End Function

' Funci�n para obtener valor de un operando (puede ser registro o n�mero)
Private Function ObtenerValorOperando(operando As String) As Long
    ' Si es un n�mero (hexadecimal o decimal)
    If EsNumero(operando) Then
        ObtenerValorOperando = ConvertirANumero(operando)
    Else
        ' Si es un registro
        ObtenerValorOperando = ObtenerValorRegistro(operando)
    End If
End Function

' Funci�n para actualizar un registro por nombre
Private Sub ActualizarRegistro(nombreRegistro As String, valor As Long)
    Select Case UCase(nombreRegistro)
        Case "EAX": EAX = valor
        Case "EBX": EBX = valor
        Case "ECX": ECX = valor
        Case "EDX": EDX = valor
        Case "ESI": ESI = valor
        Case "EDI": EDI = valor
        Case "EBP": EBP = valor
        Case "ESP": ESP = valor
    End Select
End Sub

' Funci�n para actualizar flags despu�s de ADD
Private Sub ActualizarFlagsADD(operando1 As Long, operando2 As Long, resultado As Long)
    ' Zero Flag: resultado = 0
    ZF = (resultado = 0)
    
    ' Sign Flag: bit m�s significativo = 1 (negativo)
    SF = (resultado And &H80000000) <> 0
    
    ' Carry Flag: overflow sin signo
    ' Si la suma de dos n�meros sin signo excede 2^32-1
    CF = (CLng(operando1) + CLng(operando2)) > &HFFFFFFFF
    
    ' Overflow Flag: overflow con signo
    ' Ocurre cuando sumamos dos positivos y da negativo, o dos negativos y da positivo
    OF = ((operando1 And &H80000000) = (operando2 And &H80000000)) And _
         ((operando1 And &H80000000) <> (resultado And &H80000000))
    
    ' Parity Flag: n�mero de bits 1 es par (solo mira los 8 bits bajos)
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    
    ' Auxiliary Flag: carry del bit 3 al 4 (para BCD)
    AF = ((operando1 And &HF) + (operando2 And &HF)) > &HF
End Sub
' Actualizar flags para operaci�n SUB
Private Sub ActualizarFlagsSUB(operando1 As Long, operando2 As Long, resultado As Long)
    ZF = (resultado = 0)
    SF = (resultado < 0)
    CF = (operando1 < operando2)  ' Para resta, CF indica "pr�stamo"
    OF = ((operando1 Xor operando2) And (operando1 Xor resultado) And &H80000000) <> 0
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    AF = ((operando1 And &HF) < (operando2 And &HF))
End Sub

' Actualizar flags para operaci�n MUL
Private Sub ActualizarFlagsMUL(resultado64 As Currency, resultado32 As Long)
    ' CF y OF se activan si el resultado excede 32 bits
    If resultado64 > 2147483647 Or resultado64 < -2147483648# Then
        CF = True
        OF = True
    Else
        CF = False
        OF = False
    End If
    
    ' Otros flags para MUL
    ZF = (resultado32 = 0)
    SF = (resultado32 < 0)
    PF = (ContarBits1(resultado32 And &HFF) Mod 2 = 0)
    ' AF indefinido para MUL
End Sub

' Actualizar flags para operaci�n IMUL
Private Sub ActualizarFlagsIMUL(operando1 As Long, operando2 As Long, resultado As Long)
    ' Para IMUL, CF y OF se activan si el resultado no cabe en el destino
    If resultado <> CLng(operando1) * CLng(operando2) Then
        CF = True
        OF = True
    Else
        CF = False
        OF = False
    End If
    
    ZF = (resultado = 0)
    SF = (resultado < 0)
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
End Sub

' ========== FUNCIONES AUXILIARES EXISTENTES ==========

' Funci�n auxiliar para contar bits 1
Private Function ContarBits1(valor As Long) As Integer
    Dim i As Integer
    Dim count As Integer
    Dim temp As Long
    
    temp = valor
    count = 0
    For i = 0 To 31
        If (temp And 1) Then count = count + 1
        temp = temp \ 2
    Next i
    ContarBits1 = count
End Function

' Funci�n para verificar si un string es n�mero
Private Function EsNumero(texto As String) As Boolean
    On Error GoTo ErrorHandler
    Dim temp As Double
    temp = CDbl(Replace(texto, "&H", "&H"))
    EsNumero = True
    Exit Function
ErrorHandler:
    EsNumero = False
End Function

' Funci�n para convertir string a n�mero (soporta hexadecimal con &H)
Private Function ConvertirANumero(texto As String) As Long
    If InStr(texto, "&H") > 0 Then
        ' Es hexadecimal
        ConvertirANumero = CLng("&H" & Replace(texto, "&H", ""))
    Else
        ' Es decimal
        ConvertirANumero = CLng(texto)
    End If
End Function



' ========== OPERACIONES ARITM�TICAS (YA EXISTENTES) ==========
Public Sub EjecutarADD(destino As String, origen As String)
    Dim valorDestino As Long
    Dim valorOrigen As Long
    Dim resultado As Long
    
    ' Obtener valores actuales
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    
    ' Realizar la suma
    resultado = valorDestino + valorOrigen
    
    ' Actualizar el registro destino
    ActualizarRegistro destino, resultado
    
    ' Actualizar flags despu�s de la operaci�n
    ActualizarFlagsADD valorDestino, valorOrigen, resultado
End Sub

Public Sub EjecutarSUB(destino As String, origen As String)
    Dim valorDestino As Long
    Dim valorOrigen As Long
    Dim resultado As Long
    
    ' Obtener valores actuales
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    
    ' Realizar la resta
    resultado = valorDestino - valorOrigen
    
    ' Actualizar el registro destino
    ActualizarRegistro destino, resultado
    
    ' Actualizar flags despu�s de la operaci�n
    ActualizarFlagsSUB valorDestino, valorOrigen, resultado
End Sub

Public Sub EjecutarMUL(operando As String)
    Dim valorOperando As Long
    Dim resultado As Long
    Dim resultado64 As Currency ' Para detectar overflow
    
    valorOperando = ObtenerValorOperando(operando)
    
    ' Realizar multiplicaci�n (EAX * operando)
    resultado64 = CDbl(EAX) * CDbl(valorOperando)
    resultado = EAX * valorOperando
    
    ' En x86, MUL almacena el resultado de 64 bits en EDX:EAX
    ' Para simplificar, usaremos solo 32 bits en nuestro simulador educativo
    EAX = resultado
    EDX = 0 ' En una implementaci�n real, contendr�a los bits altos
    
    ' Actualizar flags para multiplicaci�n
    ActualizarFlagsMUL resultado64, resultado
End Sub

Public Sub EjecutarDIV(operando As String)
    Dim valorOperando As Long
    Dim dividendo As Long
    Dim cociente As Long
    Dim residuo As Long
    
    valorOperando = ObtenerValorOperando(operando)
    
    ' Verificar divisi�n por cero
    If valorOperando = 0 Then
        MsgBox "ERROR: Divisi�n por cero", vbCritical, "Error de Ejecuci�n"
        Exit Sub
    End If
    
    ' En x86, DIV divide EDX:EAX por el operando
    ' Para simplificar, usaremos solo EAX como dividendo
    dividendo = EAX
    
    ' Realizar divisi�n
    cociente = dividendo \ valorOperando
    residuo = dividendo Mod valorOperando
    
    ' Almacenar resultados (cociente en EAX, residuo en EDX)
    EAX = cociente
    EDX = residuo
    
    ' La divisi�n no afecta los flags en x86
    ' Pero podemos actualizar ZF si el cociente es cero
    ZF = (cociente = 0)
    SF = (cociente < 0)
End Sub

Public Sub EjecutarIMUL(operando As String)
    Dim valorOperando As Long
    Dim resultado As Long
    
    valorOperando = ObtenerValorOperando(operando)
    
    ' Realizar multiplicaci�n con signo
    resultado = CLng(EAX) * CLng(valorOperando)
    
    EAX = resultado
    EDX = 0 ' Simplificaci�n para nuestro simulador
    
    ' Actualizar flags para multiplicaci�n con signo
    ActualizarFlagsIMUL EAX, valorOperando, resultado
End Sub

Public Sub EjecutarIDIV(operando As String)
    Dim valorOperando As Long
    Dim dividendo As Long
    Dim cociente As Long
    Dim residuo As Long
    
    valorOperando = ObtenerValorOperando(operando)
    
    ' Verificar divisi�n por cero
    If valorOperando = 0 Then
        MsgBox "ERROR: Divisi�n por cero", vbCritical, "Error de Ejecuci�n"
        Exit Sub
    End If
    
    dividendo = EAX
    
    ' Realizar divisi�n con signo
    cociente = CLng(dividendo) \ CLng(valorOperando)
    residuo = CLng(dividendo) Mod CLng(valorOperando)
    
    EAX = cociente
    EDX = residuo
    
    ' Actualizar flags b�sicos
    ZF = (cociente = 0)
    SF = (cociente < 0)
End Sub

Public Sub EjecutarMOV(destino As String, origen As String)
    Dim valorOrigen As Long
    valorOrigen = ObtenerValorOperando(origen)
    ActualizarRegistro destino, valorOrigen
    ' MOV no afecta flags
End Sub
Private Sub ActualizarFlagsLogicos(resultado As Long)
    ZF = (resultado = 0)
    SF = (resultado And &H80000000) <> 0
    PF = (ContarBits1(resultado And &HFF) Mod 2 = 0)
    CF = False
    OF = False
    AF = False
End Sub

Public Sub EjecutarAND(destino As String, origen As String)
    Dim valorDestino As Long
    Dim valorOrigen As Long
    Dim resultado As Long
    
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    resultado = valorDestino And valorOrigen
    
    ActualizarRegistro destino, resultado
    ActualizarFlagsLogicos resultado
End Sub

Public Sub EjecutarOR(destino As String, origen As String)
    Dim valorDestino As Long
    Dim valorOrigen As Long
    Dim resultado As Long
    
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    resultado = valorDestino Or valorOrigen
    
    ActualizarRegistro destino, resultado
    ActualizarFlagsLogicos resultado
End Sub

Public Sub EjecutarXOR(destino As String, origen As String)
    Dim valorDestino As Long
    Dim valorOrigen As Long
    Dim resultado As Long
    
    valorDestino = ObtenerValorRegistro(destino)
    valorOrigen = ObtenerValorOperando(origen)
    resultado = valorDestino Xor valorOrigen
    
    ActualizarRegistro destino, resultado
    ActualizarFlagsLogicos resultado
End Sub

    
Public Sub EjecutarNOT(destino As String)
    Dim valorDestino As Long
    Dim resultado As Long
    
    valorDestino = ObtenerValorRegistro(destino)
    resultado = Not valorDestino
    ActualizarRegistro destino, resultado
End Sub
' funciones que el Parser est� llamando:
   
   Public Sub EjecutarCMP(operando1 As String, operando2 As String)
       ' CMP es como SUB pero no guarda resultado, solo actualiza flags
       Dim val1 As Long, val2 As Long, resultado As Long
       val1 = ObtenerValorRegistro(operando1)
       val2 = ObtenerValorOperando(operando2)
       resultado = val1 - val2
       ActualizarFlagsSUB val1, val2, resultado
   End Sub
   
   Public Sub EjecutarTEST(operando1 As String, operando2 As String)
       ' TEST es como AND pero no guarda resultado
       Dim val1 As Long, val2 As Long, resultado As Long
       val1 = ObtenerValorRegistro(operando1)
       val2 = ObtenerValorOperando(operando2)
       resultado = val1 And val2
       ActualizarFlagsLogicos resultado
   End Sub
   
   Public Sub EjecutarINC(destino As String)
       Dim valor As Long
       valor = ObtenerValorRegistro(destino)
       valor = valor + 1
       ActualizarRegistro destino, valor
       ' INC actualiza todos los flags excepto CF
       ZF = (valor = 0)
       SF = (valor < 0)
       PF = (ContarBits1(valor And &HFF) Mod 2 = 0)
   End Sub
   
   Public Sub EjecutarDEC(destino As String)
       Dim valor As Long
       valor = ObtenerValorRegistro(destino)
       valor = valor - 1
       ActualizarRegistro destino, valor
       ZF = (valor = 0)
       SF = (valor < 0)
       PF = (ContarBits1(valor And &HFF) Mod 2 = 0)
   End Sub
   
   Public Sub EjecutarSHL(destino As String, cantidad As String)
       Dim valor As Long, shift As Integer
       valor = ObtenerValorRegistro(destino)
       shift = CInt(ObtenerValorOperando(cantidad))
       valor = valor * (2 ^ shift) ' Shift left
       ActualizarRegistro destino, valor
       ActualizarFlagsLogicos valor
   End Sub
   
   Public Sub EjecutarSHR(destino As String, cantidad As String)
       Dim valor As Long, shift As Integer
       valor = ObtenerValorRegistro(destino)
       shift = CInt(ObtenerValorOperando(cantidad))
       valor = valor \ (2 ^ shift) ' Shift right
       ActualizarRegistro destino, valor
       ActualizarFlagsLogicos valor
   End Sub
