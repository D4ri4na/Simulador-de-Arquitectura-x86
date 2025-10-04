Attribute VB_Name = "ModuloRegistros"
Attribute VB_Name = "ModuloRegistros"
Option Explicit

' --- Registros de propósito general (32 bits) ---
Public EAX As Long
Public EBX As Long
Public ECX As Long
Public EDX As Long

' --- Puntero de Instrucción ---
Public EIP As Long ' Instruction Pointer

' --- Flags de Estado ---
Public ZeroFlag As Boolean
Public CarryFlag As Boolean ' No implementado en esta lógica, pero presente
Public SignFlag As Boolean

' --- Inicializa todos los registros a su estado por defecto ---
Public Sub InicializarRegistros()
    EAX = 0
    EBX = 0
    ECX = 0
    EDX = 0
    EIP = 0
    ZeroFlag = False
    CarryFlag = False
    SignFlag = False
End Sub
