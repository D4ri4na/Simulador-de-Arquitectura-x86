Attribute VB_Name = "ModuloMemoria"
Attribute VB_Name = "ModuloMemoria"
Option Explicit

' Estructura para una línea de caché
Private Type CacheLine
    ValidBit As Boolean
    Tag As Long
    LastUsedTimestamp As Double
End Type

' Memoria Principal (RAM) - AHORA PÚBLICA
Public RAM(0 To 255) As String

' Memoria Caché
Private Cache(0 To 7) As CacheLine
Private currentTimestamp As Double

' Estadísticas de la Caché
Public CacheHits As Long
Public CacheMisses As Long

Public Sub InicializarMemoria()
    Dim i As Long
    For i = 0 To 255: RAM(i) = "NOP": Next i
    For i = 0 To 7
        Cache(i).ValidBit = False
        Cache(i).Tag = 0
        Cache(i).LastUsedTimestamp = 0
    Next i
    currentTimestamp = 0
    CacheHits = 0
    CacheMisses = 0
End Sub

Public Function LeerDesdeMemoria(direccion As Long) As String
    If direccion < 0 Or direccion >= 256 Then
        LeerDesdeMemoria = "NOP"
        Exit Function
    End If
    
    currentTimestamp = currentTimestamp + 1
    Dim lineIndex As Long, tagVal As Long
    lineIndex = direccion Mod 8
    tagVal = Int(direccion / 8)
    
    If Cache(lineIndex).ValidBit And Cache(lineIndex).Tag = tagVal Then
        CacheHits = CacheHits + 1
        Cache(lineIndex).LastUsedTimestamp = currentTimestamp
    Else
        CacheMisses = CacheMisses + 1
        Cache(lineIndex).ValidBit = True
        Cache(lineIndex).Tag = tagVal
        Cache(lineIndex).LastUsedTimestamp = currentTimestamp
    End If
    
    LeerDesdeMemoria = RAM(direccion)
End Function

Public Sub EscribirEnMemoria(direccion As Long, valor As String)
    If direccion < 0 Or direccion >= 256 Then Exit Sub
    RAM(direccion) = valor
    
    Dim lineIndex As Long, tagVal As Long
    lineIndex = direccion Mod 8
    tagVal = Int(direccion / 8)
End Sub

Public Function ObtenerEstadoCache() As Variant
    Dim estado(0 To 8, 1 To 4) As String
    Dim i As Long
    estado(0, 1) = "Índice": estado(0, 2) = "Válido": estado(0, 3) = "Tag": estado(0, 4) = "Timestamp"
    For i = 0 To 7
        estado(i + 1, 1) = CStr(i)
        estado(i + 1, 2) = IIf(Cache(i).ValidBit, "1", "0")
        estado(i + 1, 3) = Hex(Cache(i).Tag)
        estado(i + 1, 4) = CStr(Cache(i).LastUsedTimestamp)
    Next i
    ObtenerEstadoCache = estado
End Function
