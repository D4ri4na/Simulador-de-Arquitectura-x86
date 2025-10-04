Attribute VB_Name = "ModuloMemoria"
Attribute VB_Name = "ModuloMemoria"
Option Explicit

Private Type CacheLine
    ValidBit As Boolean
    Tag As Long
    LastUsedTimestamp As Double
End Type

Public RAM(0 To MEM_SIZE - 1) As String
Private Cache(0 To CACHE_LINES - 1) As CacheLine
Private currentTimestamp As Double
Public CacheHits As Long
Public CacheMisses As Long

Public Sub InicializarMemoria()
    Dim i As Long
    For i = 0 To UBound(RAM): RAM(i) = "NOP": Next i
    For i = 0 To UBound(Cache)
        Cache(i).ValidBit = False
        Cache(i).Tag = 0
        Cache(i).LastUsedTimestamp = 0
    Next i
    currentTimestamp = 0
    CacheHits = 0
    CacheMisses = 0
End Sub

Public Function LeerDesdeMemoria(direccion As Long) As String
    If direccion < 0 Or direccion >= MEM_SIZE Then
        LeerDesdeMemoria = "NOP"
        Exit Function
    End If

    currentTimestamp = currentTimestamp + 1
    Dim lineIndex As Long, tagVal As Long
    lineIndex = direccion Mod CACHE_LINES
    tagVal = Int(direccion / CACHE_LINES)

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
    If direccion < 0 Or direccion >= MEM_SIZE Then Exit Sub
    RAM(direccion) = valor
    ' Política Write-Through: No actualizamos la caché en la escritura para simplificar.
    ' Una implementación completa podría invalidar o actualizar la línea de caché.
End Sub
