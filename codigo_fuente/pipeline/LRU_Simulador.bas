Attribute VB_Name = "LRU_Simulador"
' Simulador de Caché con Política LRU
Type CacheBlock
    Valid As Boolean
    tag As Long
    Data As String
    LastUsed As Long ' Contador para LRU
    address As String
    Dirty As Boolean
End Type

Type CacheSet
    Blocks() As CacheBlock
End Type

' Variables globales de la caché
Dim Cache() As CacheSet
Dim CacheSize As Long
Dim BlockSize As Long
Dim Associativity As Long
Dim NumberOfSets As Long
Dim AccessCounter As Long
Dim HitCount As Long
Dim MissCount As Long
Dim ReplacementCount As Long

' =============================================
' INICIALIZACIÓN DE LA CACHÉ
' =============================================

Sub InitializeCache(Optional sizeKB As Long = 4, Optional blockSizeBytes As Long = 16, Optional assoc As Long = 2)
    CacheSize = sizeKB * 1024 ' Convertir a bytes
    BlockSize = blockSizeBytes
    Associativity = assoc
    
    ' Calcular número de conjuntos
    NumberOfSets = CacheSize / (BlockSize * Associativity)
    
    ' Redimensionar la caché
    ReDim Cache(0 To NumberOfSets - 1)
    
    Dim i As Long, j As Long
    For i = 0 To NumberOfSets - 1
        ReDim Cache(i).Blocks(0 To Associativity - 1)
        For j = 0 To Associativity - 1
            Cache(i).Blocks(j).Valid = False
            Cache(i).Blocks(j).tag = 0
            Cache(i).Blocks(j).Data = ""
            Cache(i).Blocks(j).LastUsed = 0
            Cache(i).Blocks(j).address = ""
            Cache(i).Blocks(j).Dirty = False
        Next j
    Next i
    
    AccessCounter = 0
    HitCount = 0
    MissCount = 0
    ReplacementCount = 0
    
    CreateCacheDisplay
    CreateCacheLog
    UpdateStatsDisplay
    
    MsgBox "Caché inicializada:" & vbCrLf & _
           "Tamaño: " & sizeKB & " KB" & vbCrLf & _
           "Tamaño de bloque: " & blockSizeBytes & " bytes" & vbCrLf & _
           "Asociatividad: " & assoc & "-way" & vbCrLf & _
           "Número de conjuntos: " & NumberOfSets, vbInformation
End Sub

' =============================================
' FUNCIONES DE ACCESO A CACHÉ
' =============================================

Function AccessMemory(address As Long) As Boolean
    AccessCounter = AccessCounter + 1
    Dim setIndex As Long, tag As Long
    Dim hit As Boolean
    
    ' Calcular conjunto y tag
    setIndex = (address \ BlockSize) Mod NumberOfSets
    tag = address \ (BlockSize * NumberOfSets)
    
    ' Buscar en la caché
    hit = SearchInCache(setIndex, tag)
    
    If hit Then
        HitCount = HitCount + 1
        LogAccess address, "HIT", setIndex, FindBlockIndex(setIndex, tag)
        AccessMemory = True
    Else
        MissCount = MissCount + 1
        LogAccess address, "MISS", setIndex, -1
        HandleCacheMiss setIndex, tag, address
        AccessMemory = False
    End If
    
    UpdateCacheDisplay
    UpdateStatsDisplay
End Function

Function SearchInCache(setIndex As Long, tag As Long) As Boolean
    Dim i As Long
    For i = 0 To Associativity - 1
        If Cache(setIndex).Blocks(i).Valid And Cache(setIndex).Blocks(i).tag = tag Then
            ' Actualizar contador LRU
            Cache(setIndex).Blocks(i).LastUsed = AccessCounter
            SearchInCache = True
            Exit Function
        End If
    Next i
    SearchInCache = False
End Function

Function FindBlockIndex(setIndex As Long, tag As Long) As Long
    Dim i As Long
    For i = 0 To Associativity - 1
        If Cache(setIndex).Blocks(i).Valid And Cache(setIndex).Blocks(i).tag = tag Then
            FindBlockIndex = i
            Exit Function
        End If
    Next i
    FindBlockIndex = -1
End Function

Sub HandleCacheMiss(setIndex As Long, tag As Long, address As Long)
    Dim blockIndex As Long
    Dim replacedAddress As String
    
    ' Buscar bloque vacío
    blockIndex = FindEmptyBlock(setIndex)
    
    If blockIndex = -1 Then
        ' No hay bloques vacíos, necesitamos reemplazar
        blockIndex = FindLRUBlock(setIndex)
        replacedAddress = Cache(setIndex).Blocks(blockIndex).address
        ReplacementCount = ReplacementCount + 1
        
        ' Log del reemplazo
        LogReplacement setIndex, blockIndex, replacedAddress, address
    Else
        LogMessage "Bloque cargado en posición vacía: Set " & setIndex & ", Bloque " & blockIndex
    End If
    
    ' Cargar el nuevo bloque
    Cache(setIndex).Blocks(blockIndex).Valid = True
    Cache(setIndex).Blocks(blockIndex).tag = tag
    Cache(setIndex).Blocks(blockIndex).Data = "Datos[" & Hex(address) & "]"
    Cache(setIndex).Blocks(blockIndex).LastUsed = AccessCounter
    Cache(setIndex).Blocks(blockIndex).address = "0x" & Hex(address)
    Cache(setIndex).Blocks(blockIndex).Dirty = False
End Sub

Function FindEmptyBlock(setIndex As Long) As Long
    Dim i As Long
    For i = 0 To Associativity - 1
        If Not Cache(setIndex).Blocks(i).Valid Then
            FindEmptyBlock = i
            Exit Function
        End If
    Next i
    FindEmptyBlock = -1
End Function

Function FindLRUBlock(setIndex As Long) As Long
    Dim i As Long, lruIndex As Long, minUsage As Long
    minUsage = AccessCounter + 1 ' Valor inicial alto
    
    For i = 0 To Associativity - 1
        If Cache(setIndex).Blocks(i).LastUsed < minUsage Then
            minUsage = Cache(setIndex).Blocks(i).LastUsed
            lruIndex = i
        End If
    Next i
    
    FindLRUBlock = lruIndex
End Function

' =============================================
' LOGGING Y VISUALIZACIÓN
' =============================================

Sub LogAccess(address As Long, accessType As String, setIndex As Long, blockIndex As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LogCaché")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = "0x" & Hex(address)
    ws.Cells(nextRow, 3).value = accessType
    ws.Cells(nextRow, 4).value = setIndex
    ws.Cells(nextRow, 5).value = IIf(blockIndex >= 0, blockIndex, "N/A")
    ws.Cells(nextRow, 6).value = "Acceso #" & AccessCounter
    
    ' Resaltar misses
    If accessType = "MISS" Then
        ws.ROWS(nextRow).Interior.Color = RGB(255, 200, 200)
    Else
        ws.ROWS(nextRow).Interior.Color = RGB(200, 255, 200)
    End If
    
    ws.Columns.AutoFit
End Sub

Sub LogReplacement(setIndex As Long, blockIndex As Long, oldAddress As String, newAddress As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LogCaché")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = "REEMPLAZO LRU"
    ws.Cells(nextRow, 3).value = "Set " & setIndex & ", Bloque " & blockIndex
    ws.Cells(nextRow, 4).value = oldAddress & " ? " & "0x" & Hex(newAddress)
    ws.Cells(nextRow, 5).value = "Bloque reemplazado por ser el menos recientemente usado"
    ws.Cells(nextRow, 6).value = "Contador LRU: " & Cache(setIndex).Blocks(blockIndex).LastUsed
    
    ' Resaltar reemplazos en amarillo
    ws.ROWS(nextRow).Interior.Color = RGB(255, 255, 200)
    
    ' Mensaje detallado
    LogMessage "REEMPLAZO LRU: Bloque en Set " & setIndex & ", Bloque " & blockIndex & _
               " (" & oldAddress & ") reemplazado por dirección " & "0x" & Hex(newAddress) & _
               " por ser el menos recientemente usado (último acceso: " & _
               Cache(setIndex).Blocks(blockIndex).LastUsed & ")"
    
    ws.Columns.AutoFit
End Sub

Sub LogMessage(message As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LogCaché")
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.ROWS.count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = message
    ws.Cells(nextRow, 3).value = ""
    ws.Cells(nextRow, 4).value = ""
    ws.Cells(nextRow, 5).value = ""
    ws.Cells(nextRow, 6).value = ""
    
    ws.Columns.AutoFit
End Sub

Sub CreateCacheDisplay()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("EstadoCaché")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "EstadoCaché"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ' Encabezados
    ws.Cells(1, 1).value = "Conjunto"
    ws.Cells(1, 2).value = "Bloque"
    ws.Cells(1, 3).value = "Válido"
    ws.Cells(1, 4).value = "Tag"
    ws.Cells(1, 5).value = "Dirección"
    ws.Cells(1, 6).value = "Datos"
    ws.Cells(1, 7).value = "Último Uso"
    ws.Cells(1, 8).value = "LRU Rank"
    
    ' Formato de encabezados
    Dim headerRange As Range
    Set headerRange = ws.Range("A1:H1")
    headerRange.Font.Bold = True
    headerRange.Interior.Color = RGB(200, 200, 200)
    
    UpdateCacheDisplay
End Sub

Sub UpdateCacheDisplay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("EstadoCaché")
    
    ' Limpiar datos anteriores (mantener encabezados)
    If ws.Cells(2, 1).value <> "" Then
        ws.Range("A2:H1000").ClearContents
        ws.Range("A2:H1000").Interior.ColorIndex = 0
    End If
    
    Dim row As Long
    row = 2
    Dim i As Long, j As Long
    
    For i = 0 To NumberOfSets - 1
        For j = 0 To Associativity - 1
            With Cache(i).Blocks(j)
                ws.Cells(row, 1).value = i
                ws.Cells(row, 2).value = j
                ws.Cells(row, 3).value = IIf(.Valid, "Sí", "No")
                ws.Cells(row, 4).value = IIf(.Valid, "0x" & Hex(.tag), "N/A")
                ws.Cells(row, 5).value = .address
                ws.Cells(row, 6).value = .Data
                ws.Cells(row, 7).value = IIf(.Valid, .LastUsed, "N/A")
                
                ' Calcular y mostrar ranking LRU
                If .Valid Then
                    ws.Cells(row, 8).value = GetLRURank(i, j)
                    
                    ' Resaltar el bloque LRU actual en cada conjunto
                    If IsLRUBlock(i, j) Then
                        ws.ROWS(row).Interior.Color = RGB(255, 200, 200) ' Rojo claro para LRU
                    ElseIf .LastUsed = AccessCounter Then
                        ws.ROWS(row).Interior.Color = RGB(200, 255, 200) ' Verde claro para más reciente
                    End If
                End If
            End With
            row = row + 1
        Next j
        ' Línea separadora entre conjuntos
        ws.ROWS(row).Interior.Color = RGB(240, 240, 240)
        row = row + 1
    Next i
    
    ws.Columns.AutoFit
End Sub

Function GetLRURank(setIndex As Long, blockIndex As Long) As Long
    If Not Cache(setIndex).Blocks(blockIndex).Valid Then
        GetLRURank = -1
        Exit Function
    End If
    
    Dim usageCounters() As Long
    ReDim usageCounters(0 To Associativity - 1)
    Dim i As Long, j As Long
    Dim rank As Long
    
    ' Recopilar contadores de uso
    For i = 0 To Associativity - 1
        If Cache(setIndex).Blocks(i).Valid Then
            usageCounters(i) = Cache(setIndex).Blocks(i).LastUsed
        Else
            usageCounters(i) = AccessCounter + 1 ' Los inválidos son los "más nuevos"
        End If
    Next i
    
    ' Calcular ranking (1 = más reciente, Associativity = menos reciente)
    For i = 0 To Associativity - 1
        rank = 1
        For j = 0 To Associativity - 1
            If i <> j And usageCounters(j) < usageCounters(i) Then
                rank = rank + 1
            End If
        Next j
        If i = blockIndex Then
            GetLRURank = rank
            Exit Function
        End If
    Next i
End Function

Function IsLRUBlock(setIndex As Long, blockIndex As Long) As Boolean
    If Not Cache(setIndex).Blocks(blockIndex).Valid Then
        IsLRUBlock = False
        Exit Function
    End If
    
    Dim lruIndex As Long
    lruIndex = FindLRUBlock(setIndex)
    IsLRUBlock = (lruIndex = blockIndex)
End Function

Sub CreateCacheLog()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LogCaché")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "LogCaché"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ' Encabezados del log
    ws.Cells(1, 1).value = "Timestamp"
    ws.Cells(1, 2).value = "Dirección/Evento"
    ws.Cells(1, 3).value = "Tipo"
    ws.Cells(1, 4).value = "Set"
    ws.Cells(1, 5).value = "Bloque"
    ws.Cells(1, 6).value = "Información"
    
    ' Formato de encabezados
    Dim headerRange As Range
    Set headerRange = ws.Range("A1:F1")
    headerRange.Font.Bold = True
    headerRange.Interior.Color = RGB(200, 200, 200)
    
    ws.Columns.AutoFit
End Sub

Sub UpdateStatsDisplay()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Estadísticas")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Estadísticas"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Cells(1, 1).value = "ESTADÍSTICAS DE CACHÉ"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.size = 14
    
    ws.Cells(3, 1).value = "Total de Accesos:"
    ws.Cells(3, 2).value = AccessCounter
    
    ws.Cells(4, 1).value = "Hits:"
    ws.Cells(4, 2).value = HitCount
    
    ws.Cells(5, 1).value = "Misses:"
    ws.Cells(5, 2).value = MissCount
    
    ws.Cells(6, 1).value = "Reemplazos:"
    ws.Cells(6, 2).value = ReplacementCount
    
    If AccessCounter > 0 Then
        ws.Cells(7, 1).value = "Tasa de Hit:"
        ws.Cells(7, 2).value = Format(HitCount / AccessCounter, "0.00%")
        
        ws.Cells(8, 1).value = "Tasa de Miss:"
        ws.Cells(8, 2).value = Format(MissCount / AccessCounter, "0.00%")
    End If
    
    ws.Cells(10, 1).value = "Configuración:"
    ws.Cells(11, 1).value = "Tamaño Caché:"
    ws.Cells(11, 2).value = CacheSize / 1024 & " KB"
    ws.Cells(12, 1).value = "Tamaño Bloque:"
    ws.Cells(12, 2).value = BlockSize & " bytes"
    ws.Cells(13, 1).value = "Asociatividad:"
    ws.Cells(13, 2).value = Associativity & "-way"
    ws.Cells(14, 1).value = "Número de Conjuntos:"
    ws.Cells(14, 2).value = NumberOfSets
    
    ws.Columns.AutoFit
End Sub

' =============================================
' FUNCIONES DE PRUEBA Y DEMOSTRACIÓN
' =============================================
Sub TestCacheLRU()
    ' Inicializar caché pequeña para demostrar reemplazos
    InitializeCache 1, 16, 2 ' 1KB, bloques de 16 bytes, 2-way
    
    ' Secuencia de accesos que forzará reemplazos LRU
Dim addresses As Variant
addresses = Array(0, 16, 32, 48, 0, 64, 16, 80, 32, 96)

' Y usarlo directamente como Variant en tu código
Dim i As Long
For i = LBound(addresses) To UBound(addresses)
    ' Tu código aquí usando addresses(i)
Next i
    
   
End Sub

Sub RunCustomTest()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Pruebas")
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Pruebas")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Pruebas"
    End If
    On Error GoTo 0
    
    ' Crear interfaz de prueba simple
    ws.Cells.Clear
    ws.Cells(1, 1).value = "PRUEBAS DE CACHÉ LRU"
    ws.Cells(1, 1).Font.Bold = True
    
    ws.Cells(3, 1).value = "Dirección a acceder (hex):"
    ws.Cells(3, 2).value = "0x"
    
    ws.Cells(5, 1).value = "Accesos de ejemplo:"
    ws.Cells(6, 1).value = "0x0"
    ws.Cells(7, 1).value = "0x10"
    ws.Cells(8, 1).value = "0x20"
    ws.Cells(9, 1).value = "0x30"
    ws.Cells(10, 1).value = "0x40"
    
    ' Botón para acceder a dirección
    Dim btn As Button
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    Set btn = ws.Buttons.Add(ws.Cells(3, 3).Left, ws.Cells(3, 3).Top, 80, 20)
    btn.OnAction = "AccessCustomAddress"
    btn.Text = "Acceder"
    
    ' Botón para prueba automática
    Set btn = ws.Buttons.Add(ws.Cells(5, 3).Left, ws.Cells(5, 3).Top, 120, 20)
    btn.OnAction = "TestCacheLRU"
    btn.Text = "Ejecutar Test LRU"
    
    ' Botón para limpiar
    Set btn = ws.Buttons.Add(ws.Cells(7, 3).Left, ws.Cells(7, 3).Top, 80, 20)
    btn.OnAction = "ClearCache"
    btn.Text = "Limpiar"
    
    ws.Columns.AutoFit
End Sub



Sub AccessCustomAddress()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Pruebas")
    
    Dim addressStr As String
    addressStr = ws.Cells(3, 2).value
    
    ' Convertir dirección hexadecimal a decimal
    Dim address As Long
    If Left(addressStr, 2) = "0x" Then
        addressStr = Mid(addressStr, 3)
    End If
    
    On Error GoTo ErrorHandler
    address = Application.WorksheetFunction.Hex2Dec(addressStr)
    AccessMemory address
    Exit Sub
    
ErrorHandler:
    MsgBox "Dirección inválida: " & addressStr, vbExclamation
End Sub

Sub ClearCache()
    InitializeCache 1, 16, 2
    MsgBox "Caché limpiada y reinicializada", vbInformation
End Sub

' =============================================
' INICIALIZACIÓN DEL SIMULADOR
' =============================================

Sub IniciarSimuladorCache()
    RunCustomTest
    InitializeCache 1, 16, 2 ' Configuración por defecto para demostración
    MsgBox "Simulador de Caché LRU inicializado." & vbCrLf & _
           "Use la pestaña 'Pruebas' para realizar accesos a memoria.", vbInformation
End Sub





