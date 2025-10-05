Attribute VB_Name = "Cache_Simulador"
' Módulo: Cache_Simulator
' Descripción: Simulador de memoria caché con visualización (Versión Corregida y Mejorada)

Option Explicit

' Constantes para la configuración de la caché
Private Const CACHE_SIZE As Integer = 16      ' 16 líneas de caché
Private Const CACHE_LINE_SIZE As Integer = 4 ' 4 bytes por línea
Private Const RAM_SIZE As Integer = 256
Private Const ROWS As Integer = 16
Private Const COLS As Integer = 16

' Estructura de línea de caché
Private Type cacheLine
    Valid As Boolean
    tag As Long
    data(0 To CACHE_LINE_SIZE - 1) As Byte
    address As Long
    AccessCount As Integer
End Type

' Tipo para instrucciones
Private Type AssemblyInstruction
    address As Long
    OriginalLine As String
    opcode As String
    Operand1 As String
    Operand2 As String
    bytes As String
    Length As Integer
    isDataDefinition As Boolean
End Type

' Variables globales
Private Cache(0 To CACHE_SIZE - 1) As cacheLine
Private RAM(0 To RAM_SIZE - 1) As Byte
Private Program() As AssemblyInstruction
Private SymbolTable As Object ' Usaremos un Dictionary para la tabla de símbolos
Private DataSectionStart As Long
Private TextSectionStart As Long
Private CurrentInstructionIndex As Integer
Private StartInstructionIndex As Integer
Private CacheHits As Integer
Private CacheMisses As Integer
Private TotalAccesses As Integer
Private LastAccessedCacheLine As Integer
Private LastAccessedRAMAddress As Integer

' ===================================================================================
' ============================ INICIALIZACIÓN Y CONTROL =============================
' ===================================================================================

' Punto de entrada principal para configurar el simulador
Sub InitializeCacheSimulator()
    ' Prepara las hojas de Excel
    Application.ScreenUpdating = False
    CreateSampleProgramSheet ' Asegura que el programa de ejemplo exista
    DrawCacheGrid
    DrawRAMGridCache
    
    ' Inicializa el estado del simulador
    ResetCacheSimulator
    Application.ScreenUpdating = True
    
    MsgBox "Simulador de Caché inicializado. Programa cargado y listo para ejecutar."
End Sub

' Reinicia el simulador a su estado inicial
Sub ResetCacheSimulator()
    Application.ScreenUpdating = False
    
    ' Limpiar RAM y caché
    ClearRAM
    ClearCache
    
    ' Reiniciar variables de estado
    CurrentInstructionIndex = -1
    StartInstructionIndex = -1
    LastAccessedCacheLine = -1
    LastAccessedRAMAddress = -1
    Set SymbolTable = CreateObject("Scripting.Dictionary")
    
    ' Leer, parsear y cargar el programa
    ReadAndParseNASM
    LoadNASMProgramIntoRAM
    
    ' Buscar el punto de inicio de la ejecución (_start)
    If SymbolTable.Exists("_start") Then
        Dim startAddr As Long
        startAddr = SymbolTable("_start")
        Dim i As Integer
        For i = 0 To UBound(Program)
            If Program(i).address = startAddr And Not Program(i).isDataDefinition Then
                StartInstructionIndex = i
                CurrentInstructionIndex = i
                Exit For
            End If
        Next i
    End If
    
    If StartInstructionIndex = -1 Then
        MsgBox "Advertencia: No se encontró la etiqueta '_start' en la sección .text. La ejecución no puede comenzar.", vbExclamation
    End If
    
    ' Actualizar visualización
    UpdateRAMDisplayCache
    UpdateCacheDisplay
    UpdateCacheStatus "LISTO", "---", "Presione 'Siguiente' o 'Ejecutar Todo'"
    
    Application.ScreenUpdating = True
End Sub

' Limpia solo la caché, manteniendo la RAM y el programa
Sub ClearCacheOnly()
    Application.ScreenUpdating = False
    ClearCache
    LastAccessedCacheLine = -1
    UpdateCacheDisplay
    UpdateCacheStats
    UpdateCacheStatus "LIMPIO", "---", "Caché vaciada. Estadísticas reiniciadas."
    Application.ScreenUpdating = True
    MsgBox "Caché limpiada. Estadísticas reiniciadas."
End Sub

' ===================================================================================
' =========================== EJECUCIÓN DEL PROGRAMA ================================
' ===================================================================================

' Ejecuta la siguiente instrucción del programa
Sub ExecuteNextInstructionCache()
    If CurrentInstructionIndex = -1 Or CurrentInstructionIndex > UBound(Program) Then
        UpdateCacheStatus "COMPLETADO", "---", "El programa ha finalizado."
        MsgBox "Programa completado."
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Limpiar resaltados de la ejecución anterior
    LastAccessedCacheLine = -1
    LastAccessedRAMAddress = -1
    UpdateRAMDisplayCache
    
    Dim instruction As AssemblyInstruction
    instruction = Program(CurrentInstructionIndex)
    
    ' Saltar directivas o definiciones de datos
    If instruction.isDataDefinition Or LCase(instruction.opcode) = "section" Or LCase(instruction.opcode) = "global" Then
        CurrentInstructionIndex = CurrentInstructionIndex + 1
        ExecuteNextInstructionCache ' Llama recursivamente para pasar a la siguiente real
        Exit Sub
    End If
    
    ' Simular accesos a memoria para la instrucción actual
    SimulateInstructionCacheAccess instruction
    
    ' Actualizar estado visual
    UpdateExecutionStatusCache instruction
    
    ' Avanzar al siguiente índice
    CurrentInstructionIndex = CurrentInstructionIndex + 1
    
    If CurrentInstructionIndex > UBound(Program) Then
        UpdateCacheStatus "COMPLETADO", "---", "Programa finalizado."
    End If
    
    Application.ScreenUpdating = True
End Sub

' Ejecuta el programa completo de una vez
Sub ExecuteFullProgramCache()
    If CurrentInstructionIndex = -1 Or CurrentInstructionIndex > UBound(Program) Then
        MsgBox "El programa ya ha finalizado. Por favor, reinicie el simulador.", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Do While CurrentInstructionIndex <= UBound(Program)
        ' Avanzar por las líneas que no son instrucciones ejecutables
        Do While Program(CurrentInstructionIndex).isDataDefinition Or LCase(Program(CurrentInstructionIndex).opcode) = "section" Or LCase(Program(CurrentInstructionIndex).opcode) = "global"
            CurrentInstructionIndex = CurrentInstructionIndex + 1
            If CurrentInstructionIndex > UBound(Program) Then Exit Do
        Loop
        If CurrentInstructionIndex > UBound(Program) Then Exit Do
        
        Dim instruction As AssemblyInstruction
        instruction = Program(CurrentInstructionIndex)
        
        SimulateInstructionCacheAccess instruction
        UpdateExecutionStatusCache instruction
        
        CurrentInstructionIndex = CurrentInstructionIndex + 1
    Loop
    
    UpdateCacheStatus "COMPLETADO", "---", "Programa finalizado."
    Application.ScreenUpdating = True
    MsgBox "Ejecución completa."
End Sub

' Simula los accesos a memoria que realizaría una instrucción
Sub SimulateInstructionCacheAccess(instruction As AssemblyInstruction)
    Dim memAddr As Long
    
    ' Acceso para leer la propia instrucción desde la RAM
    MemoryAccess instruction.address, False ' Lectura de la instrucción
    
    ' Detectar y simular accesos a memoria en operandos
    If InStr(instruction.Operand1, "[") > 0 Then
        memAddr = ExtractAddressFromOperand(instruction.Operand1)
        If LCase(instruction.opcode) = "mov" Then ' MOV [mem], reg es escritura
            MemoryAccess memAddr, True, &HAA ' Escritura con dato de ejemplo (0xAA)
        Else ' ADD [mem], reg; CMP [mem], reg etc. son lectura
            MemoryAccess memAddr, False
        End If
    End If
    
    If InStr(instruction.Operand2, "[") > 0 Then
        memAddr = ExtractAddressFromOperand(instruction.Operand2)
        MemoryAccess memAddr, False ' Siempre es lectura para el segundo operando (ej: MOV reg, [mem])
    End If
End Sub

' ===================================================================================
' ======================== LÓGICA DE CACHÉ Y MEMORIA ================================
' ===================================================================================

' Simula un acceso a una dirección de memoria, gestionando la caché
Function MemoryAccess(address As Long, isWrite As Boolean, Optional data As Byte = 0) As Boolean
    Dim cacheIndex As Long
    Dim tag As Long
    Dim offset As Integer
    Dim hit As Boolean
    
    ' Calcular componentes de la dirección para el mapeo directo
    offset = address Mod CACHE_LINE_SIZE
    cacheIndex = (address \ CACHE_LINE_SIZE) Mod CACHE_SIZE
    tag = address \ (CACHE_SIZE * CACHE_LINE_SIZE)
    
    TotalAccesses = TotalAccesses + 1
    LastAccessedRAMAddress = address
    LastAccessedCacheLine = cacheIndex
    
    ' Verificar si es HIT o MISS
    If Cache(cacheIndex).Valid And Cache(cacheIndex).tag = tag Then
        ' Cache HIT
        hit = True
        CacheHits = CacheHits + 1
        Cache(cacheIndex).AccessCount = Cache(cacheIndex).AccessCount + 1
        UpdateCacheStatus "HIT", "0x" & Hex(address), "Línea: " & cacheIndex & ", Tag: 0x" & Hex(tag)
    Else
        ' Cache MISS
        hit = False
        CacheMisses = CacheMisses + 1
        ' Cargar el bloque completo desde RAM a la línea de caché
        LoadCacheLineFromRAM cacheIndex, address, tag
        UpdateCacheStatus "MISS", "0x" & Hex(address), "Línea: " & cacheIndex & " cargada desde RAM"
    End If
    
    ' Si es una operación de escritura (Write-Through)
    If isWrite Then
        RAM(address) = data ' Siempre escribe en RAM
        If hit Then ' Si fue hit, también actualiza la caché
            Cache(cacheIndex).data(offset) = data
        End If
        ' Si fue miss, LoadCacheLineFromRAM ya trajo el bloque viejo; ahora actualizamos el byte específico
        Cache(cacheIndex).data(offset) = data
        UpdateRAMDisplayCache ' Actualizar visualización de RAM por la escritura
    End If
    
    ' Actualizar visualización de la caché y estadísticas
    UpdateCacheDisplay
    
    MemoryAccess = hit
End Function

' Carga un bloque de memoria de la RAM a una línea específica de la caché
Sub LoadCacheLineFromRAM(cacheIndex As Long, address As Long, tag As Long)
    Dim i As Integer
    Dim baseAddress As Long
    
    ' Calcular la dirección de inicio del bloque en RAM
    baseAddress = (address \ CACHE_LINE_SIZE) * CACHE_LINE_SIZE
    
    With Cache(cacheIndex)
        .Valid = True
        .tag = tag
        .address = baseAddress ' Guardamos la dirección base del bloque
        .AccessCount = 1 ' Es el primer acceso a esta nueva línea
        
        ' Copiar los datos del bloque desde la RAM a la línea de caché
        For i = 0 To CACHE_LINE_SIZE - 1
            If baseAddress + i < RAM_SIZE Then
                .data(i) = RAM(baseAddress + i)
            Else
                .data(i) = 0 ' Fuera de los límites de la RAM
            End If
        Next i
    End With
End Sub

' Permite al usuario realizar un acceso manual a memoria
Sub ManualMemoryAccess()
    Dim addrStr As String
    Dim addr As Long
    Dim accessType As String
    
    addrStr = InputBox("Ingrese la dirección de memoria (en decimal o hexadecimal con '0x'):", "Acceso Manual a Memoria")
    If addrStr = "" Then Exit Sub
    
    On Error Resume Next
    If LCase(Left(addrStr, 2)) = "0x" Then
        addr = Application.WorksheetFunction.Hex2Dec(Mid(addrStr, 3))
    Else
        addr = CLng(addrStr)
    End If
    If Err.Number <> 0 Then
        MsgBox "Dirección inválida.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    If addr < 0 Or addr >= RAM_SIZE Then
        MsgBox "La dirección está fuera del rango de la RAM (0-" & RAM_SIZE - 1 & ").", vbCritical
        Exit Sub
    End If
    
    accessType = InputBox("Ingrese el tipo de acceso ('R' para Lectura, 'W' para Escritura):", "Tipo de Acceso")
    If UCase(accessType) = "R" Then
        MemoryAccess addr, False
    ElseIf UCase(accessType) = "W" Then
        MemoryAccess addr, True, &HFF ' Escribir un valor de ejemplo
    Else
        MsgBox "Tipo de acceso no válido.", vbInformation
    End If
End Sub


' ===================================================================================
' ===================== LECTURA Y PARSEO DE CÓDIGO NASM =============================
' ===================================================================================

' Orquesta el proceso de lectura y parseo en dos pasadas
Sub ReadAndParseNASM()
    Dim ws As Worksheet
    Set ws = Worksheets("ProgramaNASM")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.ROWS.count, "A").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No se encontró programa en la hoja 'ProgramaNASM'.", vbCritical
        Exit Sub
    End If

    ' Contar líneas válidas para dimensionar el array
    Dim validLines() As String
    Dim validLineCount As Integer
    validLineCount = 0
    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) <> "" Then
            validLineCount = validLineCount + 1
            ReDim Preserve validLines(1 To validLineCount)
            validLines(validLineCount) = Trim(ws.Cells(i, 1).value)
        End If
    Next i

    If validLineCount = 0 Then
        MsgBox "No se encontraron líneas de código válidas.", vbCritical
        Exit Sub
    End If

    ReDim Program(0 To validLineCount - 1)
    
    DataSectionStart = &H0
    TextSectionStart = &H80

    ' --- PRIMERA PASADA: Construir la Tabla de Símbolos ---
    ParsePassOne validLines
    
    ' --- SEGUNDA PASADA: Parsear instrucciones y resolver operandos ---
    ParsePassTwo validLines
End Sub

' Primera pasada: encuentra etiquetas y las añade a la tabla de símbolos
Sub ParsePassOne(lines() As String)
    Dim currentAddress As Long
    Dim currentSection As String
    currentSection = ".data" ' Asumir .data por defecto
    currentAddress = DataSectionStart
    Dim i As Integer

    For i = LBound(lines) To UBound(lines)
        Dim line As String, cleanLine As String
        line = lines(i)
        cleanLine = Trim(LCase(Split(line, ";")(0))) ' Ignorar comentarios y espacios

        If InStr(cleanLine, "section .text") > 0 Then
            currentSection = ".text"
            currentAddress = TextSectionStart
        ElseIf InStr(cleanLine, "section .data") > 0 Then
            currentSection = ".data"
            currentAddress = DataSectionStart
        ElseIf cleanLine <> "" Then ' Solo procesar líneas no vacías
            Dim parts() As String
            parts = Split(Trim(line), " ")
            
            ' Es una etiqueta (ej: _start:)
            If Right(parts(0), 1) = ":" Then
                Dim labelName As String
                labelName = Left(parts(0), Len(parts(0)) - 1)
                SymbolTable(labelName) = currentAddress
                ' Las etiquetas no consumen espacio de memoria
            ElseIf UBound(parts) >= 1 Then
                ' Es una definición de datos (dd, db, dw)
                If LCase(parts(1)) = "dd" Or LCase(parts(1)) = "db" Or LCase(parts(1)) = "dw" Then
                    SymbolTable(parts(0)) = currentAddress
                    ' Avanzar la dirección según el tamaño del dato
                    Select Case LCase(parts(1))
                        Case "dd": currentAddress = currentAddress + 4
                        Case "dw": currentAddress = currentAddress + 2
                        Case "db": currentAddress = currentAddress + 1
                    End Select
                ElseIf currentSection = ".text" And LCase(parts(0)) <> "global" Then
                    ' Es una instrucción, avanzar 4 bytes (simplificación)
                    currentAddress = currentAddress + 4
                End If
            End If
        End If
    Next i
End Sub

' Segunda pasada: parsea cada línea en la estructura de instrucción
Sub ParsePassTwo(lines() As String)
    Dim currentAddress As Long
    Dim currentSection As String
    currentSection = ".data" ' Asumir .data por defecto
    currentAddress = DataSectionStart
    Dim i As Integer

    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = lines(i)
        
        With Program(i - 1)
            .OriginalLine = line
            
            Dim cleanLine As String
            cleanLine = LCase(Split(line, ";")(0))
            
            If InStr(cleanLine, "section .text") > 0 Then
                currentSection = ".text"
                currentAddress = TextSectionStart
                .opcode = "section"
                .address = currentAddress
            ElseIf InStr(cleanLine, "section .data") > 0 Then
                currentSection = ".data"
                currentAddress = DataSectionStart
                .opcode = "section"
                .address = currentAddress
            Else
                .address = currentAddress
                ParseInstructionLine line, i - 1
                
                ' Avanzar dirección si no es una etiqueta
                If Not Right(.opcode, 1) = ":" And .opcode <> "global" Then
                    currentAddress = currentAddress + .Length
                End If
            End If
        End With
    Next i
End Sub

' Parsea una línea de instrucción individual
Sub ParseInstructionLine(line As String, index As Integer)
    Dim mainParts() As String, parts() As String, commentPart As String
    
    ' Separar comentarios
    commentPart = Split(line, ";")(0)
    
    ' Separar operandos por coma
    mainParts = Split(commentPart, ",")
    
    ' Separar opcode del primer operando (usando espacio como separador)
    parts = Split(Trim(mainParts(0)), " ")
    
    With Program(index)
        .opcode = Trim(parts(0))
        
        ' Verificar y asignar operandos de manera segura
        If UBound(parts) >= 1 Then
            .Operand1 = Trim(parts(1))
        Else
            .Operand1 = ""
        End If
        
        If UBound(mainParts) >= 1 Then
            .Operand2 = Trim(mainParts(1))
        Else
            .Operand2 = ""
        End If
    End With
End Sub


' Carga los bytes del programa en la RAM simulada
Sub LoadNASMProgramIntoRAM()
    Dim i As Integer, j As Integer
    Dim bytes() As String
    Dim byteValue As Byte
    
    For i = 0 To UBound(Program)
        If Not Program(i).isDataDefinition And Program(i).bytes <> "" Then
            bytes = Split(Program(i).bytes, " ")
            For j = 0 To UBound(bytes)
                If Program(i).address + j < RAM_SIZE Then
                    byteValue = CInt("&H" & bytes(j))
                    RAM(Program(i).address + j) = byteValue
                End If
            Next j
        End If
    Next i
End Sub

' Extrae la dirección de un operando como [num1] o [128]
Function ExtractAddressFromOperand(operand As String) As Long
    Dim cleanOperand As String
    cleanOperand = Replace(Replace(operand, "[", ""), "]", "")
    
    ' Si es una variable, buscar en la tabla de símbolos
    If SymbolTable.Exists(cleanOperand) Then
        ExtractAddressFromOperand = SymbolTable(cleanOperand)
    Else ' Si no, asumir que es una dirección numérica
        On Error Resume Next
        ExtractAddressFromOperand = CLng(cleanOperand)
        If Err.Number <> 0 Then ExtractAddressFromOperand = -1 ' Dirección inválida
        On Error GoTo 0
    End If
End Function

' Genera bytes de máquina falsos para la simulación visual
Function GenerateSimpleBytes(opcode As String) As String
    Select Case LCase(opcode)
        Case "mov": GenerateSimpleBytes = "B8 12 34 56"
        Case "add": GenerateSimpleBytes = "01 C0 90 90"
        Case "xor": GenerateSimpleBytes = "31 DB 90 90"
        Case "int": GenerateSimpleBytes = "CD 80 90 90"
        Case Else: GenerateSimpleBytes = "90 90 90 90" ' NOP
    End Select
End Function


' ===================================================================================
' ================== DIBUJO Y ACTUALIZACIÓN DE LA INTERFAZ ==========================
' ===================================================================================

' Dibuja la cuadrícula de la caché y los controles
Sub DrawCacheGrid()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("Cache")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "Cache"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    ws.Activate
    
    With ws
        .Range("B1").value = "MEMORIA CACHÉ - Mapeo Directo"
        .Range("B1:G1").Merge
        .Range("B1").Font.Bold = True
        .Range("B1").Font.size = 16
        
        .Range("I1").value = "ESTADÍSTICAS DE CACHÉ": .Range("I1:K1").Merge: .Range("I1").Font.Bold = True
        .Range("I2").value = "Total Accesos:": .Range("J2").value = 0
        .Range("I3").value = "Cache Hits:": .Range("J3").value = 0: .Range("J3").Interior.color = RGB(180, 255, 180)
        .Range("I4").value = "Cache Misses:": .Range("J4").value = 0: .Range("J4").Interior.color = RGB(255, 180, 180)
        .Range("I5").value = "Hit Rate:": .Range("J5").value = "0%"
        
        Dim headers As Variant
        headers = Array("Línea", "Válido", "Tag (Hex)", "Dirección Bloque", "Datos (Hex)", "Accesos")
        .Range("A3").Resize(1, UBound(headers) + 1).value = headers
        With .Range("A3").Resize(1, UBound(headers) + 1)
            .Font.Bold = True
            .Interior.color = RGB(220, 220, 220)
            .HorizontalAlignment = xlCenter
        End With

        Dim i As Integer
        For i = 0 To CACHE_SIZE - 1
            .Cells(i + 4, 1).value = i
        Next i
        
        With .Range("A4").Resize(CACHE_SIZE, UBound(headers) + 1)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        
        .Columns("A:G").AutoFit
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 20
    End With
    
    CreateCacheControlButtons
End Sub

' Dibuja la cuadrícula de la RAM y el panel de estado
Sub DrawRAMGridCache()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("RAM")
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "RAM"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    With ws
        .Range("B1").value = "MEMORIA RAM": .Range("B1:J1").Merge: .Range("B1").Font.Bold = True: .Range("B1").Font.size = 16
        
        .Range("S2").value = "ESTADO DEL SIMULADOR": .Range("S2:U2").Merge: .Range("S2").Font.Bold = True
        .Range("S3").value = "Última Instrucción:"
        .Range("S4").value = "Último Acceso:"
        .Range("S5").value = "Dirección:"
        .Range("S6").value = "Detalles:"
        
        .Range("T3:T6").value = "---"
        .Range("T3:T6").HorizontalAlignment = xlLeft
        .Range("S3:S6").Font.Bold = True
        .Columns("S").AutoFit
        .Columns("T").ColumnWidth = 30
        
        ' Encabezados de columnas (0-F)
        Dim i As Integer
        For i = 0 To COLS - 1
            .Cells(8, i + 2).value = Hex(i)
        Next i
        With .Range(.Cells(8, 2), .Cells(8, COLS + 1))
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.color = RGB(220, 220, 220)
        End With

        ' Encabezados de filas y celdas
        Dim r As Integer, addr As Long
        For r = 0 To ROWS - 1
            addr = r * COLS
            .Cells(r + 9, 1).value = "0x" & Format(Hex(addr), "00")
            For i = 0 To COLS - 1
                .Cells(r + 9, i + 2).value = "00"
            Next i
        Next r
        
        With .Range("A9").Resize(ROWS, 1)
            .Font.Bold = True
            .Interior.color = RGB(220, 220, 220)
        End With
        With .Range("B9").Resize(ROWS, COLS)
            .HorizontalAlignment = xlCenter
            .Font.Name = "Courier New"
            .Borders.LineStyle = xlContinuous
        End With
    End With
End Sub

' Actualiza la visualización de la RAM
Sub UpdateRAMDisplayCache()
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    Dim r As Integer, c As Integer, addr As Long
    
    For r = 0 To ROWS - 1
        For c = 0 To COLS - 1
            addr = r * COLS + c
            With ws.Cells(r + 9, c + 2)
                .value = Format(Hex(RAM(addr)), "00")
                
                ' Color de fondo por sección
                If addr >= TextSectionStart Then
                    .Interior.color = RGB(220, 255, 220) ' Verde para .text
                Else
                    .Interior.color = RGB(220, 220, 255) ' Azul para .data
                End If
                .Font.Bold = False
            End With
        Next c
    Next r
    
    ' Resaltar el último acceso a RAM
    If LastAccessedRAMAddress <> -1 Then
        r = (LastAccessedRAMAddress \ COLS)
        c = (LastAccessedRAMAddress Mod COLS)
        With ws.Cells(r + 9, c + 2)
            .Interior.color = RGB(255, 255, 0) ' Amarillo
            .Font.Bold = True
        End With
    End If
End Sub

' Actualiza la visualización de la caché
Sub UpdateCacheDisplay()
    Dim ws As Worksheet
    Set ws = Worksheets("Cache")
    Dim i As Integer, j As Integer
    
    For i = 0 To CACHE_SIZE - 1
        With Cache(i)
            ' Válido
            ws.Cells(i + 4, 2).value = IIf(.Valid, "Sí", "No")
            ' Tag
            ws.Cells(i + 4, 3).value = "0x" & Hex(.tag)
            ' Dirección y Datos
            If .Valid Then
                ws.Cells(i + 4, 4).value = "0x" & Format(Hex(.address), "00")
                Dim dataStr As String
                dataStr = ""
                For j = 0 To CACHE_LINE_SIZE - 1
                    dataStr = dataStr & Format(Hex(.data(j)), "00") & " "
                Next j
                ws.Cells(i + 4, 5).value = Trim(dataStr)
            Else
                ws.Cells(i + 4, 4).value = "---"
                ws.Cells(i + 4, 5).value = "00 00 00 00"
            End If
            ' Accesos
            ws.Cells(i + 4, 6).value = .AccessCount
        End With
        
        ' Color de fondo base (sin el resaltado de acceso)
        If Cache(i).Valid Then
            ws.Range("B" & i + 4 & ":F" & i + 4).Interior.color = RGB(220, 255, 220) ' Verde claro
        Else
            ws.Range("B" & i + 4 & ":F" & i + 4).Interior.color = RGB(255, 220, 220) ' Rojo claro
        End If
    Next i
    
    ' Resaltar la última línea accedida
    If LastAccessedCacheLine <> -1 Then
        ws.Range("A" & LastAccessedCacheLine + 4 & ":F" & LastAccessedCacheLine + 4).Interior.color = RGB(255, 255, 0) ' Amarillo
    End If
    
    UpdateCacheStats
End Sub

' Actualiza el panel de estado general en la hoja RAM
Sub UpdateCacheStatus(status As String, address As String, details As String)
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    
    ws.Range("T4").value = status
    ws.Range("T5").value = address
    ws.Range("T6").value = details
    
    ' Color según el estado
    Select Case status
        Case "HIT": ws.Range("T4").Interior.color = RGB(0, 255, 0)
        Case "MISS": ws.Range("T4").Interior.color = RGB(255, 0, 0)
        Case "COMPLETADO": ws.Range("T4").Interior.color = RGB(0, 255, 255)
        Case "LISTO", "LIMPIO": ws.Range("T4").Interior.color = RGB(200, 200, 255)
        Case Else: ws.Range("T4").Interior.color = RGB(255, 255, 255)
    End Select
End Sub

' Actualiza el indicador de la instrucción actual en la hoja RAM
Sub UpdateExecutionStatusCache(instruction As AssemblyInstruction)
    Dim ws As Worksheet
    Set ws = Worksheets("RAM")
    ws.Range("T3").value = instruction.OriginalLine
    ws.Range("T3").Interior.color = RGB(255, 255, 0) ' Amarillo
End Sub

' Actualiza las estadísticas de la caché
Sub UpdateCacheStats()
    Dim ws As Worksheet
    Set ws = Worksheets("Cache")
    
    ws.Range("J2").value = TotalAccesses
    ws.Range("J3").value = CacheHits
    ws.Range("J4").value = CacheMisses
    
    If TotalAccesses > 0 Then
        ws.Range("J5").value = Format((CacheHits / TotalAccesses), "0.0%")
    Else
        ws.Range("J5").value = "0%"
    End If
End Sub

' ===================================================================================
' ========================= UTILIDADES Y CONFIGURACIÓN ==============================
' ===================================================================================

' Limpia la memoria RAM
Sub ClearRAM()
    Dim i As Long
    For i = 0 To RAM_SIZE - 1
        RAM(i) = 0
    Next i
End Sub

' Limpia la caché y resetea las estadísticas
Sub ClearCache()
    Dim i As Integer, j As Integer
    For i = 0 To CACHE_SIZE - 1
        Cache(i).Valid = False
        Cache(i).tag = 0
        Cache(i).address = 0
        Cache(i).AccessCount = 0
        For j = 0 To CACHE_LINE_SIZE - 1
            Cache(i).data(j) = 0
        Next j
    Next i
    
    CacheHits = 0
    CacheMisses = 0
    TotalAccesses = 0
End Sub


' Crea los botones de control en la hoja de la Caché
Sub CreateCacheControlButtons()
    Dim ws As Worksheet
    Set ws = Worksheets("Cache")
    
    ' Limpiar botones existentes para evitar duplicados
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    Dim btn As Button
    Set btn = ws.Buttons.Add(50, 350, 120, 30)
    btn.OnAction = "ExecuteNextInstructionCache"
    btn.Characters.Text = "Siguiente Instrucción"
    
    Set btn = ws.Buttons.Add(180, 350, 120, 30)
    btn.OnAction = "ExecuteFullProgramCache"
    btn.Characters.Text = "Ejecutar Todo"
    
    Set btn = ws.Buttons.Add(310, 350, 120, 30)
    btn.OnAction = "ResetCacheSimulator"
    btn.Characters.Text = "Reiniciar Simulador"
    
    Set btn = ws.Buttons.Add(440, 350, 120, 30)
    btn.OnAction = "ManualMemoryAccess"
    btn.Characters.Text = "Acceso Manual"
    
    Set btn = ws.Buttons.Add(570, 350, 100, 30)
    btn.OnAction = "ClearCacheOnly"
    btn.Characters.Text = "Limpiar Caché"
End Sub

' Crea una hoja con un programa NASM de ejemplo si no existe
Sub CreateSampleProgramSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("ProgramaNASM")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "ProgramaNASM"
    End If
    
    If ws.Range("A2").value <> "" Then Exit Sub ' No sobrescribir si ya hay algo
    
    ws.Cells.Clear
    ws.Range("A1").value = "Código NASM (Puede editar este código y reiniciar el simulador)"
    ws.Range("A1").Font.Bold = True
    
    Dim programCode As Variant
    programCode = Array( _
        "section .data", _
        "    num1 dd 10", _
        "    num2 dd 20", _
        "    result dd 0", _
        "", _
        "section .text", _
        "    global _start", _
        "", _
        "_start:", _  ' <--- AQUÍ ESTABA EL ERROR: FALTABA ESTA LÍNEA
        "    ; Cargar num1 en EAX", _
        "    mov eax, [num1]", _
        "    ; Sumar num2 a EAX", _
        "    add eax, [num2]", _
        "    ; Guardar resultado", _
        "    mov [result], eax", _
        "    ; Acceso a dirección fija para demostrar caché", _
        "    mov ebx, [128]", _
        "    ; Salir del programa", _
        "    mov eax, 1", _
        "    xor ebx, ebx", _
        "    int 0x80")
        
    ws.Range("A2").Resize(UBound(programCode) + 1, 1).value = Application.Transpose(programCode)
    ws.Columns("A").ColumnWidth = 50
    ws.Columns("A").WrapText = True
End Sub

