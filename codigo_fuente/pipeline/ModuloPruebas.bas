Attribute VB_Name = "ModuloPruebas"
Attribute VB_Name = "ModuloPruebas"
Option Explicit

Private TestsPasados As Long
Private TestsFallados As Long

' =========================================================
'      FUNCIÓN PRINCIPAL - EJECUTA ESTA MACRO
' =========================================================
Public Sub EjecutarTodasLasPruebas()
    TestsPasados = 0
    TestsFallados = 0
    Debug.Print "============================================"
    Debug.Print "         INICIANDO PRUEBAS UNITARIAS"
    Debug.Print "============================================"
    
    ' --- Grupo de Pruebas de Instrucciones ---
    Call Test_MOV_Inmediato_A_Registro
    Call Test_ADD_Registro_A_Registro
    
    ' --- Grupo de Pruebas de Memoria y Caché ---
    Call Test_Cache_Miss_En_Primera_Lectura
    Call Test_Cache_Hit_En_Segunda_Lectura
    
    ' --- Grupo de Pruebas del Pipeline ---
    Call Test_Pipeline_Flujo_Simple_Sin_Riesgos
    Call Test_Pipeline_Hazard_De_Control_JMP
    
    ' --- Resumen Final ---
    Debug.Print "--------------------------------------------"
    Debug.Print "PRUEBAS FINALIZADAS"
    Debug.Print "Total Pasados: " & TestsPasados
    Debug.Print "Total Fallados: " & TestsFallados
    Debug.Print "============================================"
    
    If TestsFallados > 0 Then
        MsgBox "Algunas pruebas fallaron. Revisa la Ventana de Inmediato (Ctrl+G) para más detalles.", vbCritical, "Resultado de Pruebas"
    Else
        MsgBox "¡Todas las pruebas pasaron exitosamente!", vbInformation, "Resultado de Pruebas"
    End If
End Sub

' =========================================================
'                PRUEBAS DE INSTRUCCIONES
' =========================================================
Private Sub Test_MOV_Inmediato_A_Registro()
    ' Arrange: Prepara el estado inicial
    InicializarSimulador
    RAM(0) = "MOV EAX,123"
    
    ' Act: Ejecuta el código
    EjecutarCiclos 5 ' 5 ciclos para que la instrucción complete el pipeline
    
    ' Assert: Verifica el resultado
    AssertEquals 123, EAX, "Test_MOV_Inmediato_A_Registro"
End Sub

Private Sub Test_ADD_Registro_A_Registro()
    ' Arrange
    InicializarSimulador
    EAX = 10
    EBX = 5
    RAM(0) = "ADD EAX,EBX"
    
    ' Act
    EjecutarCiclos 5
    
    ' Assert
    AssertEquals 15, EAX, "Test_ADD_Registro_A_Registro"
End Sub

' =========================================================
'                PRUEBAS DE MEMORIA Y CACHÉ
' =========================================================
Private Sub Test_Cache_Miss_En_Primera_Lectura()
    ' Arrange
    InicializarSimulador
    RAM(0) = "MOV EAX,1"
    
    ' Act
    EjecutarUnCiclo ' El primer ciclo es IF, que causa la lectura
    
    ' Assert
    AssertEquals 1, CacheMisses, "Test_Cache_Miss_En_Primera_Lectura"
    AssertEquals 0, CacheHits, "Test_Cache_Miss_En_Primera_Lectura"
End Sub

Private Sub Test_Cache_Hit_En_Segunda_Lectura()
    ' Arrange
    InicializarSimulador
    RAM(0) = "MOV EAX,1"
    EjecutarUnCiclo ' Causa el primer miss
    
    ' Act
    EIP = 0 ' Forzamos leer la misma dirección otra vez
    EjecutarUnCiclo
    
    ' Assert
    AssertEquals 1, CacheHits, "Test_Cache_Hit_En_Segunda_Lectura"
End Sub

' =========================================================
'                PRUEBAS DEL PIPELINE
' =========================================================
Private Sub Test_Pipeline_Flujo_Simple_Sin_Riesgos()
    ' Arrange
    InicializarSimulador
    RAM(0) = "MOV EAX,1"
    RAM(1) = "MOV EBX,2"
    
    ' Act
    EjecutarUnCiclo
    ' Assert: Ciclo 1
    AssertEquals "MOV EAX,1", IF_ID.Instruccion, "Test_Pipeline_Flujo_Simple (Ciclo 1)"
    
    EjecutarUnCiclo
    ' Assert: Ciclo 2
    AssertEquals "MOV EBX,2", IF_ID.Instruccion, "Test_Pipeline_Flujo_Simple (Ciclo 2 - IF)"
    AssertEquals "MOV EAX,1", ID_EX.Instruccion, "Test_Pipeline_Flujo_Simple (Ciclo 2 - ID)"
End Sub

Private Sub Test_Pipeline_Hazard_De_Control_JMP()
    ' Arrange
    InicializarSimulador
    RAM(0) = "MOV EAX,1"
    RAM(1) = "JMP 10"
    RAM(2) = "MOV EBX,99" ' Esta instrucción debe ser anulada
    
    ' Act
    EjecutarCiclos 3 ' Ciclo 1: IF(MOV), Ciclo 2: ID(MOV), IF(JMP), Ciclo 3: EX(MOV), ID(JMP), IF(MOV EBX)
    
    ' Assert: En la etapa EX del JMP, la instrucción siguiente se anula
    AssertEquals "NOP", IF_ID.Instruccion, "Test_Pipeline_JMP_Anula_IF"
    AssertEquals 10, EIP, "Test_Pipeline_JMP_Actualiza_EIP"
End Sub

' =========================================================
'           HERRAMIENTAS AUXILIARES PARA PRUEBAS
' =========================================================
Private Sub AssertEquals(expected As Variant, actual As Variant, testName As String)
    If actual <> expected Then
        Debug.Print "FAIL: " & testName & " -> Esperado: " & expected & ", Obtenido: " & actual
        TestsFallados = TestsFallados + 1
    Else
        Debug.Print "PASS: " & testName
        TestsPasados = TestsPasados + 1
    End If
End Sub

Private Sub EjecutarCiclos(numeroDeCiclos As Long)
    Dim i As Long
    For i = 1 To numeroDeCiclos
        EjecutarUnCiclo
    Next i
End Sub
