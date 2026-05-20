Attribute VB_Name = "CNG_R2_N1_FT"
Option Explicit
' ======================================================================================
'@Nombre descriptivo: "Formato de Texto Minusculas, Acentos y Puntuacion"
'@nivel Nivel 1
'@regla R2 Flexibles
'@desc: Aplica validación de datos para forzar mayúsculas y sin acentos.
'@use   Pestaña de Validaciones Particulares apartado Flexibles.
'@fecha 11-Mar-2026
' ======================================================================================

Public Function EjecutarFormatoTexto(ByRef wbCenso As Workbook) As Range 'Se pasa el libro por medio de la variable wbCenso
    
    Dim wsDestino As Worksheet 'En wsDestino se guardara la variable de la celda
    Dim prevCalc As XlCalculation
    
    ' 1. Verificación de seguridad
    If wbCenso Is Nothing Then
        MsgBox "No se ha detectado el libro del censo. Por favor, cárguelo de nuevo.", vbCritical
        Exit Function
    End If

    ' Forzamos que trabaje en la hoja que el usuario está viendo del censo
    Set wsDestino = wbCenso.ActiveSheet
    
    ' 2. Manejo de errores
    On Error GoTo Err_Fail

    ' 3. Preparar entornog
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationAutomatic 'Se aplica el calculo automatico para evitar valores obsoletos en el excel durante la ejecucion de la macro
    
    ' --- INICIO DE LÓGICA DE NEGOCIO ---
    Dim rangoEntrada As Range
    Dim i As Long, colAuxiliar As Long
    Dim colLetraAux As String
    Dim filaReferencia As Long
    Dim columnaLimpia As Boolean
    
    ' Paso 1: Selección de rango
    On Error Resume Next
    Set rangoEntrada = Application.InputBox("Selecciona el rango donde se ingresará la validación (Celdas Amarillas):", "PASO 1", Type:=8)
    On Error GoTo Err_Fail
    
    If rangoEntrada Is Nothing Then Exit Function

    filaReferencia = rangoEntrada.Row

    ' Buscar columna auxiliar limpia (AG en adelante)
    colAuxiliar = 33
    columnaLimpia = False
    
    Do While columnaLimpia = False
        ' ESENCIAL: Referenciar siempre a wsDestino
        If wsDestino.Cells(filaReferencia, colAuxiliar).HasFormula = False And _
           wsDestino.Cells(filaReferencia, colAuxiliar).value = "" Then
            columnaLimpia = True
        Else
            colAuxiliar = colAuxiliar + 1
        End If
        If colAuxiliar > 16384 Then Exit Do
    Loop
    
    ' Obtener la letra de la columna encontrada
    colLetraAux = Split(wsDestino.Cells(1, colAuxiliar).Address, "$")(1)

    ' --- Título en la columna auxiliar ---
    With wsDestino.Cells(filaReferencia - 1, colAuxiliar)
        .value = "Validacion Mayusc"
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
    End With

    ' --- Bucle para aplicar la fórmula y la validación ---
    For i = 1 To rangoEntrada.Rows.count
        Dim celdaActual As Range
        Dim celdaAux As Range
        
        Set celdaActual = rangoEntrada.Cells(i, 1)
        ' Referencia explícita a la celda en la hoja de destino
        Set celdaAux = wsDestino.Cells(rangoEntrada.Row + i - 1, colAuxiliar)
        
        ' Aplicar la fórmula de validación (Exactamente la tuya, pero calificada)
        celdaAux.Formula = "=IF(OR(" & celdaActual.Address(False, False) & "="" "",AND(EXACT(UPPER(TRIM(SUBSTITUTE(" & celdaActual.Address(False, False) & ",CHAR(160),"" ""))),TRIM(SUBSTITUTE(" & celdaActual.Address(False, False) & ",CHAR(160),"" ""))),SUMPRODUCT(--ISNUMBER(MATCH(MID(TRIM(SUBSTITUTE(" & celdaActual.Address(False, False) & ",CHAR(160),"" "")),ROW(INDIRECT(""1:""&LEN(TRIM(SUBSTITUTE(" & celdaActual.Address(False, False) & ",CHAR(160),"" ""))))),1),{""A"",""B"",""C"",""D"",""E"",""F"",""G"",""H"",""I"",""J"",""K"",""L"",""M"",""N"",""O"",""P"",""Q"",""R"",""S"",""T"",""U"",""V"",""W"",""X"",""Y"",""Z"",""Ñ"",""0"",""1"",""2"",""3"",""4"",""5"",""6"",""7"",""8"",""9"","" ""},0)))=LEN(TRIM(SUBSTITUTE(" & celdaActual.Address(False, False) & ",CHAR(160),"" ""))))),0,1)"

        ' Configuración de Data Validation en la celda del censo
        With celdaActual.Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
                 Formula1:="=$" & colLetraAux & "$" & celdaAux.Row & "=0"
            .IgnoreBlank = True
            .ErrorTitle = "Dato inválido"
            .ErrorMessage = "El nombre de la posición y/o designación debe registrarse en mayúsculas, sin comillas ni signos de acentuación, puntuación, paréntesis y abreviaturas."
            .ShowError = True
        End With
    Next i

    ' Ajustar columna y finalizar
    wsDestino.Columns(colAuxiliar).AutoFit
    MsgBox "Validación aplicada con éxito en la columna " & colLetraAux, vbInformation

    ' Restaurar cálculo
    Application.Calculation = prevCalc
    Set EjecutarFormatoTexto = rangoEntrada
    Exit Function

Err_Fail:
    Application.Calculation = prevCalc
    MsgBox "Error al aplicar el formato de texto: " & Err.Description, vbCritical
End Function


