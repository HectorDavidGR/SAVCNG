using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core;

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionNumeroTelefonico : IValidacionExcel
    {
        // =====================================================================
        // CONSTANTES
        // =====================================================================

        private const int TotalColumnasAuxiliares = 3;
        private const int MaxColumnasExcel = 16384; // XFD

        // =====================================================================
        // CONTRATO PRINCIPAL
        // =====================================================================

        public ResultadoValidacion Ejecutar(
            Excel.Application excelApp,
            Excel.Workbook libroCenso,
            Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;

            Excel.Range seleccionAuxiliar = null;
            Excel.Range celdaInicioAux = null;

            string etapaActual =
                "Inicialización";

            try
            {
                // =============================================================
                // 1. VALIDACIONES GENERALES
                // =============================================================

                if (excelApp == null)
                {
                    return CrearError(
                        "No fue posible obtener la instancia de Excel.");
                }

                if (libroCenso == null)
                {
                    return CrearError(
                        "No existe un libro de censo cargado.");
                }

                if (rangoCapturado == null)
                {
                    return CrearError(
                        "No existe un rango seleccionado para aplicar la validación.");
                }

                wsActual =
                    (Excel.Worksheet)
                    rangoCapturado.Worksheet;

                // =============================================================
                // FASE 1:
                // AUDITORÍA DE COEXISTENCIA
                // (DETECCIÓN DE REGLAS PREVIAS)
                //
                // Basada en ValidacionFormatoTexto.cs:
                //
                // - Detecta Formatos Condicionales previos.
                // - Detecta Data Validation previa.
                // - SÍ: conserva colores / alertas / bloqueos anteriores.
                // - NO: elimina Formatos Condicionales previos.
                // - CANCELAR: no realiza cambios.
                //
                // La Data Validation de Número telefónico reemplaza la
                // Data Validation previa porque es una restricción de captura.
                // =============================================================

                etapaActual =
                    "Auditoría de coexistencia";

                bool tieneFormatosPrevios =
                    false;

                Excel.FormatConditions formatosPrevios =
                    null;

                try
                {
                    formatosPrevios =
                        rangoCapturado.FormatConditions;

                    tieneFormatosPrevios =
                        formatosPrevios != null &&
                        formatosPrevios.Count > 0;
                }
                catch
                {
                    tieneFormatosPrevios =
                        false;
                }
                finally
                {
                    ExcelHelper.LiberarCom(
                        formatosPrevios);
                }

                bool tieneValidacionPrevia =
                    false;

                try
                {
                    var tipo =
                        rangoCapturado.Validation.Type;

                    tieneValidacionPrevia =
                        true;
                }
                catch
                {
                    // Silenciado:
                    // no existe Data Validation previa.
                    tieneValidacionPrevia =
                        false;
                }

                bool limpiarFormatos =
                    false;

                if (tieneFormatosPrevios ||
                    tieneValidacionPrevia)
                {
                    excelApp.ScreenUpdating =
                        true;

                    DialogResult respLimpieza =
                        MessageBox.Show(
                            "Se detectaron configuraciones previas en el rango seleccionado.\n\n" +

                            "NOTA: Al ser una restricción de captura estricta, la regla de " +
                            "Validación de datos existente se sobreescribirá, pero...\n\n" +

                            "¿Deseas CONSERVAR los colores, alertas o bloqueos " +
                            "(Formatos Condicionales) aplicados previamente?\n\n" +

                            "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                            "NO = BORRAR todo el historial de Formatos Condicionales " +
                            "y aplicar únicamente la validación de Número telefónico.\n" +
                            "CANCELAR = No realizar cambios.",

                            "SAVCNG - Auditoría de Coexistencia",

                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Warning);

                    if (respLimpieza ==
                        DialogResult.Cancel)
                    {
                        return new ResultadoValidacion
                        {
                            Exito = false,

                            Mensaje =
                                "Proceso cancelado.\n" +
                                "No se alteró la plantilla.",

                            AlertaInyectada = false
                        };
                    }

                    if (respLimpieza ==
                        DialogResult.No)
                    {
                        limpiarFormatos =
                            true;
                    }
                }

                // =============================================================
                // FASE 1.1:
                // VALIDACIÓN DEL RANGO CAPTURADO
                // =============================================================

                etapaActual =
                    "Validación del rango de Número telefónico";

                int totalAreas =
                    rangoCapturado.Areas.Count;

                if (totalAreas <= 0)
                {
                    return CrearError(
                        "El rango seleccionado no contiene celdas válidas.");
                }

                // -------------------------------------------------------------
                // VALIDAR ESTRUCTURA DEL RANGO
                //
                // Permite columnas formadas por celdas combinadas.
                //
                // Ejemplo del cuestionario:
                //
                // Número telefónico de contacto
                // (diez dígitos)
                // -------------------------------------------------------------

                for (int i = 1;
                     i <= totalAreas;
                     i++)
                {
                    Excel.Range area = null;
                    Excel.Range primeraCelda = null;
                    Excel.Range mergeArea = null;

                    try
                    {
                        area =
                            (Excel.Range)
                            rangoCapturado.Areas[i];

                        primeraCelda =
                            ObtenerPrimeraCelda(
                                area);

                        if (primeraCelda == null)
                        {
                            return CrearError(
                                "No fue posible identificar la primera celda " +
                                "del rango seleccionado.");
                        }

                        mergeArea =
                            (Excel.Range)
                            primeraCelda.MergeArea;

                        int anchoCapturado =
                            area.Columns.Count;

                        int anchoCombinado =
                            mergeArea.Columns.Count;

                        if (anchoCapturado !=
                            anchoCombinado)
                        {
                            return CrearError(
                                "El rango seleccionado debe contener únicamente " +
                                "el campo 'Número telefónico de contacto (diez dígitos)'.");
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(
                            mergeArea);

                        ExcelHelper.LiberarCom(
                            primeraCelda);

                        ExcelHelper.LiberarCom(
                            area);
                    }
                }

                // =============================================================
                // FASE 1.5:
                // DETECCIÓN DE PREGUNTAS
                // =============================================================

                etapaActual =
                    "Detección de preguntas";

                HashSet<string> preguntasUnicas =
                    new HashSet<string>();

                for (int i = 1;
                     i <= totalAreas;
                     i++)
                {
                    Excel.Range area = null;

                    try
                    {
                        area =
                            (Excel.Range)
                            rangoCapturado.Areas[i];

                        string idPregunta =
                            ExcelHelper.ObtenerNumeroPregunta(
                                area);

                        if (!string.IsNullOrEmpty(
                                idPregunta))
                        {
                            preguntasUnicas.Add(
                                idPregunta);
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(
                            area);
                    }
                }

                string preguntasDetectadas =
                    preguntasUnicas.Count > 0
                        ? string.Join(
                            ", ",
                            preguntasUnicas)
                        : "ND";

                // =============================================================
                // FASE 2:
                // SELECCIÓN DE LA COLUMNA AUXILIAR
                //
                // MISMO FLUJO DE ValidacionCorreoElectronicoDV.cs:
                //
                // - Se mantiene seleccionado el rango capturado.
                // - Se abre InputBox Type = 8.
                // - El usuario selecciona con el mouse dónde comenzar.
                // - Si la selección no sirve, se vuelve directamente al
                //   InputBox SIN MessageBox intermedio.
                // =============================================================

                etapaActual =
                    "Selección de Fórmula Auxiliar";

                try
                {
                    wsActual.Activate();
                    rangoCapturado.Select();
                }
                catch
                {
                    // No crítico.
                }

                excelApp.ScreenUpdating =
                    true;

                int columnaAuxiliarInicial =
                    0;

                bool auxiliarSeleccionado =
                    false;

                while (!auxiliarSeleccionado)
                {
                    object resAuxiliar =
                        excelApp.InputBox(

                            "Indica a partir de qué columna libre deseas colocar las fórmulas " +
                            "auxiliares de la Validación de Número telefónico.\n\n" +

                            "Considera que el sistema requerirá 3 columnas consecutivas " +
                            "y espacio libre hacia abajo proporcional al número de filas " +
                            "que seleccionaste.",

                            "SAVCNG - Selección de Fórmula Auxiliar",

                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            8); // 8 = Range

                    // ---------------------------------------------------------
                    // CANCELACIÓN
                    // ---------------------------------------------------------

                    if (resAuxiliar is bool &&
                        (bool)resAuxiliar == false)
                    {
                        return new ResultadoValidacion
                        {
                            Exito = false,

                            Mensaje =
                                "Configuración de columna auxiliar cancelada.",

                            AlertaInyectada = false
                        };
                    }

                    Excel.Worksheet hojaAuxiliar =
                        null;

                    try
                    {
                        // =====================================================
                        // INTENTAR OBTENER EL RANGE DIRECTAMENTE
                        // =====================================================

                        seleccionAuxiliar =
                            resAuxiliar as Excel.Range;

                        // =====================================================
                        // RESPALDO:
                        //
                        // Si Excel devuelve Int32 u otro tipo,
                        // recuperar Application.Selection.
                        //
                        // Evita:
                        //
                        // System.InvalidCastException
                        // Int32 -> Excel.Range
                        // =====================================================

                        if (seleccionAuxiliar == null)
                        {
                            object seleccionActual =
                                null;

                            try
                            {
                                seleccionActual =
                                    excelApp.Selection;

                                seleccionAuxiliar =
                                    seleccionActual as Excel.Range;

                                if (seleccionAuxiliar != null)
                                {
                                    /*
									 * seleccionAuxiliar será liberado
									 * posteriormente.
									 */
                                    seleccionActual =
                                        null;
                                }
                            }
                            catch
                            {
                                seleccionAuxiliar =
                                    null;
                            }
                            finally
                            {
                                if (seleccionActual != null)
                                {
                                    ExcelHelper.LiberarCom(
                                        seleccionActual);
                                }
                            }
                        }

                        // =====================================================
                        // NO SE OBTUVO UN RANGE
                        //
                        // VOLVER DIRECTAMENTE AL INPUTBOX.
                        // =====================================================

                        if (seleccionAuxiliar == null)
                        {
                            continue;
                        }

                        // =====================================================
                        // OBTENER PRIMERA CELDA
                        // =====================================================

                        celdaInicioAux =
                            ObtenerPrimeraCelda(
                                seleccionAuxiliar);

                        if (celdaInicioAux == null)
                        {
                            continue;
                        }

                        // =====================================================
                        // VALIDAR MISMA HOJA
                        //
                        // Si no corresponde:
                        //
                        // volver directamente al InputBox.
                        // =====================================================

                        hojaAuxiliar =
                            (Excel.Worksheet)
                            celdaInicioAux.Worksheet;

                        if (!string.Equals(
                                hojaAuxiliar.Name,
                                wsActual.Name,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        int columnaTemporal =
                            celdaInicioAux.Column;

                        // =====================================================
                        // COMPROBAR ESPACIO PARA 3 COLUMNAS
                        // =====================================================

                        if (columnaTemporal +
                            TotalColumnasAuxiliares - 1 >
                            MaxColumnasExcel)
                        {
                            continue;
                        }

                        // =====================================================
                        // EVITAR TRASLAPE
                        //
                        // No muestra advertencia.
                        //
                        // Si existe traslape vuelve directamente al selector.
                        // =====================================================

                        if (ColumnasSeTraslapan(
                                rangoCapturado,
                                columnaTemporal,
                                columnaTemporal + 2))
                        {
                            continue;
                        }

                        // =====================================================
                        // SELECCIÓN CORRECTA
                        // =====================================================

                        columnaAuxiliarInicial =
                            columnaTemporal;

                        auxiliarSeleccionado =
                            true;
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(
                            hojaAuxiliar);

                        ExcelHelper.LiberarCom(
                            celdaInicioAux);

                        celdaInicioAux =
                            null;

                        ExcelHelper.LiberarCom(
                            seleccionAuxiliar);

                        seleccionAuxiliar =
                            null;
                    }
                }

                // =============================================================
                // OBTENER LAS 3 COLUMNAS AUXILIARES
                //
                // Ejemplo:
                //
                // AK = Longitud
                // AL = Solo dígitos
                // AM = Válido
                // =============================================================

                string colLongitud =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial);

                string colSoloDigitos =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 1);

                string colValido =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 2);

                // =============================================================
                // FASE 3:
                // ENCABEZADOS AUXILIARES
                // =============================================================

                etapaActual =
                    "Inserción de encabezados auxiliares";

                excelApp.ScreenUpdating =
                    false;

                int filaInicioGeneral =
                    ObtenerFilaMinima(
                        rangoCapturado);

                if (filaInicioGeneral <= 2)
                {
                    return CrearError(
                        "No existen dos filas superiores disponibles para colocar " +
                        "el título de la validación y los encabezados de las fórmulas auxiliares.");
                }

                int filaTitulos =
                    filaInicioGeneral - 1;

                int filaTituloValidacion =
                    filaInicioGeneral - 2;

                Excel.Range tituloValidacion = null;
                Excel.Range tituloLongitud = null;
                Excel.Range tituloSoloDigitos = null;
                Excel.Range tituloValido = null;

                try
                {
                    // ---------------------------------------------------------
                    // TÍTULO GENERAL DE LA VALIDACIÓN
                    //
                    // Combina las 3 columnas auxiliares:
                    //
                    // Longitud | Solo dígitos | Válido
                    // ---------------------------------------------------------

                    tituloValidacion =
                        wsActual.Range[
                            colLongitud +
                            filaTituloValidacion.ToString(),
                            colValido +
                            filaTituloValidacion.ToString()];

                    try
                    {
                        tituloValidacion.UnMerge();
                    }
                    catch
                    {
                        // No crítico.
                    }

                    tituloValidacion.Merge(
                        Type.Missing);

                    tituloValidacion.Value2 =
                        "SAV_Número_telefónico";

                    FormatearTitulo(
                        tituloValidacion);

                    // ---------------------------------------------------------
                    // ENCABEZADOS AUXILIARES
                    // ---------------------------------------------------------

                    tituloLongitud =
                        wsActual.Range[
                            colLongitud +
                            filaTitulos.ToString()];

                    tituloSoloDigitos =
                        wsActual.Range[
                            colSoloDigitos +
                            filaTitulos.ToString()];

                    tituloValido =
                        wsActual.Range[
                            colValido +
                            filaTitulos.ToString()];

                    tituloLongitud.Value2 =
                        "SAV_Longitud";

                    tituloSoloDigitos.Value2 =
                        "SAV_Solo_dígitos";

                    tituloValido.Value2 =
                        "SAV_Válido";

                    FormatearTitulo(
                        tituloLongitud);

                    FormatearTitulo(
                        tituloSoloDigitos);

                    FormatearTitulo(
                        tituloValido);
                }
                finally
                {
                    ExcelHelper.LiberarCom(
                        tituloValido);

                    ExcelHelper.LiberarCom(
                        tituloSoloDigitos);

                    ExcelHelper.LiberarCom(
                        tituloLongitud);

                    ExcelHelper.LiberarCom(
                        tituloValidacion);
                }

                // =============================================================
                // FASE 4:
                // PROCESAMIENTO MASIVO
                //
                // CON DATA VALIDATION
                // CON FORMATO CONDICIONAL
                // SIN MENSAJE CONJUNTO.
                // =============================================================

                etapaActual =
                    "Procesamiento de números telefónicos";

                int totalTelefonosProcesados =
                    0;

                // -------------------------------------------------------------
                // COLORES DEL FORMATO CONDICIONAL DE ERROR
                //
                // Mismo esquema visual utilizado en ValidacionHorarioDV:
                //
                // Fondo rojo claro
                // Texto rojo oscuro
                // -------------------------------------------------------------

                int colorFondoError =
                    System.Drawing.ColorTranslator.ToOle(
                        System.Drawing.Color.FromArgb(
                            255,
                            199,
                            206));

                int colorTextoError =
                    System.Drawing.ColorTranslator.ToOle(
                        System.Drawing.Color.FromArgb(
                            156,
                            0,
                            6));

                for (int j = 1;
                     j <= totalAreas;
                     j++)
                {
                    Excel.Range area = null;

                    try
                    {
                        area =
                            (Excel.Range)
                            rangoCapturado.Areas[j];

                        /*
						 * La Data Validation previa NO se elimina de forma
						 * masiva sobre "area", porque el rango puede contener
						 * varias celdas combinadas independientes.
						 *
						 * La eliminación se realiza más adelante, fila por fila,
						 * directamente sobre mergeTelefono.
						 */

                        int filaInicio =
                            area.Row;

                        // =====================================================
                        // PROCESAR FILA POR FILA
                        // =====================================================

                        for (int filaOffset = 0;
                             filaOffset <
                             area.Rows.Count;
                             filaOffset++)
                        {
                            Excel.Range celdaTelefono = null;
                            Excel.Range mergeTelefono = null;

                            Excel.Range auxLongitud = null;
                            Excel.Range auxSoloDigitos = null;
                            Excel.Range auxValido = null;

                            Excel.Validation validacion = null;

                            Excel.FormatConditions formatosCondicionales = null;
                            Excel.FormatCondition formatoError = null;

                            try
                            {
                                int fila =
                                    filaInicio +
                                    filaOffset;

                                // -------------------------------------------------
                                // CELDA DEL NÚMERO TELEFÓNICO
                                // -------------------------------------------------

                                celdaTelefono =
                                    wsActual.Range[
                                        NumeroColumnaALetra(
                                            area.Column) +
                                        fila.ToString()];

                                mergeTelefono =
                                    (Excel.Range)
                                    celdaTelefono.MergeArea;

                                string dirTelefonoRel =
                                    celdaTelefono.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                // -------------------------------------------------
                                // AUXILIARES
                                // -------------------------------------------------

                                auxLongitud =
                                    wsActual.Range[
                                        colLongitud +
                                        fila.ToString()];

                                auxSoloDigitos =
                                    wsActual.Range[
                                        colSoloDigitos +
                                        fila.ToString()];

                                auxValido =
                                    wsActual.Range[
                                        colValido +
                                        fila.ToString()];

                                string dirLongitudRel =
                                    auxLongitud.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                string dirSoloDigitosRel =
                                    auxSoloDigitos.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                string dirValidoAbs =
                                    auxValido.get_Address(
                                        true,
                                        true,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                // -------------------------------------------------
                                // INSERTAR FÓRMULAS AUXILIARES
                                // -------------------------------------------------

                                auxLongitud.Formula =
                                    ConstruirFormulaLongitud(
                                        dirTelefonoRel);

                                auxSoloDigitos.Formula =
                                    ConstruirFormulaSoloDigitos(
                                        dirTelefonoRel);

                                auxValido.Formula =
                                    ConstruirFormulaValido(
                                        dirTelefonoRel,
                                        dirLongitudRel,
                                        dirSoloDigitosRel);

                                // -------------------------------------------------
                                // FORMATO DE AUXILIARES
                                // -------------------------------------------------

                                FormatearAuxiliar(
                                    auxLongitud);

                                FormatearAuxiliar(
                                    auxSoloDigitos);

                                FormatearAuxiliar(
                                    auxValido);

                                auxValido.Font.Bold =
                                    true;

                                // -------------------------------------------------
                                // FORMATO NUMÉRICO DE AUXILIARES
                                //
                                // Válido y Solo dígitos se muestran como enteros
                                // 0 o 1. Longitud también se muestra como entero.
                                // -------------------------------------------------

                                auxLongitud.NumberFormat =
                                    "0";

                                auxSoloDigitos.NumberFormat =
                                    "0";

                                auxValido.NumberFormat =
                                    "0";

                                // -------------------------------------------------
                                // EL TELÉFONO SE MANEJA COMO TEXTO
                                //
                                // Esto también permite conservar un cero inicial
                                // en caso de ser capturado.
                                // -------------------------------------------------

                                mergeTelefono.NumberFormat =
                                    "@";

                                // -------------------------------------------------
                                // IGNORAR "NÚMERO ALMACENADO COMO TEXTO"
                                //
                                // El teléfono se conserva como texto para respetar
                                // exactamente los 10 dígitos y posibles ceros
                                // iniciales, pero no queremos que Excel muestre
                                // el triángulo verde ni el menú de comprobación
                                // "Número almacenado como texto".
                                //
                                // Se desactiva únicamente para esta celda/rango,
                                // sin modificar las opciones globales de Excel.
                                // -------------------------------------------------

                                IgnorarErrorNumeroComoTexto(
                                    celdaTelefono);

                                // =================================================
                                // ELIMINAR DATA VALIDATION DE ESA CELDA
                                // =================================================

                                try
                                {
                                    validacion =
                                        mergeTelefono.Validation;

                                    validacion.Delete();
                                }
                                catch
                                {
                                    // Sin validación anterior.
                                }
                                finally
                                {
                                    ExcelHelper.LiberarCom(
                                        validacion);

                                    validacion =
                                        null;
                                }

                                // =================================================
                                // CREAR DATA VALIDATION
                                //
                                // El dato se acepta únicamente cuando:
                                //
                                // Válido = 1
                                // =================================================

                                validacion =
                                    mergeTelefono.Validation;

                                string formulaDV =
                                    ExcelHelper.TraducirFormulaLocal(
                                        wsActual,
                                        "=" +
                                        dirValidoAbs +
                                        "=1",
                                        fila);

                                validacion.Add(
                                    Excel.XlDVType.xlValidateCustom,
                                    Excel.XlDVAlertStyle.xlValidAlertStop,
                                    Type.Missing,
                                    formulaDV,
                                    Type.Missing);

                                validacion.IgnoreBlank =
                                    true;

                                validacion.ShowError =
                                    true;

                                validacion.ErrorTitle =
                                    "Número telefónico inválido";

                                /*
								 * Excel limita la longitud del mensaje
								 * de Validación de datos.
								 *
								 * Además de indicar el formato, se recuerda
								 * que el dato debe corresponder a un medio
								 * institucional de contacto.
								 */

                                validacion.ErrorMessage =
                                    "El número telefónico de contacto debe registrarse con diez dígitos.\n\n" +

                                    "Capture solo números, sin espacios, guiones, paréntesis ni prefijo +52.\n\n" +

                                    "Ejemplos válidos:\n" +
                                    "• 5512345678\n" +
                                    "• 8123456789\n" +
                                    "• 9991234567";

                                // -------------------------------------------------
                                // LIMPIEZA DE FORMATOS CONDICIONALES PREVIOS
                                //
                                // Resultado de la Auditoría de coexistencia:
                                //
                                // SÍ  = conservar las reglas visuales anteriores.
                                // NO  = eliminar las reglas visuales anteriores.
                                //
                                // La limpieza se hace fila por fila sobre
                                // MergeArea para respetar celdas combinadas.
                                // -------------------------------------------------

                                if (limpiarFormatos)
                                {
                                    Excel.FormatConditions formatosALimpiar =
                                        null;

                                    try
                                    {
                                        formatosALimpiar =
                                            mergeTelefono.FormatConditions;

                                        if (formatosALimpiar != null)
                                        {
                                            formatosALimpiar.Delete();
                                        }
                                    }
                                    catch
                                    {
                                        // No crítico:
                                        // no había formatos condicionales previos.
                                    }
                                    finally
                                    {
                                        ExcelHelper.LiberarCom(
                                            formatosALimpiar);
                                    }
                                }

                                // -------------------------------------------------
                                // FORMATO CONDICIONAL
                                //
                                // Se resalta la celda del número telefónico cuando:
                                //
                                // 1. La celda NO está vacía.
                                // 2. La columna auxiliar Válido = 0.
                                //
                                // Esto complementa la Data Validation:
                                //
                                // - Data Validation impide nuevas capturas incorrectas.
                                // - Formato condicional identifica visualmente teléfonos
                                //   incorrectos que ya existían o que fueron pegados.
                                // -------------------------------------------------

                                string formulaFCEng =
                                    "=AND(" +
                                    dirTelefonoRel +
                                    "<>\"\"," +
                                    dirValidoAbs +
                                    "=0)";

                                string formulaFCLocal =
                                    ExcelHelper.TraducirFormulaLocal(
                                        wsActual,
                                        formulaFCEng,
                                        fila);

                                formatosCondicionales =
                                    mergeTelefono.FormatConditions;

                                formatoError =
                                    (Excel.FormatCondition)
                                    formatosCondicionales.Add(
                                        Excel.XlFormatConditionType.xlExpression,
                                        Type.Missing,
                                        formulaFCLocal);

                                formatoError.Interior.Color =
                                    colorFondoError;

                                formatoError.Font.Color =
                                    colorTextoError;

                                formatoError.StopIfTrue =
                                    false;

                                // Prioridad alta para que el error sea visible incluso
                                // cuando existan otros formatos condicionales.
                                try
                                {
                                    formatoError.SetFirstPriority();
                                }
                                catch
                                {
                                    // No crítico.
                                }

                                totalTelefonosProcesados++;
                            }
                            finally
                            {
                                ExcelHelper.LiberarCom(
                                    formatoError);

                                ExcelHelper.LiberarCom(
                                    formatosCondicionales);

                                ExcelHelper.LiberarCom(
                                    validacion);

                                ExcelHelper.LiberarCom(
                                    auxValido);

                                ExcelHelper.LiberarCom(
                                    auxSoloDigitos);

                                ExcelHelper.LiberarCom(
                                    auxLongitud);

                                ExcelHelper.LiberarCom(
                                    mergeTelefono);

                                ExcelHelper.LiberarCom(
                                    celdaTelefono);
                            }
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(
                            area);
                    }
                }

                // =============================================================
                // FASE 5:
                // REGISTRO EN BITÁCORA
                // =============================================================

                etapaActual =
                    "Registro en bitácora";

                string rangoAuxiliares =
                    colLongitud +
                    ":" +
                    colValido;

                string estadoAuditoria;

                if (!tieneFormatosPrevios &&
                    !tieneValidacionPrevia)
                {
                    estadoAuditoria =
                        "No se detectaron configuraciones previas.";
                }
                else if (limpiarFormatos)
                {
                    estadoAuditoria =
                        "Se eliminaron los Formatos Condicionales previos.";
                }
                else
                {
                    estadoAuditoria =
                        "Se conservaron los Formatos Condicionales previos.";
                }

                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    "Validación Número telefónico",
                    rangoCapturado.Address.Replace(
                        "$",
                        ""),
                    "Data Validation + Formato Condicional + " +
                    "3 Fórmulas Auxiliares (" +
                    rangoAuxiliares +
                    "). Auditoría de coexistencia: " +
                    estadoAuditoria
                );

                // =============================================================
                // RESTAURAR SELECCIÓN ORIGINAL
                // =============================================================

                try
                {
                    wsActual.Activate();
                    rangoCapturado.Select();
                }
                catch
                {
                    // No crítico.
                }

                return new ResultadoValidacion
                {
                    Exito = true,

                    Mensaje =
                        "Validación de Número telefónico aplicada correctamente.\n\n" +

                        "• Números telefónicos procesados: " +
                        totalTelefonosProcesados.ToString() +
                        ".\n" +

                        "• Fórmulas auxiliares: " +
                        rangoAuxiliares +
                        ".\n" +

                        "• Título de la validación: fila " +
                        filaTituloValidacion.ToString() +
                        ".\n" +

                        "• Títulos auxiliares: fila " +
                        filaTitulos.ToString() +
                        ".\n" +

                        "• Válido: 1 = Correcto / 0 = Error.\n" +
                        "• Se aplicó formato condicional para identificar teléfonos incorrectos.\n" +
                        "• Auditoría de coexistencia: " +
                        estadoAuditoria,

                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion
                {
                    Exito = false,

                    Mensaje =
                        "Error crítico al aplicar la Validación de Número telefónico.\n\n" +

                        "Etapa: " +
                        etapaActual +
                        "\n\n" +

                        "Tipo: " +
                        ex.GetType().FullName +
                        "\n\n" +

                        "Detalle: " +
                        ex.Message,

                    AlertaInyectada = false
                };
            }
            finally
            {
                ExcelHelper.LiberarCom(
                    celdaInicioAux);

                ExcelHelper.LiberarCom(
                    seleccionAuxiliar);

                ExcelHelper.LiberarCom(
                    wsActual);

                if (excelApp != null)
                {
                    excelApp.ScreenUpdating =
                        true;
                }
            }
        }

        // =====================================================================
        // AUXILIAR 1:
        // LONGITUD
        //
        // Ejemplo:
        //
        // 5512345678
        //
        // Resultado:
        //
        // 10
        //
        // Si está vacío:
        //
        // ""
        // =====================================================================

        private static string ConstruirFormulaLongitud(
            string celda)
        {
            return
                "=IF(" +

                celda +
                "=\"\"," +

                "\"\"," +

                "LEN(" +
                celda +
                "&\"\")" +

                ")";
        }

        // =====================================================================
        // AUXILIAR 2:
        // SOLO DÍGITOS
        //
        // 1 = TODOS LOS CARACTERES SON 0-9
        // 0 = EXISTE ALGÚN CARÁCTER NO NUMÉRICO
        //
        // Ejemplos:
        //
        // 5512345678   -> 1
        // 55 12345678  -> 0
        // 55-1234-5678 -> 0
        // +525512345678 -> 0
        // ABC1234567   -> 0
        //
        // Vacío:
        //
        // ""
        // =====================================================================

        private static string ConstruirFormulaSoloDigitos(
            string celda)
        {
            string texto =
                celda +
                "&\"\"";

            return
                "=IF(" +

                celda +
                "=\"\"," +

                "\"\"," +

                "IFERROR(" +

                "IF(" +

                "SUMPRODUCT(" +

                "--ISNUMBER(" +

                "FIND(" +

                "MID(" +

                texto +
                "," +

                "ROW(" +

                "INDIRECT(" +

                "\"1:\"&MAX(1,LEN(" +
                texto +
                "))" +

                ")" +

                ")," +

                "1" +

                ")," +

                "\"0123456789\"" +

                ")" +

                ")" +

                ")=LEN(" +
                texto +
                ")," +

                "1," +

                "0" +

                ")," +

                "0" +

                ")" +

                ")";
        }

        // =====================================================================
        // AUXILIAR 3:
        // VÁLIDO
        //
        // 1 = CORRECTO
        // 0 = ERROR
        //
        // REGLAS:
        //
        // - Vacío = 1
        // - Longitud = 10
        // - Solo dígitos = 1
        //
        // La obligatoriedad corresponde a ValidacionBlancos.
        // =====================================================================

        private static string ConstruirFormulaValido(
            string telefono,
            string longitud,
            string soloDigitos)
        {
            return
                "=IF(" +

                telefono +
                "=\"\"," +

                "1," +

                "IFERROR(" +

                "IF(" +

                "AND(" +

                longitud +
                "=10," +

                soloDigitos +
                "=1" +

                ")," +

                "1," +

                "0" +

                ")," +

                "0" +

                ")" +

                ")";
        }

        // =====================================================================
        // OBTENER PRIMERA CELDA
        //
        // MISMA TÉCNICA DE ValidacionCorreoElectronicoDV.cs
        // =====================================================================

        private static Excel.Range ObtenerPrimeraCelda(
            Excel.Range rango)
        {
            if (rango == null)
            {
                return null;
            }

            Excel.Range celdas =
                null;

            object primeraObj =
                null;

            try
            {
                celdas =
                    rango.Cells;

                primeraObj =
                    celdas.Item[1, 1];

                Excel.Range primeraCelda =
                    primeraObj as Excel.Range;

                if (primeraCelda != null)
                {
                    /*
					 * El objeto será liberado por el método llamador.
					 */

                    primeraObj =
                        null;

                    return primeraCelda;
                }

                return null;
            }
            finally
            {
                if (primeraObj != null)
                {
                    ExcelHelper.LiberarCom(
                        primeraObj);
                }

                ExcelHelper.LiberarCom(
                    celdas);
            }
        }

        // =====================================================================
        // IGNORAR COMPROBACIÓN "NÚMERO ALMACENADO COMO TEXTO"
        //
        // El teléfono se almacena intencionalmente como texto para conservar
        // los 10 dígitos tal como fueron capturados. Esta rutina deshabilita
        // únicamente el aviso xlNumberAsText para la celda correspondiente,
        // sin cambiar ErrorCheckingOptions.NumberAsText a nivel de Excel.
        // =====================================================================

        private static void IgnorarErrorNumeroComoTexto(
            Excel.Range celda)
        {
            if (celda == null)
            {
                return;
            }

            Excel.Errors errores = null;
            Excel.Error errorNumeroComoTexto = null;

            try
            {
                errores =
                    celda.Errors;

                if (errores == null)
                {
                    return;
                }

                errorNumeroComoTexto =
                    errores.Item[
                        Excel.XlErrorChecks.xlNumberAsText];

                if (errorNumeroComoTexto != null)
                {
                    errorNumeroComoTexto.Ignore =
                        true;
                }
            }
            catch
            {
                // No crítico.
                //
                // Si Excel no expone la comprobación para una celda vacía,
                // la validación principal continúa funcionando normalmente.
            }
            finally
            {
                ExcelHelper.LiberarCom(
                    errorNumeroComoTexto);

                ExcelHelper.LiberarCom(
                    errores);
            }
        }

        // =====================================================================
        // FORMATO DE TÍTULOS
        // =====================================================================

        private static void FormatearTitulo(
            Excel.Range celda)
        {
            if (celda == null)
            {
                return;
            }

            celda.Font.Name =
                "Arial";

            celda.Font.Size =
                9;

            celda.Font.Bold =
                true;

            celda.HorizontalAlignment =
                Excel.XlHAlign.xlHAlignCenter;

            celda.VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;

            celda.WrapText =
                true;
        }

        // =====================================================================
        // FORMATO DE AUXILIARES
        // =====================================================================

        private static void FormatearAuxiliar(
            Excel.Range celda)
        {
            if (celda == null)
            {
                return;
            }

            celda.Font.Name =
                "Arial";

            celda.Font.Size =
                9;

            celda.HorizontalAlignment =
                Excel.XlHAlign.xlHAlignCenter;

            celda.VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;
        }

        // =====================================================================
        // OBTENER FILA INICIAL
        // =====================================================================

        private static int ObtenerFilaMinima(
            Excel.Range rango)
        {
            int filaMinima =
                int.MaxValue;

            int totalAreas =
                rango.Areas.Count;

            for (int i = 1;
                 i <= totalAreas;
                 i++)
            {
                Excel.Range area = null;

                try
                {
                    area =
                        (Excel.Range)
                        rango.Areas[i];

                    if (area.Row <
                        filaMinima)
                    {
                        filaMinima =
                            area.Row;
                    }
                }
                finally
                {
                    ExcelHelper.LiberarCom(
                        area);
                }
            }

            return filaMinima;
        }

        // =====================================================================
        // COMPROBAR TRASLAPE
        //
        // Igual que en ValidacionCorreoElectronicoDV:
        //
        // si existe traslape no muestra MessageBox,
        // simplemente vuelve al selector.
        // =====================================================================

        private static bool ColumnasSeTraslapan(
            Excel.Range rango,
            int columnaInicio,
            int columnaFin)
        {
            int totalAreas =
                rango.Areas.Count;

            for (int i = 1;
                 i <= totalAreas;
                 i++)
            {
                Excel.Range area = null;

                try
                {
                    area =
                        (Excel.Range)
                        rango.Areas[i];

                    int inicioArea =
                        area.Column;

                    int finArea =
                        area.Column +
                        area.Columns.Count - 1;

                    if (columnaInicio <=
                            finArea &&
                        columnaFin >=
                            inicioArea)
                    {
                        return true;
                    }
                }
                finally
                {
                    ExcelHelper.LiberarCom(
                        area);
                }
            }

            return false;
        }

        // =====================================================================
        // NÚMERO DE COLUMNA -> LETRA
        // =====================================================================

        private static string NumeroColumnaALetra(
            int numeroColumna)
        {
            if (numeroColumna <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    "numeroColumna");
            }

            string resultado =
                "";

            while (numeroColumna > 0)
            {
                numeroColumna--;

                int residuo =
                    numeroColumna % 26;

                resultado =
                    ((char)('A' + residuo)).ToString() +
                    resultado;

                numeroColumna =
                    numeroColumna / 26;
            }

            return resultado;
        }

        // =====================================================================
        // DTO DE ERROR
        // =====================================================================

        private static ResultadoValidacion CrearError(
            string mensaje)
        {
            return new ResultadoValidacion
            {
                Exito = false,
                Mensaje = mensaje,
                AlertaInyectada = false
            };
        }
    }
}