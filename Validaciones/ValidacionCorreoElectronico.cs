using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core;

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionCorreoElectronico : IValidacionExcel
    {
        // =====================================================================
        // CONSTANTES
        // =====================================================================

        private const int TotalColumnasAuxiliares = 5;
        private const int MaxColumnasExcel = 16384; // XFD

        // =====================================================================
        // MÉTODO PRINCIPAL
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
                // VALIDACIONES GENERALES
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
                // La Data Validation de Correo electrónico reemplaza la
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
                            "y aplicar únicamente la validación de Correo electrónico.\n" +
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
                // VALIDAR RANGO CAPTURADO
                // =============================================================

                etapaActual =
                    "Validación del rango de Correo electrónico";

                int totalAreas =
                    rangoCapturado.Areas.Count;

                if (totalAreas <= 0)
                {
                    return CrearError(
                        "El rango seleccionado no contiene celdas válidas.");
                }

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
                                "el campo 'Correo electrónico de contacto'.");
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
                // DETECTAR PREGUNTAS
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
                // SELECCIÓN DE COLUMNA AUXILIAR
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
                            "auxiliares de la Validación de Correo electrónico.\n\n" +

                            "Considera que el sistema requerirá 5 columnas consecutivas " +
                            "y espacio libre hacia abajo proporcional al número de filas " +
                            "que seleccionaste.",

                            "SAVCNG - Selección de Fórmula Auxiliar",

                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            8);

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
                        // -----------------------------------------------------
                        // INTENTAR OBTENER RANGE DIRECTAMENTE
                        // -----------------------------------------------------

                        seleccionAuxiliar =
                            resAuxiliar as Excel.Range;

                        // -----------------------------------------------------
                        // RESPALDO PARA CUANDO EXCEL DEVUELVE Int32
                        // -----------------------------------------------------

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

                        if (seleccionAuxiliar == null)
                        {
                            continue;
                        }

                        celdaInicioAux =
                            ObtenerPrimeraCelda(
                                seleccionAuxiliar);

                        if (celdaInicioAux == null)
                        {
                            continue;
                        }

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

                        // -----------------------------------------------------
                        // ESPACIO HASTA XFD
                        // -----------------------------------------------------

                        if (columnaTemporal +
                            TotalColumnasAuxiliares - 1 >
                            MaxColumnasExcel)
                        {
                            continue;
                        }

                        // -----------------------------------------------------
                        // EVITAR TRASLAPE CON CORREO
                        // -----------------------------------------------------

                        if (ColumnasSeTraslapan(
                                rangoCapturado,
                                columnaTemporal,
                                columnaTemporal + 4))
                        {
                            continue;
                        }

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
                // COLUMNAS AUXILIARES
                // =============================================================

                string colUsuario =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial);

                string colArroba =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 1);

                string colSLD =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 2);

                string colTLD =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 3);

                string colValido =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 4);

                // =============================================================
                // FASE 3:
                // TÍTULO Y ENCABEZADOS AUXILIARES
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

                Excel.Range tituloUsuario = null;
                Excel.Range tituloArroba = null;
                Excel.Range tituloSLD = null;
                Excel.Range tituloTLD = null;
                Excel.Range tituloValido = null;

                try
                {
                    // ---------------------------------------------------------
                    // TÍTULO GENERAL
                    // ---------------------------------------------------------

                    tituloValidacion =
                        wsActual.Range[
                            colUsuario +
                            filaTituloValidacion.ToString(),
                            colValido +
                            filaTituloValidacion.ToString()];

                    /*
					 * Si las cinco celdas ya forman exactamente un rango
					 * combinado, no es necesario volver a combinarlo.
					 *
					 * Si no están combinadas, primero se descombinan posibles
					 * combinaciones previas contenidas completamente en el
					 * rango y después se combinan las cinco celdas.
					 */

                    bool tituloYaCombinado =
                        false;

                    try
                    {
                        object mergeCells =
                            tituloValidacion.MergeCells;

                        if (mergeCells is bool)
                        {
                            tituloYaCombinado =
                                (bool)mergeCells;
                        }
                    }
                    catch
                    {
                        tituloYaCombinado =
                            false;
                    }

                    if (!tituloYaCombinado)
                    {
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
                    }

                    tituloValidacion.Value2 =
                        "SAV_Correo_electrónico";

                    FormatearTitulo(
                        tituloValidacion);

                    // ---------------------------------------------------------
                    // ENCABEZADOS AUXILIARES
                    // ---------------------------------------------------------

                    tituloUsuario =
                        wsActual.Range[
                            colUsuario +
                            filaTitulos.ToString()];

                    tituloArroba =
                        wsActual.Range[
                            colArroba +
                            filaTitulos.ToString()];

                    tituloSLD =
                        wsActual.Range[
                            colSLD +
                            filaTitulos.ToString()];

                    tituloTLD =
                        wsActual.Range[
                            colTLD +
                            filaTitulos.ToString()];

                    tituloValido =
                        wsActual.Range[
                            colValido +
                            filaTitulos.ToString()];

                    tituloUsuario.Value2 =
                        "SAV_Usuario";

                    tituloArroba.Value2 =
                        "SAV_@";

                    tituloSLD.Value2 =
                        "SAV_SLD";

                    tituloTLD.Value2 =
                        "SAV_TLD";

                    tituloValido.Value2 =
                        "SAV_Válido";

                    FormatearTitulo(
                        tituloUsuario);

                    FormatearTitulo(
                        tituloArroba);

                    FormatearTitulo(
                        tituloSLD);

                    FormatearTitulo(
                        tituloTLD);

                    FormatearTitulo(
                        tituloValido);
                }
                finally
                {
                    ExcelHelper.LiberarCom(
                        tituloValido);

                    ExcelHelper.LiberarCom(
                        tituloTLD);

                    ExcelHelper.LiberarCom(
                        tituloSLD);

                    ExcelHelper.LiberarCom(
                        tituloArroba);

                    ExcelHelper.LiberarCom(
                        tituloUsuario);

                    ExcelHelper.LiberarCom(
                        tituloValidacion);
                }

                // =============================================================
                // FASE 4:
                // PROCESAMIENTO
                //
                // DATA VALIDATION + FORMATO CONDICIONAL
                // =============================================================

                etapaActual =
                    "Procesamiento de correos electrónicos";

                int totalCorreosProcesados =
                    0;

                // -------------------------------------------------------------
                // COLORES DE ERROR
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
						 * =====================================================
						 * IMPORTANTE
						 * =====================================================
						 *
						 * NO SE UTILIZA:
						 *
						 *	 area.Validation.Delete();
						 *
						 * El rango "area" puede contener varias celdas
						 * combinadas independientes, por ejemplo:
						 *
						 * Z23:AB23
						 * Z24:AB24
						 * Z25:AB25
						 *
						 * Intentar borrar la Validación de datos sobre todo
						 * Z23:AB42 puede provocar el mensaje nativo de Excel:
						 *
						 * "Esta acción no se puede realizar en una celda
						 * combinada."
						 *
						 * La Validación de datos se elimina más adelante,
						 * individualmente sobre cada MergeArea.
						 */

                        int filaInicio =
                            area.Row;

                        int totalFilasArea =
                            area.Rows.Count;

                        // =====================================================
                        // PROCESAR FILA POR FILA
                        // =====================================================

                        for (int filaOffset = 0;
                             filaOffset < totalFilasArea;
                             filaOffset++)
                        {
                            Excel.Range celdaCorreo = null;
                            Excel.Range mergeCorreo = null;

                            Excel.Range auxUsuario = null;
                            Excel.Range auxArroba = null;
                            Excel.Range auxSLD = null;
                            Excel.Range auxTLD = null;
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
                                // CELDA DEL CORREO
                                // -------------------------------------------------

                                celdaCorreo =
                                    wsActual.Range[
                                        NumeroColumnaALetra(
                                            area.Column) +
                                        fila.ToString()];

                                mergeCorreo =
                                    (Excel.Range)
                                    celdaCorreo.MergeArea;

                                string dirCorreoRel =
                                    celdaCorreo.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                // -------------------------------------------------
                                // AUXILIARES
                                // -------------------------------------------------

                                auxUsuario =
                                    wsActual.Range[
                                        colUsuario +
                                        fila.ToString()];

                                auxArroba =
                                    wsActual.Range[
                                        colArroba +
                                        fila.ToString()];

                                auxSLD =
                                    wsActual.Range[
                                        colSLD +
                                        fila.ToString()];

                                auxTLD =
                                    wsActual.Range[
                                        colTLD +
                                        fila.ToString()];

                                auxValido =
                                    wsActual.Range[
                                        colValido +
                                        fila.ToString()];

                                string dirUsuarioRel =
                                    auxUsuario.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                string dirArrobaRel =
                                    auxArroba.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                string dirSLDRel =
                                    auxSLD.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                string dirTLDRel =
                                    auxTLD.get_Address(
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
                                // FÓRMULAS AUXILIARES
                                // -------------------------------------------------

                                auxUsuario.Formula =
                                    ConstruirFormulaUsuario(
                                        dirCorreoRel);

                                auxArroba.Formula =
                                    ConstruirFormulaArroba(
                                        dirCorreoRel);

                                auxSLD.Formula =
                                    ConstruirFormulaSLD(
                                        dirCorreoRel);

                                auxTLD.Formula =
                                    ConstruirFormulaTLD(
                                        dirCorreoRel);

                                auxValido.Formula =
                                    ConstruirFormulaValido(
                                        dirCorreoRel,
                                        dirUsuarioRel,
                                        dirArrobaRel,
                                        dirSLDRel,
                                        dirTLDRel);

                                // -------------------------------------------------
                                // FORMATO AUXILIARES
                                // -------------------------------------------------

                                FormatearAuxiliar(
                                    auxUsuario);

                                FormatearAuxiliar(
                                    auxArroba);

                                FormatearAuxiliar(
                                    auxSLD);

                                FormatearAuxiliar(
                                    auxTLD);

                                FormatearAuxiliar(
                                    auxValido);

                                auxValido.NumberFormat =
                                    "0";

                                auxValido.Font.Bold =
                                    true;

                                // -------------------------------------------------
                                // CORREO COMO TEXTO
                                // -------------------------------------------------

                                mergeCorreo.NumberFormat =
                                    "@";

                                // =================================================
                                // ELIMINAR DATA VALIDATION
                                //
                                // ÚNICAMENTE SOBRE LA CELDA COMBINADA DE ESTA FILA
                                // =================================================

                                try
                                {
                                    validacion =
                                        mergeCorreo.Validation;

                                    validacion.Delete();
                                }
                                catch
                                {
                                    // No existía validación anterior.
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
                                // =================================================

                                validacion =
                                    mergeCorreo.Validation;

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
                                    "Formato de correo inválido";

                                validacion.ErrorMessage =
                                    "El correo electrónico debe registrarse con un formato válido.\n\n" +

                                    "Debe contener usuario, @, dominio (SLD) y extensión (TLD).\n\n" +

                                    "Ejemplos válidos:\n" +
                                    "• contacto@institucion.gob.mx\n" +
                                    "• informacion@empresa.com\n" +
                                    "• servicios@dependencia.mx";

                                // -------------------------------------------------
                                // LIMPIEZA DE FORMATOS CONDICIONALES PREVIOS
                                //
                                // Resultado de la Auditoría de coexistencia:
                                //
                                // SÍ  = conservar las reglas visuales anteriores.
                                // NO  = eliminar las reglas visuales anteriores.
                                //
                                // Se hace fila por fila sobre MergeArea para
                                // respetar las celdas combinadas del cuestionario.
                                // -------------------------------------------------

                                if (limpiarFormatos)
                                {
                                    Excel.FormatConditions formatosALimpiar =
                                        null;

                                    try
                                    {
                                        formatosALimpiar =
                                            mergeCorreo.FormatConditions;

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

                                // =================================================
                                // FORMATO CONDICIONAL
                                //
                                // CORREO NO VACÍO + VÁLIDO = 0
                                // =================================================

                                string formulaFCEng =
                                    "=AND(" +
                                    dirCorreoRel +
                                    "<>\"\"," +
                                    dirValidoAbs +
                                    "=0)";

                                string formulaFCLocal =
                                    ExcelHelper.TraducirFormulaLocal(
                                        wsActual,
                                        formulaFCEng,
                                        fila);

                                formatosCondicionales =
                                    mergeCorreo.FormatConditions;

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

                                try
                                {
                                    formatoError.SetFirstPriority();
                                }
                                catch
                                {
                                    // No crítico.
                                }

                                totalCorreosProcesados++;
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
                                    auxTLD);

                                ExcelHelper.LiberarCom(
                                    auxSLD);

                                ExcelHelper.LiberarCom(
                                    auxArroba);

                                ExcelHelper.LiberarCom(
                                    auxUsuario);

                                ExcelHelper.LiberarCom(
                                    mergeCorreo);

                                ExcelHelper.LiberarCom(
                                    celdaCorreo);
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
                // BITÁCORA
                // =============================================================

                etapaActual =
                    "Registro en bitácora";

                string rangoAuxiliares =
                    colUsuario +
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
                    "Validación Correo electrónico",
                    rangoCapturado.Address.Replace(
                        "$",
                        ""),
                    "Data Validation + Formato Condicional + " +
                    "5 Fórmulas Auxiliares (" +
                    rangoAuxiliares +
                    "). Auditoría de coexistencia: " +
                    estadoAuditoria
                );

                // =============================================================
                // RESTAURAR SELECCIÓN
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
                        "Validación de Correo Electrónico aplicada correctamente.\n\n" +

                        "• Correos procesados: " +
                        totalCorreosProcesados.ToString() +
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

                        "• Se aplicó Data Validation y formato condicional.\n" +

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
                        "Error crítico al aplicar la Validación de Correo Electrónico.\n\n" +

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
        // USUARIO
        // =====================================================================

        private static string ConstruirFormulaUsuario(
            string celda)
        {
            return
                "=IF(" +

                celda +
                "=\"\"," +

                "\"\"," +

                "IFERROR(" +

                "LEFT(" +

                celda +
                "," +

                "FIND(\"@\"," +
                celda +
                ")-1" +

                ")," +

                "\"\"" +

                ")" +

                ")";
        }

        // =====================================================================
        // AUXILIAR 2:
        // ARROBA
        // =====================================================================

        private static string ConstruirFormulaArroba(
            string celda)
        {
            return
                "=IF(" +

                celda +
                "=\"\"," +

                "\"\"," +

                "IF(" +

                "LEN(" +
                celda +
                ")-" +

                "LEN(" +

                "SUBSTITUTE(" +
                celda +
                ",\"@\",\"\"" +

                ")" +

                ")=1," +

                "\"@\"," +

                "\"\"" +

                ")" +

                ")";
        }

        // =====================================================================
        // AUXILIAR 3:
        // SLD
        // =====================================================================

        private static string ConstruirFormulaSLD(
            string celda)
        {
            return
                "=IF(" +

                celda +
                "=\"\"," +

                "\"\"," +

                "IFERROR(" +

                "MID(" +

                celda +
                "," +

                "FIND(\"@\"," +
                celda +
                ")+1," +

                "FIND(\".\"," +
                celda +
                "," +

                "FIND(\"@\"," +
                celda +
                ")+2" +

                ")-" +

                "FIND(\"@\"," +
                celda +
                ")-1" +

                ")," +

                "\"\"" +

                ")" +

                ")";
        }

        // =====================================================================
        // AUXILIAR 4:
        // TLD
        // =====================================================================

        private static string ConstruirFormulaTLD(
            string celda)
        {
            return
                "=IF(" +

                celda +
                "=\"\"," +

                "\"\"," +

                "IFERROR(" +

                "MID(" +

                celda +
                "," +

                "FIND(\".\"," +
                celda +
                "," +

                "FIND(\"@\"," +
                celda +
                ")+2" +

                ")," +

                "LEN(" +
                celda +
                ")" +

                ")," +

                "\"\"" +

                ")" +

                ")";
        }

        // =====================================================================
        // AUXILIAR 5:
        // VÁLIDO
        //
        // 1 = CORRECTO
        // 0 = ERROR
        //
        // VACÍO = 1
        // =====================================================================

        private static string ConstruirFormulaValido(
            string correo,
            string usuario,
            string arroba,
            string sld,
            string tld)
        {
            string caracteresUsuario =
                ConstruirExpresionCaracteresPermitidos(
                    usuario,
                    "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789._%+-");

            string caracteresSLD =
                ConstruirExpresionCaracteresPermitidos(
                    sld,
                    "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-");

            string caracteresTLD =
                ConstruirExpresionCaracteresPermitidos(
                    tld,
                    "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-.");

            // -------------------------------------------------------------
            // TLD SIN PUNTO INICIAL
            // -------------------------------------------------------------

            string tldSinPunto =
                "IF(" +

                "LEFT(" +
                tld +
                ",1)=\".\"," +

                "RIGHT(" +

                tld +
                "," +

                "LEN(" +
                tld +
                ")-1" +

                ")," +

                tld +

                ")";

            // -------------------------------------------------------------
            // ÚLTIMO SEGMENTO DEL TLD
            // -------------------------------------------------------------

            string tldFinal =
                "TRIM(" +

                "RIGHT(" +

                "SUBSTITUTE(" +

                tldSinPunto +
                "," +

                "\".\"," +

                "REPT(\" \",100)" +

                ")," +

                "100" +

                ")" +

                ")";

            string caracteresTldFinal =
                ConstruirExpresionCaracteresPermitidos(
                    tldFinal,
                    "ABCDEFGHIJKLMNOPQRSTUVWXYZ");

            // -------------------------------------------------------------
            // EXACTAMENTE UNA @
            // -------------------------------------------------------------

            string unaArroba =
                "LEN(" +
                correo +
                ")-" +

                "LEN(" +

                "SUBSTITUTE(" +
                correo +
                ",\"@\",\"\"" +

                ")" +

                ")=1";

            // -------------------------------------------------------------
            // USUARIO VÁLIDO
            // -------------------------------------------------------------

            string usuarioValido =
                "AND(" +

                usuario +
                "<>\"\"," +

                "LEN(" +
                usuario +
                ")<=64," +

                "LEFT(" +
                usuario +
                ",1)<>\".\"," +

                "RIGHT(" +
                usuario +
                ",1)<>\".\"," +

                "ISERROR(" +
                "FIND(\"..\"," +
                usuario +
                ")" +
                ")," +

                caracteresUsuario +

                ")";

            // -------------------------------------------------------------
            // SLD VÁLIDO
            // -------------------------------------------------------------

            string sldValido =
                "AND(" +

                sld +
                "<>\"\"," +

                "LEN(" +
                sld +
                ")<=63," +

                "LEFT(" +
                sld +
                ",1)<>\"-\"," +

                "RIGHT(" +
                sld +
                ",1)<>\"-\"," +

                caracteresSLD +

                ")";

            // -------------------------------------------------------------
            // TLD VÁLIDO
            // -------------------------------------------------------------

            string tldValido =
                "AND(" +

                tld +
                "<>\"\"," +

                "LEFT(" +
                tld +
                ",1)=\".\"," +

                "LEN(" +
                tld +
                ")>=3," +

                "RIGHT(" +
                tld +
                ",1)<>\".\"," +

                "ISERROR(" +
                "FIND(\"..\"," +
                tld +
                ")" +
                ")," +

                "ISERROR(" +
                "FIND(\"-.\"," +
                tld +
                ")" +
                ")," +

                "ISERROR(" +
                "FIND(\".-\"," +
                tld +
                ")" +
                ")," +

                caracteresTLD +
                "," +

                "LEN(" +
                tldFinal +
                ")>=2," +

                "LEN(" +
                tldFinal +
                ")<=63," +

                caracteresTldFinal +

                ")";

            // -------------------------------------------------------------
            // ESTRUCTURA GENERAL
            // -------------------------------------------------------------

            string estructuraGeneral =
                "AND(" +

                "LEN(" +
                correo +
                ")<=254," +

                correo +
                "=TRIM(" +
                correo +
                ")," +

                correo +
                "=CLEAN(" +
                correo +
                ")," +

                "ISERROR(" +
                "FIND(\" \"," +
                correo +
                ")" +
                ")," +

                unaArroba +
                "," +

                arroba +
                "=\"@\"" +

                ")";

            string longitudDominio =
                "LEN(" +
                sld +
                ")+LEN(" +
                tld +
                ")<=253";

            // -------------------------------------------------------------
            // RESULTADO
            // -------------------------------------------------------------

            return
                "=IF(" +

                correo +
                "=\"\"," +

                "1," +

                "IFERROR(" +

                "IF(" +

                "AND(" +

                estructuraGeneral +
                "," +

                usuarioValido +
                "," +

                sldValido +
                "," +

                tldValido +
                "," +

                longitudDominio +

                ")," +

                "1," +

                "0" +

                ")," +

                "0" +

                ")" +

                ")";
        }

        // =====================================================================
        // CARACTERES PERMITIDOS
        // =====================================================================

        private static string ConstruirExpresionCaracteresPermitidos(
            string referencia,
            string permitidos)
        {
            return
                "SUMPRODUCT(" +

                "--ISNUMBER(" +

                "FIND(" +

                "MID(" +

                "UPPER(" +
                referencia +
                ")," +

                "ROW(" +

                "INDIRECT(" +

                "\"1:\"&MAX(1,LEN(" +
                referencia +
                "))" +

                ")" +

                ")," +

                "1" +

                ")," +

                "\"" +
                permitidos +
                "\"" +

                ")" +

                ")" +

                ")=LEN(" +
                referencia +
                ")";
        }

        // =====================================================================
        // OBTENER PRIMERA CELDA
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
        // FORMATO AUXILIAR
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
        // FILA MÍNIMA
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
        // COLUMNA NUMÉRICA -> LETRA
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
        // CREAR ERROR
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