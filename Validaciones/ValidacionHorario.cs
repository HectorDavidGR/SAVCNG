using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core;

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionHorario : IValidacionExcel
    {
        // =====================================================================
        // CONSTANTES
        // =====================================================================

        private const int TotalColumnasAuxiliares = 7;
        private const int ColumnaMinimaAuxiliar = 32; // AF
        private const int MaxColumnasExcel = 16384;   // XFD

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

                            "NOTA: La Validación de Horario reemplazará la regla de " +
                            "Validación de datos existente, pero...\n\n" +

                            "¿Deseas CONSERVAR los colores, alertas o bloqueos " +
                            "(Formatos Condicionales) aplicados previamente?\n\n" +

                            "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                            "NO = BORRAR todo el historial de Formatos Condicionales " +
                            "y aplicar únicamente la validación de Horario.\n" +
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
                    "Validación del rango de Horario de atención";

                int totalAreas =
                    rangoCapturado.Areas.Count;

                if (totalAreas <= 0)
                {
                    return CrearError(
                        "El rango seleccionado no contiene celdas válidas.");
                }

                // -------------------------------------------------------------
                // VALIDAR LA ESTRUCTURA DEL RANGO
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
                            (Excel.Range)
                            area.Cells[1, 1];

                        mergeArea =
                            (Excel.Range)
                            primeraCelda.MergeArea;

                        int anchoCapturado =
                            area.Columns.Count;

                        int anchoCombinado =
                            mergeArea.Columns.Count;

                        /*
						 * Ejemplo esperado:
						 *
						 * W23:Y42
						 *
						 * Si W:Y forman una celda combinada por fila,
						 * el ancho del rango debe coincidir con el ancho
						 * del MergeArea.
						 */

                        if (anchoCapturado !=
                            anchoCombinado)
                        {
                            return CrearError(
                                "El rango seleccionado debe contener únicamente " +
                                "el campo 'Horario de atención (formato de 24 horas)'.");
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

                object resAuxiliar =
                    excelApp.InputBox(
                        "Indica a partir de qué columna libre deseas colocar las fórmulas " +
                        "auxiliares de la Validación de Horario.\n\n" +

                        "Considera que el sistema requerirá 7 columnas consecutivas " +
                        "y espacio libre hacia abajo proporcional al número de filas " +
                        "que seleccionaste.",

                        "SAVCNG - Selección de Fórmula Auxiliar",

                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        8);

                // -------------------------------------------------------------
                // CANCELAR
                // -------------------------------------------------------------

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

                int columnaAuxiliarInicial =
                    0;

                try
                {
                    seleccionAuxiliar =
                        (Excel.Range)
                        resAuxiliar;

                    celdaInicioAux =
                        (Excel.Range)
                        seleccionAuxiliar.Cells[1, 1];

                    columnaAuxiliarInicial =
                        celdaInicioAux.Column;

                    // ---------------------------------------------------------
                    // VALIDAR QUE LA CELDA ESTÉ EN LA MISMA HOJA
                    // ---------------------------------------------------------

                    Excel.Worksheet hojaAuxiliar = null;

                    try
                    {
                        hojaAuxiliar =
                            (Excel.Worksheet)
                            celdaInicioAux.Worksheet;

                        if (!string.Equals(
                                hojaAuxiliar.Name,
                                wsActual.Name,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            return CrearError(
                                "La fórmula auxiliar debe colocarse en la misma hoja " +
                                "que el rango de Horario de atención.");
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(
                            hojaAuxiliar);
                    }

                    // ---------------------------------------------------------
                    // SOLO AF EN ADELANTE
                    // ---------------------------------------------------------

                    if (columnaAuxiliarInicial <
                        ColumnaMinimaAuxiliar)
                    {
                        return CrearError(
                            "Las fórmulas auxiliares deben colocarse " +
                            "a partir de la columna AF.");
                    }

                    // ---------------------------------------------------------
                    // ESPACIO PARA LAS 7 COLUMNAS
                    // ---------------------------------------------------------

                    if (columnaAuxiliarInicial +
                        TotalColumnasAuxiliares - 1 >
                        MaxColumnasExcel)
                    {
                        return CrearError(
                            "No existe espacio suficiente para colocar " +
                            "las 7 columnas auxiliares.");
                    }

                    // ---------------------------------------------------------
                    // EVITAR TRASLAPE CON EL RANGO DE HORARIO
                    // ---------------------------------------------------------

                    if (ColumnasSeTraslapan(
                            rangoCapturado,
                            columnaAuxiliarInicial,
                            columnaAuxiliarInicial + 6))
                    {
                        return CrearError(
                            "Las columnas seleccionadas para las fórmulas auxiliares " +
                            "se traslapan con el rango de Horario de atención.\n\n" +
                            "Selecciona otra columna libre.");
                    }
                }
                finally
                {
                    ExcelHelper.LiberarCom(
                        celdaInicioAux);

                    celdaInicioAux =
                        null;

                    ExcelHelper.LiberarCom(
                        seleccionAuxiliar);

                    seleccionAuxiliar =
                        null;
                }

                // =============================================================
                // OBTENER LAS 7 COLUMNAS AUXILIARES
                // =============================================================

                string colInicio1 =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial);

                string colFin1 =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 1);

                string colDias1 =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 2);

                string colInicio2 =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 3);

                string colFin2 =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 4);

                string colDias2 =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 5);

                string colValido =
                    NumeroColumnaALetra(
                        columnaAuxiliarInicial + 6);

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
                Excel.Range tituloInicio1 = null;
                Excel.Range tituloFin1 = null;
                Excel.Range tituloDias1 = null;

                Excel.Range tituloInicio2 = null;
                Excel.Range tituloFin2 = null;
                Excel.Range tituloDias2 = null;

                Excel.Range tituloValido = null;

                try
                {
                    // ---------------------------------------------------------
                    // TÍTULO GENERAL DE LA VALIDACIÓN
                    //
                    // Se coloca una fila arriba de los 7 encabezados auxiliares
                    // y se combinan las 7 celdas correspondientes.
                    // ---------------------------------------------------------

                    tituloValidacion =
                        wsActual.Range[
                            colInicio1 +
                            filaTituloValidacion.ToString(),
                            colValido +
                            filaTituloValidacion.ToString()];

                    // Si la validación ya se había aplicado previamente,
                    // normalizar el rango antes de volver a combinarlo.
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
                        "SAV_Horario_atención";

                    FormatearTitulo(
                        tituloValidacion);

                    tituloInicio1 =
                        wsActual.Range[
                            colInicio1 +
                            filaTitulos.ToString()];

                    tituloFin1 =
                        wsActual.Range[
                            colFin1 +
                            filaTitulos.ToString()];

                    tituloDias1 =
                        wsActual.Range[
                            colDias1 +
                            filaTitulos.ToString()];

                    tituloInicio2 =
                        wsActual.Range[
                            colInicio2 +
                            filaTitulos.ToString()];

                    tituloFin2 =
                        wsActual.Range[
                            colFin2 +
                            filaTitulos.ToString()];

                    tituloDias2 =
                        wsActual.Range[
                            colDias2 +
                            filaTitulos.ToString()];

                    tituloValido =
                        wsActual.Range[
                            colValido +
                            filaTitulos.ToString()];

                    tituloInicio1.Value2 =
                        "SAV_Inicio_1";

                    tituloFin1.Value2 =
                        "SAV_Fin_1";

                    tituloDias1.Value2 =
                        "SAV_Días_1";

                    tituloInicio2.Value2 =
                        "SAV_Inicio_2";

                    tituloFin2.Value2 =
                        "SAV_Fin_2";

                    tituloDias2.Value2 =
                        "SAV_Días_2";

                    tituloValido.Value2 =
                        "SAV_Válido";

                    FormatearTitulo(
                        tituloInicio1);

                    FormatearTitulo(
                        tituloFin1);

                    FormatearTitulo(
                        tituloDias1);

                    FormatearTitulo(
                        tituloInicio2);

                    FormatearTitulo(
                        tituloFin2);

                    FormatearTitulo(
                        tituloDias2);

                    FormatearTitulo(
                        tituloValido);
                }
                finally
                {
                    ExcelHelper.LiberarCom(
                        tituloValido);

                    ExcelHelper.LiberarCom(
                        tituloDias2);

                    ExcelHelper.LiberarCom(
                        tituloFin2);

                    ExcelHelper.LiberarCom(
                        tituloInicio2);

                    ExcelHelper.LiberarCom(
                        tituloDias1);

                    ExcelHelper.LiberarCom(
                        tituloFin1);

                    ExcelHelper.LiberarCom(
                        tituloInicio1);

                    ExcelHelper.LiberarCom(
                        tituloValidacion);
                }

                // =============================================================
                // FASE 4:
                // PROCESAMIENTO MASIVO POR ÁREAS
                // =============================================================

                etapaActual =
                    "Procesamiento de horarios";

                int totalHorariosProcesados =
                    0;

                // -------------------------------------------------------------
                // COLORES DEL FORMATO CONDICIONAL DE ERROR
                //
                // Mismo esquema visual de Excel para datos incorrectos:
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

                        // -----------------------------------------------------
                        // ELIMINAR VALIDACIÓN DE DATOS ANTERIOR
                        // -----------------------------------------------------

                        try
                        {
                            area.Validation.Delete();
                        }
                        catch
                        {
                            // No existía validación previa.
                        }

                        int filaInicio =
                            area.Row;

                        // =====================================================
                        // FILA POR FILA
                        // =====================================================

                        for (int filaOffset = 0;
                             filaOffset <
                             area.Rows.Count;
                             filaOffset++)
                        {
                            Excel.Range celdaHorario = null;
                            Excel.Range mergeHorario = null;

                            Excel.Range auxInicio1 = null;
                            Excel.Range auxFin1 = null;
                            Excel.Range auxDias1 = null;

                            Excel.Range auxInicio2 = null;
                            Excel.Range auxFin2 = null;
                            Excel.Range auxDias2 = null;

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
                                // CELDA DE HORARIO
                                // -------------------------------------------------

                                celdaHorario =
                                    wsActual.Range[
                                        NumeroColumnaALetra(
                                            area.Column) +
                                        fila.ToString()];

                                mergeHorario =
                                    (Excel.Range)
                                    celdaHorario.MergeArea;

                                string dirHorarioRel =
                                    celdaHorario.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                // -------------------------------------------------
                                // AUXILIARES
                                // -------------------------------------------------

                                auxInicio1 =
                                    wsActual.Range[
                                        colInicio1 +
                                        fila.ToString()];

                                auxFin1 =
                                    wsActual.Range[
                                        colFin1 +
                                        fila.ToString()];

                                auxDias1 =
                                    wsActual.Range[
                                        colDias1 +
                                        fila.ToString()];

                                auxInicio2 =
                                    wsActual.Range[
                                        colInicio2 +
                                        fila.ToString()];

                                auxFin2 =
                                    wsActual.Range[
                                        colFin2 +
                                        fila.ToString()];

                                auxDias2 =
                                    wsActual.Range[
                                        colDias2 +
                                        fila.ToString()];

                                auxValido =
                                    wsActual.Range[
                                        colValido +
                                        fila.ToString()];

                                string dirInicio1Rel =
                                    auxInicio1.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                string dirFin1Rel =
                                    auxFin1.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                string dirDias1Rel =
                                    auxDias1.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                string dirInicio2Rel =
                                    auxInicio2.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                string dirFin2Rel =
                                    auxFin2.get_Address(
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false);

                                string dirDias2Rel =
                                    auxDias2.get_Address(
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

                                auxInicio1.Formula =
                                    ConstruirFormulaInicio1(
                                        dirHorarioRel);

                                auxFin1.Formula =
                                    ConstruirFormulaFin1(
                                        dirHorarioRel);

                                auxDias1.Formula =
                                    ConstruirFormulaDias1(
                                        dirHorarioRel);

                                auxInicio2.Formula =
                                    ConstruirFormulaInicio2(
                                        dirHorarioRel);

                                auxFin2.Formula =
                                    ConstruirFormulaFin2(
                                        dirHorarioRel);

                                auxDias2.Formula =
                                    ConstruirFormulaDias2(
                                        dirHorarioRel);

                                auxValido.Formula =
                                    ConstruirFormulaValido(
                                        dirHorarioRel,
                                        dirInicio1Rel,
                                        dirFin1Rel,
                                        dirDias1Rel,
                                        dirInicio2Rel,
                                        dirFin2Rel,
                                        dirDias2Rel);

                                // -------------------------------------------------
                                // FORMATO DE AUXILIARES
                                // -------------------------------------------------

                                FormatearAuxiliar(
                                    auxInicio1);

                                FormatearAuxiliar(
                                    auxFin1);

                                FormatearAuxiliar(
                                    auxDias1);

                                FormatearAuxiliar(
                                    auxInicio2);

                                FormatearAuxiliar(
                                    auxFin2);

                                FormatearAuxiliar(
                                    auxDias2);

                                FormatearAuxiliar(
                                    auxValido);

                                auxValido.Font.Bold =
                                    true;

                                // -------------------------------------------------
                                // EL HORARIO SE MANEJA COMO TEXTO
                                // -------------------------------------------------

                                mergeHorario.NumberFormat =
                                    "@";

                                // -------------------------------------------------
                                // VALIDACIÓN DE DATOS
                                //
                                // La columna Válido contiene:
                                //
                                // 1 = Correcto
                                // 0 = Error
                                // -------------------------------------------------

                                try
                                {
                                    validacion =
                                        mergeHorario.Validation;

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

                                validacion =
                                    mergeHorario.Validation;

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
                                    "Formato de horario inválido";

                                validacion.ErrorMessage =
                                    "El horario debe registrarse en formato de 24 horas.\n\n" +

                                    "Ejemplos válidos:\n" +
                                    "• 9:00 - 15:00\n" +
                                    "• 00:00 - 24:00\n" +
                                    "• 8:00 - 14:00, 16:00 - 18:00\n" +
                                    "• 9:00 - 18:00 (L-J), 09:00 - 15:00 (V)";

                                // -------------------------------------------------
                                // LIMPIEZA DE FORMATOS CONDICIONALES PREVIOS
                                //
                                // La auditoría de coexistencia permite decidir:
                                //
                                // SÍ  = conservar reglas visuales anteriores.
                                // NO  = eliminar reglas visuales anteriores.
                                //
                                // La eliminación se hace fila por fila sobre
                                // MergeArea para respetar celdas combinadas.
                                // -------------------------------------------------

                                if (limpiarFormatos)
                                {
                                    Excel.FormatConditions formatosALimpiar =
                                        null;

                                    try
                                    {
                                        formatosALimpiar =
                                            mergeHorario.FormatConditions;

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
                                // Se resalta la celda de Horario de atención
                                // cuando:
                                //
                                // 1. La celda NO está vacía.
                                // 2. La columna auxiliar Válido = 0.
                                //
                                // Esto complementa la Data Validation:
                                //
                                // - Data Validation impide nuevas capturas incorrectas.
                                // - Formato condicional identifica visualmente datos
                                //   incorrectos que ya existían o que fueron pegados.
                                // -------------------------------------------------

                                string formulaFCEng =
                                    "=AND(" +
                                    dirHorarioRel +
                                    "<>\"\"," +
                                    dirValidoAbs +
                                    "=0)";

                                string formulaFCLocal =
                                    ExcelHelper.TraducirFormulaLocal(
                                        wsActual,
                                        formulaFCEng,
                                        fila);

                                formatosCondicionales =
                                    mergeHorario.FormatConditions;

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

                                /*
								 * La regla se coloca con prioridad alta para que
								 * el error sea visible aun cuando la celda ya tenga
								 * otros formatos condicionales.
								 */
                                try
                                {
                                    formatoError.SetFirstPriority();
                                }
                                catch
                                {
                                    // No crítico.
                                }

                                totalHorariosProcesados++;
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
                                    auxDias2);

                                ExcelHelper.LiberarCom(
                                    auxFin2);

                                ExcelHelper.LiberarCom(
                                    auxInicio2);

                                ExcelHelper.LiberarCom(
                                    auxDias1);

                                ExcelHelper.LiberarCom(
                                    auxFin1);

                                ExcelHelper.LiberarCom(
                                    auxInicio1);

                                ExcelHelper.LiberarCom(
                                    mergeHorario);

                                ExcelHelper.LiberarCom(
                                    celdaHorario);
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
                    colInicio1 +
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
                    "Validación Horario de atención",
                    rangoCapturado.Address.Replace(
                        "$",
                        ""),
                    "Data Validation + Formato Condicional + " +
                    "7 Fórmulas Auxiliares (" +
                    rangoAuxiliares +
                    "). Auditoría de coexistencia: " +
                    estadoAuditoria
                );

                // =============================================================
                // RESTAURAR RANGO ORIGINAL
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
                        "Validación de Horario aplicada correctamente.\n\n" +

                        "• Horarios procesados: " +
                        totalHorariosProcesados.ToString() +
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

                        "• Formato condicional: resalta horarios incorrectos.\n" +

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
                        "Error crítico al aplicar la Validación de Horario.\n\n" +

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
        // FÓRMULA AUXILIAR 1:
        // INICIO DEL PRIMER PERIODO
        // =====================================================================

        private static string ConstruirFormulaInicio1(
            string celda)
        {
            return
                "=IF(" +
                celda +
                "=\"\"," +
                "\"\"," +
                "IFERROR(" +
                "TRIM(" +
                "LEFT(" +
                celda +
                "," +
                "FIND(\"-\"," +
                celda +
                ")-1" +
                ")" +
                ")," +
                "\"\"" +
                ")" +
                ")";
        }

        // =====================================================================
        // FÓRMULA AUXILIAR 2:
        // FIN DEL PRIMER PERIODO
        // =====================================================================

        private static string ConstruirFormulaFin1(
            string celda)
        {
            return
                "=IF(" +
                celda +
                "=\"\"," +
                "\"\"," +

                "IFERROR(" +

                "TRIM(" +

                "MID(" +
                celda +
                "," +

                "FIND(\"-\"," +
                celda +
                ")+1," +

                "MIN(" +

                "IFERROR(" +
                "FIND(\",\"," +
                celda +
                ")," +
                "LEN(" +
                celda +
                ")+1" +
                ")," +

                "IFERROR(" +
                "FIND(\"(\"," +
                celda +
                ")," +
                "LEN(" +
                celda +
                ")+1" +
                ")" +

                ")-" +

                "FIND(\"-\"," +
                celda +
                ")-1" +

                ")" +

                ")," +

                "\"\"" +

                ")" +

                ")";
        }

        // =====================================================================
        // FÓRMULA AUXILIAR 3:
        // DÍAS DEL PRIMER PERIODO
        // =====================================================================

        private static string ConstruirFormulaDias1(
            string celda)
        {
            return
                "=IF(" +

                celda +
                "=\"\"," +

                "\"\"," +

                "IFERROR(" +

                "IF(" +

                "FIND(\"(\"," +
                celda +
                ")<" +

                "IFERROR(" +
                "FIND(\",\"," +
                celda +
                ")," +
                "LEN(" +
                celda +
                ")+1" +
                ")," +

                "TRIM(" +

                "MID(" +
                celda +
                "," +

                "FIND(\"(\"," +
                celda +
                ")+1," +

                "FIND(\")\"," +
                celda +
                "," +

                "FIND(\"(\"," +
                celda +
                ")+1" +

                ")-" +

                "FIND(\"(\"," +
                celda +
                ")-1" +

                ")" +

                ")," +

                "\"\"" +

                ")," +

                "\"\"" +

                ")" +

                ")";
        }

        // =====================================================================
        // FÓRMULA AUXILIAR 4:
        // INICIO DEL SEGUNDO PERIODO
        // =====================================================================

        private static string ConstruirFormulaInicio2(
            string celda)
        {
            return
                "=IF(" +

                "OR(" +

                celda +
                "=\"\"," +

                "ISERROR(" +
                "FIND(\",\"," +
                celda +
                ")" +
                ")" +

                ")," +

                "\"\"," +

                "IFERROR(" +

                "TRIM(" +

                "MID(" +
                celda +
                "," +

                "FIND(\",\"," +
                celda +
                ")+1," +

                "FIND(\"-\"," +
                celda +
                "," +

                "FIND(\",\"," +
                celda +
                ")+1" +

                ")-" +

                "FIND(\",\"," +
                celda +
                ")-1" +

                ")" +

                ")," +

                "\"\"" +

                ")" +

                ")";
        }

        // =====================================================================
        // FÓRMULA AUXILIAR 5:
        // FIN DEL SEGUNDO PERIODO
        // =====================================================================

        private static string ConstruirFormulaFin2(
            string celda)
        {
            return
                "=IF(" +

                "OR(" +

                celda +
                "=\"\"," +

                "ISERROR(" +
                "FIND(\",\"," +
                celda +
                ")" +
                ")" +

                ")," +

                "\"\"," +

                "IFERROR(" +

                "TRIM(" +

                "MID(" +
                celda +
                "," +

                "FIND(\"-\"," +
                celda +
                "," +

                "FIND(\",\"," +
                celda +
                ")+1" +

                ")+1," +

                "IFERROR(" +

                "FIND(\"(\"," +
                celda +
                "," +

                "FIND(\",\"," +
                celda +
                ")+1" +

                ")," +

                "LEN(" +
                celda +
                ")+1" +

                ")-" +

                "FIND(\"-\"," +
                celda +
                "," +

                "FIND(\",\"," +
                celda +
                ")+1" +

                ")-1" +

                ")" +

                ")," +

                "\"\"" +

                ")" +

                ")";
        }

        // =====================================================================
        // FÓRMULA AUXILIAR 6:
        // DÍAS DEL SEGUNDO PERIODO
        // =====================================================================

        private static string ConstruirFormulaDias2(
            string celda)
        {
            return
                "=IF(" +

                "OR(" +

                celda +
                "=\"\"," +

                "ISERROR(" +
                "FIND(\",\"," +
                celda +
                ")" +
                ")" +

                ")," +

                "\"\"," +

                "IFERROR(" +

                "TRIM(" +

                "MID(" +
                celda +
                "," +

                "FIND(\"(\"," +
                celda +
                "," +

                "FIND(\",\"," +
                celda +
                ")+1" +

                ")+1," +

                "FIND(\")\"," +
                celda +
                "," +

                "FIND(\"(\"," +
                celda +
                "," +

                "FIND(\",\"," +
                celda +
                ")+1" +

                ")+1" +

                ")-" +

                "FIND(\"(\"," +
                celda +
                "," +

                "FIND(\",\"," +
                celda +
                ")+1" +

                ")-1" +

                ")" +

                ")," +

                "\"\"" +

                ")" +

                ")";
        }

        // =====================================================================
        // FÓRMULA AUXILIAR 7:
        // VÁLIDO
        //
        // 1 = CORRECTO
        // 0 = ERROR
        //
        // Una celda vacía devuelve 1.
        // La obligatoriedad corresponde a ValidacionBlancos.
        // =====================================================================

        private static string ConstruirFormulaValido(
            string celdaOriginal,
            string inicio1,
            string fin1,
            string dias1,
            string inicio2,
            string fin2,
            string dias2)
        {
            string inicio1Valido =
                ExpresionHoraValida(
                    inicio1,
                    false);

            string fin1Valido =
                ExpresionHoraValida(
                    fin1,
                    true);

            string inicio2Valido =
                ExpresionHoraValida(
                    inicio2,
                    false);

            string fin2Valido =
                ExpresionHoraValida(
                    fin2,
                    true);

            string minutosInicio1 =
                ExpresionMinutos(
                    inicio1);

            string minutosFin1 =
                ExpresionMinutos(
                    fin1);

            string minutosInicio2 =
                ExpresionMinutos(
                    inicio2);

            string minutosFin2 =
                ExpresionMinutos(
                    fin2);

            string dias1Validos =
                ExpresionDiasValidos(
                    dias1);

            string dias2Validos =
                ExpresionDiasValidos(
                    dias2);

            // -------------------------------------------------------------
            // PRIMER PERIODO
            // -------------------------------------------------------------

            string primerPeriodoValido =
                "AND(" +

                inicio1Valido +
                "," +

                fin1Valido +
                "," +

                "IFERROR(" +

                minutosInicio1 +
                "<" +
                minutosFin1 +

                ",FALSE)" +

                ")";

            // -------------------------------------------------------------
            // SEGUNDO PERIODO
            // -------------------------------------------------------------

            string segundoPeriodoValido =
                "IF(" +

                inicio2 +
                "=\"\"," +

                "AND(" +

                fin2 +
                "=\"\"," +

                dias2 +
                "=\"\"" +

                ")," +

                "AND(" +

                inicio2Valido +
                "," +

                fin2Valido +
                "," +

                "IFERROR(" +

                minutosInicio2 +
                "<" +
                minutosFin2 +

                ",FALSE)" +

                ")" +

                ")";

            // -------------------------------------------------------------
            // COHERENCIA DE DÍAS
            //
            // Si existen dos periodos:
            //
            // - ambos pueden no indicar días;
            // - o ambos deben indicar días.
            // -------------------------------------------------------------

            string coherenciaDias =
                "IF(" +

                inicio2 +
                "=\"\"," +

                "TRUE," +

                "OR(" +

                "AND(" +

                dias1 +
                "=\"\"," +

                dias2 +
                "=\"\"" +

                ")," +

                "AND(" +

                dias1 +
                "<>\"\"," +

                dias2 +
                "<>\"\"" +

                ")" +

                ")" +

                ")";

            // -------------------------------------------------------------
            // TRASLAPE DE DÍAS
            // -------------------------------------------------------------

            string indiceInicioDias1 =
                ExpresionIndiceDiaInicio(
                    dias1);

            string indiceFinDias1 =
                ExpresionIndiceDiaFin(
                    dias1);

            string indiceInicioDias2 =
                ExpresionIndiceDiaInicio(
                    dias2);

            string indiceFinDias2 =
                ExpresionIndiceDiaFin(
                    dias2);

            string diasSeTraslapan =
                "AND(" +

                dias1 +
                "<>\"\"," +

                dias2 +
                "<>\"\"," +

                "MAX(" +

                indiceInicioDias1 +
                "," +

                indiceInicioDias2 +

                ")<=" +

                "MIN(" +

                indiceFinDias1 +
                "," +

                indiceFinDias2 +

                ")" +

                ")";

            // -------------------------------------------------------------
            // ORDEN ENTRE PERIODOS
            //
            // 8:00 - 14:00, 16:00 - 18:00
            //
            // requiere:
            //
            // Fin 1 <= Inicio 2
            //
            // Cuando los días son distintos:
            //
            // 9:00 - 18:00 (L-J), 09:00 - 15:00 (V)
            //
            // las horas pueden traslaparse.
            // -------------------------------------------------------------

            string ordenPeriodos =
                "IF(" +

                inicio2 +
                "=\"\"," +

                "TRUE," +

                "IF(" +

                "AND(" +

                dias1 +
                "=\"\"," +

                dias2 +
                "=\"\"" +

                ")," +

                "IFERROR(" +

                minutosFin1 +
                "<=" +
                minutosInicio2 +

                ",FALSE)," +

                "IF(" +

                diasSeTraslapan +
                "," +

                "IFERROR(" +

                minutosFin1 +
                "<=" +
                minutosInicio2 +

                ",FALSE)," +

                "TRUE" +

                ")" +

                ")" +

                ")";

            // -------------------------------------------------------------
            // RECONSTRUCCIÓN DEL HORARIO
            //
            // Se eliminan los espacios para comparar la estructura.
            // -------------------------------------------------------------

            string horarioReconstruido =
                inicio1 +

                "&\"-\"&" +

                fin1 +

                "&IF(" +

                dias1 +
                "<>\"\"," +

                "\"(\"&" +
                dias1 +
                "&\")\"," +

                "\"\"" +

                ")" +

                "&IF(" +

                inicio2 +
                "<>\"\"," +

                "\",\"&" +

                inicio2 +

                "&\"-\"&" +

                fin2 +

                "&IF(" +

                dias2 +
                "<>\"\"," +

                "\"(\"&" +
                dias2 +
                "&\")\"," +

                "\"\"" +

                ")," +

                "\"\"" +

                ")";

            string estructuraValida =
                "SUBSTITUTE(" +

                "UPPER(TRIM(" +
                celdaOriginal +
                "))," +

                "\" \"," +
                "\"\"" +

                ")=" +

                "SUBSTITUTE(" +

                "UPPER(" +
                horarioReconstruido +
                ")," +

                "\" \"," +
                "\"\"" +

                ")";

            // -------------------------------------------------------------
            // RESULTADO FINAL
            // -------------------------------------------------------------

            return
                "=IF(" +

                celdaOriginal +
                "=\"\"," +

                "1," +

                "IFERROR(" +

                "IF(" +

                "AND(" +

                primerPeriodoValido +
                "," +

                dias1Validos +
                "," +

                segundoPeriodoValido +
                "," +

                dias2Validos +
                "," +

                coherenciaDias +
                "," +

                ordenPeriodos +
                "," +

                estructuraValida +

                ")," +

                "1," +

                "0" +

                ")," +

                "0" +

                ")" +

                ")";
        }

        // =====================================================================
        // VALIDACIÓN DE HORA
        //
        // PERMITE:
        //
        // 9:00
        // 09:00
        // 00:00
        //
        // 24:00 solo se permite como hora final.
        // =====================================================================

        private static string ExpresionHoraValida(
            string celda,
            bool permitir24)
        {
            string posicionDosPuntos =
                "FIND(\":\"," +
                celda +
                ")";

            string textoHora =
                "LEFT(" +
                celda +
                "," +
                posicionDosPuntos +
                "-1)";

            string textoMinuto =
                "RIGHT(" +
                celda +
                ",2)";

            string valorHora =
                "VALUE(" +
                textoHora +
                ")";

            string valorMinuto =
                "VALUE(" +
                textoMinuto +
                ")";

            string digitoHora1 =
                "LEFT(" +
                textoHora +
                ",1)";

            string digitoHora2 =
                "RIGHT(" +
                textoHora +
                ",1)";

            string digitoMinuto1 =
                "LEFT(" +
                textoMinuto +
                ",1)";

            string digitoMinuto2 =
                "RIGHT(" +
                textoMinuto +
                ",1)";

            string limiteHora;

            if (permitir24)
            {
                limiteHora =
                    "AND(" +

                    valorHora +
                    "<=24," +

                    "IF(" +

                    valorHora +
                    "=24," +

                    valorMinuto +
                    "=0," +

                    "TRUE" +

                    ")" +

                    ")";
            }
            else
            {
                limiteHora =
                    valorHora +
                    "<=23";
            }

            return
                "IFERROR(" +

                "AND(" +

                celda +
                "<>\"\"," +

                // ---------------------------------------------------------
                // EXACTAMENTE UN :
                // ---------------------------------------------------------

                "LEN(" +
                celda +
                ")-" +

                "LEN(" +

                "SUBSTITUTE(" +
                celda +
                ",\":\",\"\"" +

                ")" +

                ")=1," +

                // ---------------------------------------------------------
                // H:MM O HH:MM
                // ---------------------------------------------------------

                "OR(" +

                posicionDosPuntos +
                "=2," +

                posicionDosPuntos +
                "=3" +

                ")," +

                // ---------------------------------------------------------
                // DOS DÍGITOS DESPUÉS DE :
                // ---------------------------------------------------------

                "LEN(" +
                celda +
                ")-" +

                posicionDosPuntos +
                "=2," +

                // ---------------------------------------------------------
                // PRIMER DÍGITO DE HORA
                // ---------------------------------------------------------

                "ISNUMBER(" +

                "FIND(" +

                digitoHora1 +
                "," +

                "\"0123456789\"" +

                ")" +

                ")," +

                // ---------------------------------------------------------
                // SEGUNDO DÍGITO DE HORA SI EXISTE
                // ---------------------------------------------------------

                "IF(" +

                "LEN(" +
                textoHora +
                ")=2," +

                "ISNUMBER(" +

                "FIND(" +

                digitoHora2 +
                "," +

                "\"0123456789\"" +

                ")" +

                ")," +

                "TRUE" +

                ")," +

                // ---------------------------------------------------------
                // MINUTOS NUMÉRICOS
                // ---------------------------------------------------------

                "ISNUMBER(" +

                "FIND(" +

                digitoMinuto1 +
                "," +

                "\"0123456789\"" +

                ")" +

                ")," +

                "ISNUMBER(" +

                "FIND(" +

                digitoMinuto2 +
                "," +

                "\"0123456789\"" +

                ")" +

                ")," +

                // ---------------------------------------------------------
                // LÍMITES DE HORA / MINUTO
                // ---------------------------------------------------------

                valorHora +
                ">=0," +

                limiteHora +
                "," +

                valorMinuto +
                ">=0," +

                valorMinuto +
                "<=59" +

                ")," +

                "FALSE" +

                ")";
        }

        // =====================================================================
        // CONVERTIR HH:MM A MINUTOS
        // =====================================================================

        private static string ExpresionMinutos(
            string celda)
        {
            string posicionDosPuntos =
                "FIND(\":\"," +
                celda +
                ")";

            string valorHora =
                "VALUE(" +

                "LEFT(" +

                celda +
                "," +

                posicionDosPuntos +
                "-1" +

                ")" +

                ")";

            string valorMinuto =
                "VALUE(" +

                "RIGHT(" +

                celda +
                "," +

                "2" +

                ")" +

                ")";

            return
                "IFERROR(" +

                "(" +
                valorHora +
                "*60)+" +

                valorMinuto +

                "," +

                "-1" +

                ")";
        }

        // =====================================================================
        // VALIDACIÓN DE DÍAS
        //
        // PERMITE:
        //
        // L
        // M
        // X
        // J
        // V
        // S
        // D
        //
        // Y RANGOS:
        //
        // L-J
        // L-V
        // S-D
        // =====================================================================

        private static string ExpresionDiasValidos(
            string celda)
        {
            string normalizado =
                "UPPER(" +

                "SUBSTITUTE(" +

                celda +
                "," +

                "\" \"," +
                "\"\"" +

                ")" +

                ")";

            string indiceInicio =
                "FIND(" +

                "LEFT(" +

                normalizado +
                "," +

                "1" +

                ")," +

                "\"LMXJVSD\"" +

                ")";

            string indiceFin =
                "FIND(" +

                "RIGHT(" +

                normalizado +
                "," +

                "1" +

                ")," +

                "\"LMXJVSD\"" +

                ")";

            return
                "IF(" +

                celda +
                "=\"\"," +

                "TRUE," +

                "IFERROR(" +

                "OR(" +

                // ---------------------------------------------------------
                // UN SOLO DÍA
                // ---------------------------------------------------------

                "AND(" +

                "LEN(" +
                normalizado +
                ")=1," +

                "ISNUMBER(" +
                indiceInicio +
                ")" +

                ")," +

                // ---------------------------------------------------------
                // RANGO DE DÍAS
                // ---------------------------------------------------------

                "AND(" +

                "LEN(" +
                normalizado +
                ")=3," +

                "MID(" +
                normalizado +
                ",2,1)=\"-\"," +

                "ISNUMBER(" +
                indiceInicio +
                ")," +

                "ISNUMBER(" +
                indiceFin +
                ")," +

                indiceInicio +
                "<" +
                indiceFin +

                ")" +

                ")," +

                "FALSE" +

                ")" +

                ")";
        }

        // =====================================================================
        // ÍNDICE DEL DÍA INICIAL
        // =====================================================================

        private static string ExpresionIndiceDiaInicio(
            string celda)
        {
            string normalizado =
                "UPPER(" +

                "SUBSTITUTE(" +

                celda +
                "," +

                "\" \"," +
                "\"\"" +

                ")" +

                ")";

            return
                "IF(" +

                celda +
                "=\"\"," +

                "0," +

                "IFERROR(" +

                "FIND(" +

                "LEFT(" +
                normalizado +
                ",1)," +

                "\"LMXJVSD\"" +

                ")," +

                "0" +

                ")" +

                ")";
        }

        // =====================================================================
        // ÍNDICE DEL DÍA FINAL
        // =====================================================================

        private static string ExpresionIndiceDiaFin(
            string celda)
        {
            string normalizado =
                "UPPER(" +

                "SUBSTITUTE(" +

                celda +
                "," +

                "\" \"," +
                "\"\"" +

                ")" +

                ")";

            return
                "IF(" +

                celda +
                "=\"\"," +

                "0," +

                "IFERROR(" +

                "FIND(" +

                "RIGHT(" +
                normalizado +
                ",1)," +

                "\"LMXJVSD\"" +

                ")," +

                "0" +

                ")" +

                ")";
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
        // OBTENER FILA INICIAL DEL RANGO
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
        // COMPROBAR TRASLAPE ENTRE COLUMNAS
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