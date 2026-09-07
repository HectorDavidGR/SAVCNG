using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Consumo obligatorio de la Fachada y el DTO[cite: 3, 4]

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionFechas : IValidacionExcel
    {
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                // =========================================================================
                // FASE 1: VALIDACIÓN ESTRUCTURAL DINÁMICA (SOPORTA CELDAS COMBINADAS)
                // =========================================================================
                Excel.Range primerArea = null;
                Excel.Range pTopLeftDia = null, pTopLeftMes = null, pTopLeftAno = null;
                Excel.Range pMDia = null, pMMes = null, pMAno = null;

                try
                {
                    primerArea = (Excel.Range)rangoCapturado.Areas[1];

                    pTopLeftDia = (Excel.Range)primerArea.Cells[1, 1];
                    pMDia = pTopLeftDia.MergeArea;
                    int offsetMes = pMDia.Columns.Count + 1;

                    if (offsetMes > primerArea.Columns.Count) throw new Exception();

                    pTopLeftMes = (Excel.Range)primerArea.Cells[1, offsetMes];
                    pMMes = pTopLeftMes.MergeArea;
                    int offsetAno = offsetMes + pMMes.Columns.Count;

                    if (offsetAno > primerArea.Columns.Count) throw new Exception();

                    pTopLeftAno = (Excel.Range)primerArea.Cells[1, offsetAno];
                    pMAno = pTopLeftAno.MergeArea;

                    int widthTotal = pMDia.Columns.Count + pMMes.Columns.Count + pMAno.Columns.Count;

                    if (primerArea.Columns.Count != widthTotal) throw new Exception();
                }
                catch
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "El rango seleccionado debe contener exactamente 3 campos continuos (Día, Mes, Año), incluso si están formadas por celdas combinadas.", AlertaInyectada = false };
                }
                finally
                {
                    ExcelHelper.LiberarCom(pMAno); ExcelHelper.LiberarCom(pMMes); ExcelHelper.LiberarCom(pMDia);
                    ExcelHelper.LiberarCom(pTopLeftAno); ExcelHelper.LiberarCom(pTopLeftMes); ExcelHelper.LiberarCom(pTopLeftDia);
                }

                // =========================================================================
                // FASE 2: UX DE CONFIGURACIÓN DE LÍMITES Y ÁREA DE BANDERAS
                // =========================================================================
                object resInferior = excelApp.InputBox("Indica el AÑO MÍNIMO aceptado:\n\n(Ej. 1900 o 1990).", "SAVCNG - Límite Inferior", "1900", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);
                if (resInferior is bool && (bool)resInferior == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };

                object resSuperior = excelApp.InputBox("Indica el AÑO MÁXIMO aceptado:\n\n(Ej. 2026).", "SAVCNG - Límite Superior", "2026", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);
                if (resSuperior is bool && (bool)resSuperior == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };

                if (!int.TryParse(resInferior.ToString(), out int minYear) || !int.TryParse(resSuperior.ToString(), out int maxYear) || minYear > maxYear)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Límites de año inválidos.", AlertaInyectada = false };
                }

                Excel.Range rangoAlertaDefinido = null;
                try
                {
                    rangoAlertaDefinido = (Excel.Range)excelApp.InputBox(
                        "Selecciona la celda INICIAL (columna AF en adelante) donde se insertara la formula auxiliar:\n\n(1 = Error, 0 = Correcto).",
                        "SAVCNG - Formula Auxiliar", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                }
                catch
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Selección de formula auxiliar cancelada.", AlertaInyectada = false };
                }

                // =========================================================================
                // FASE 3: AUDITORÍA DE COEXISTENCIA AVANZADA (ZERO LEAKS)
                // =========================================================================
                bool tieneValidacionPrevia = false;
                bool tieneFormatosPrevios = false;
                bool tieneAlertaPrevia = false;
                bool limpiarFormatos = true;

                try { var tipo = rangoCapturado.Validation.Type; tieneValidacionPrevia = true; } catch { }

                Excel.FormatConditions conds = null;
                try
                {
                    conds = rangoCapturado.FormatConditions;
                    if (conds.Count > 0) tieneFormatosPrevios = true;
                }
                catch { }
                finally { ExcelHelper.LiberarCom(conds); }

                Excel.Range celdaAlertaBaseCheck = null;
                try
                {
                    celdaAlertaBaseCheck = (Excel.Range)rangoAlertaDefinido.Cells[1, 1];
                    string fCheck = celdaAlertaBaseCheck.Formula?.ToString() ?? "";
                    if (fCheck.StartsWith("=") || (celdaAlertaBaseCheck.Value2 != null && !string.IsNullOrEmpty(celdaAlertaBaseCheck.Value2.ToString())))
                    {
                        tieneAlertaPrevia = true;
                    }
                }
                catch { }
                finally { ExcelHelper.LiberarCom(celdaAlertaBaseCheck); }

                if (tieneValidacionPrevia || tieneFormatosPrevios || tieneAlertaPrevia)
                {
                    DialogResult resp = MessageBox.Show(
                         "Se detectaron configuraciones previas en el rango seleccionado.\n\n" +
                        "NOTA: Al ser una restricción de captura estricta, la regla se sobreescribirá, pero...\n\n" +
                        "¿Deseas CONSERVAR las formulas, alertas o bloqueos (Formatos Condicionales) aplicados previamente?\n\n" +
                        "SÍ = CONSERVAR todo.\n" +
                        "NO = BORRAR todo el historial y limpiar las celdas.",
                        "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                    if (resp == DialogResult.Cancel)
                        return new ResultadoValidacion { Exito = false, Mensaje = "Proceso cancelado por el usuario.", AlertaInyectada = false };

                    if (resp == DialogResult.Yes)
                        limpiarFormatos = false;
                }

                System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();
                int totalAreas = rangoCapturado.Areas.Count;

                for (int i = 1; i <= totalAreas; i++)
                {
                    Excel.Range areaIndividual = null;
                    try
                    {
                        areaIndividual = (Excel.Range)rangoCapturado.Areas[i];
                        string idPregunta = ExcelHelper.ObtenerNumeroPregunta(areaIndividual);
                        if (!string.IsNullOrEmpty(idPregunta)) preguntasUnicas.Add(idPregunta);
                    }
                    finally { ExcelHelper.LiberarCom(areaIndividual); }
                }
                string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 4: INYECCIÓN MATRICIAL (EVALUACIÓN POR FILA EN COLUMNA AUXILIAR)
                // =========================================================================
                try
                {
                    int colorFondoRojo = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 199, 206));
                    int colorTextoRojo = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));

                    for (int j = 1; j <= totalAreas; j++)
                    {
                        Excel.Range area = null;
                        Excel.Range celdaDia = null, celdaMes = null, celdaAno = null;
                        Excel.Range mDia = null, mMes = null, mAno = null;
                        Excel.Range rangoDia = null, rangoMes = null, rangoAno = null;
                        Excel.Range colEvalDia = null, colEvalMes = null, colEvalAno = null;
                        Excel.FormatCondition fcDia = null, fcMes = null, fcAno = null;

                        Excel.Range celdaBaseAux = null;
                        Excel.Range targetAux = null;
                        Excel.Range primeraCeldaAux = null;

                        try
                        {
                            area = (Excel.Range)rangoCapturado.Areas[j];
                            area.Validation.Delete();
                            if (limpiarFormatos) { area.FormatConditions.Delete(); }

                            celdaDia = (Excel.Range)area.Cells[1, 1];
                            mDia = celdaDia.MergeArea; int wDia = mDia.Columns.Count;

                            celdaMes = (Excel.Range)area.Cells[1, wDia + 1];
                            mMes = celdaMes.MergeArea; int wMes = mMes.Columns.Count;

                            celdaAno = (Excel.Range)area.Cells[1, wDia + wMes + 1];
                            mAno = celdaAno.MergeArea; int wAno = mAno.Columns.Count;

                            rangoDia = celdaDia.get_Resize(area.Rows.Count, wDia);
                            rangoMes = celdaMes.get_Resize(area.Rows.Count, wMes);
                            rangoAno = celdaAno.get_Resize(area.Rows.Count, wAno);

                            colEvalDia = celdaDia.get_Resize(area.Rows.Count, 1);
                            colEvalMes = celdaMes.get_Resize(area.Rows.Count, 1);
                            colEvalAno = celdaAno.get_Resize(area.Rows.Count, 1);

                            rangoDia.NumberFormat = "00";
                            rangoMes.NumberFormat = "00";
                            rangoAno.NumberFormat = "0000";

                            string relDia = celdaDia.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                            string relMes = celdaMes.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                            string relAno = celdaAno.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                            string formDiaEng = $"=OR(TRIM({relDia})=\"NS\", TRIM({relDia})=\"NA\", AND(ISNUMBER({relDia}), {relDia}>0, {relDia}<=IF(AND(ISNUMBER({relMes}), ISNUMBER({relAno})), DAY(DATE({relAno}, {relMes}+1, 0)), 31), NOT(OR(TRIM({relMes})=\"NS\", TRIM({relMes})=\"NA\"))))";
                            string formMesEng = $"=OR(TRIM({relMes})=\"NS\", TRIM({relMes})=\"NA\", AND(ISNUMBER({relMes}), {relMes}>=1, {relMes}<=12, NOT(OR(TRIM({relAno})=\"NS\", TRIM({relAno})=\"NA\"))))";
                            string formAnoEng = $"=OR(TRIM({relAno})=\"NS\", TRIM({relAno})=\"NA\", AND(ISNUMBER({relAno}), {relAno}>={minYear}, {relAno}<={maxYear}))";

                            rangoDia.Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, ExcelHelper.TraducirFormulaLocal(wsActual, formDiaEng, area.Row), Type.Missing);
                            rangoDia.Validation.IgnoreBlank = true; rangoDia.Validation.ShowError = true;

                            rangoMes.Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, ExcelHelper.TraducirFormulaLocal(wsActual, formMesEng, area.Row), Type.Missing);
                            rangoMes.Validation.IgnoreBlank = true; rangoMes.Validation.ShowError = true;

                            rangoAno.Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, ExcelHelper.TraducirFormulaLocal(wsActual, formAnoEng, area.Row), Type.Missing);
                            rangoAno.Validation.IgnoreBlank = true; rangoAno.Validation.ShowError = true;

                            string safeRelAno = $"IF(ISNUMBER({relAno}), {relAno}, 2000)";
                            string safeRelMes = $"IF(ISNUMBER({relMes}), {relMes}, 1)";

                            string fcDiaEng = $"=AND(ISNUMBER({relDia}), OR({relDia}>DAY(DATE({safeRelAno}, {safeRelMes}+1, 0)), TRIM({relMes})=\"NS\", TRIM({relMes})=\"NA\"))";
                            string fcMesEng = $"=AND(ISNUMBER({relMes}), OR({relMes}<1, {relMes}>12, TRIM({relAno})=\"NS\", TRIM({relAno})=\"NA\"))";
                            string fcAnoEng = $"=AND(ISNUMBER({relAno}), OR({relAno}<{minYear}, {relAno}>{maxYear}))";

                            fcDia = (Excel.FormatCondition)rangoDia.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, ExcelHelper.TraducirFormulaLocal(wsActual, fcDiaEng, area.Row));
                            fcDia.Interior.Color = colorFondoRojo; fcDia.Font.Color = colorTextoRojo;

                            fcMes = (Excel.FormatCondition)rangoMes.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, ExcelHelper.TraducirFormulaLocal(wsActual, fcMesEng, area.Row));
                            fcMes.Interior.Color = colorFondoRojo; fcMes.Font.Color = colorTextoRojo;

                            fcAno = (Excel.FormatCondition)rangoAno.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, ExcelHelper.TraducirFormulaLocal(wsActual, fcAnoEng, area.Row));
                            fcAno.Interior.Color = colorFondoRojo; fcAno.Font.Color = colorTextoRojo;

                            int offsetFilas = area.Row - primerArea.Row;
                            celdaBaseAux = (Excel.Range)rangoAlertaDefinido.Cells[1, 1];
                            targetAux = celdaBaseAux.get_Offset(offsetFilas, 0).get_Resize(area.Rows.Count, 1);

                            targetAux.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            targetAux.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            targetAux.Font.Bold = true;

                            string formulaVieja = "";
                            primeraCeldaAux = (Excel.Range)targetAux.Cells[1, 1];

                            if (primeraCeldaAux.HasFormula != null && (bool)primeraCeldaAux.HasFormula)
                            {
                                string fv = primeraCeldaAux.Formula.ToString();
                                if (fv.StartsWith("=")) formulaVieja = fv.Substring(1);
                            }
                            else
                            {
                                string txt = primeraCeldaAux.Value2?.ToString() ?? "";
                                if (!string.IsNullOrEmpty(txt)) formulaVieja = txt;
                            }

                            string condErrorEng = $"OR(" +
                                $"AND(ISNUMBER({relDia}), OR({relDia}>DAY(DATE({safeRelAno}, {safeRelMes}+1, 0)), TRIM({relMes})=\"NS\", TRIM({relMes})=\"NA\")), " +
                                $"AND(ISNUMBER({relMes}), OR({relMes}<1, {relMes}>12, TRIM({relAno})=\"NS\", TRIM({relAno})=\"NA\")), " +
                                $"AND(ISNUMBER({relAno}), OR({relAno}<{minYear}, {relAno}>{maxYear}))" +
                            $")";

                            string logicaNueva = $"IF({condErrorEng}, 1, 0)";
                            string formulaAlertaEng = "";

                            if (!string.IsNullOrEmpty(formulaVieja) && !limpiarFormatos)
                            {
                                formulaAlertaEng = $"=IF(SUM({formulaVieja}, {logicaNueva})>0, 1, 0)";
                            }
                            else
                            {
                                formulaAlertaEng = $"={logicaNueva}";
                            }

                            targetAux.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaAlertaEng, area.Row);
                        }
                        finally
                        {
                            ExcelHelper.LiberarCom(primeraCeldaAux); ExcelHelper.LiberarCom(targetAux); ExcelHelper.LiberarCom(celdaBaseAux);
                            ExcelHelper.LiberarCom(colEvalAno); ExcelHelper.LiberarCom(colEvalMes); ExcelHelper.LiberarCom(colEvalDia);
                            ExcelHelper.LiberarCom(fcAno); ExcelHelper.LiberarCom(fcMes); ExcelHelper.LiberarCom(fcDia);
                            ExcelHelper.LiberarCom(rangoAno); ExcelHelper.LiberarCom(rangoMes); ExcelHelper.LiberarCom(rangoDia);
                            ExcelHelper.LiberarCom(mAno); ExcelHelper.LiberarCom(mMes); ExcelHelper.LiberarCom(mDia);
                            ExcelHelper.LiberarCom(celdaAno); ExcelHelper.LiberarCom(celdaMes); ExcelHelper.LiberarCom(celdaDia);
                            ExcelHelper.LiberarCom(area);
                        }
                    }
                }
                finally
                {
                    ExcelHelper.LiberarCom(rangoAlertaDefinido);
                    ExcelHelper.LiberarCom(primerArea);
                }

                // =========================================================================
                // FASE 4.5: SUB-FLUJO DESACOPLADO (ALERTA GLOBAL UNIFICADA)
                // =========================================================================
                excelApp.ScreenUpdating = true; // Mostramos visualmente el avance
                DialogResult respMensaje = MessageBox.Show(
                    "¿Deseas agregar un MENSAJE DE TEXTO ligado a la(s) formula(s) auxiliar(es) ingresada(s)?\n\n(Debera seleccionar el rango de la(s) formula(s) axuliar(es)).",
                    "SAVCNG - Mensaje Descriptivo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (respMensaje == DialogResult.Yes)
                {
                    Excel.Range rangoBanderasSelect = null;
                    Excel.Range rangoMensajesSelect = null;
                    try
                    {
                        rangoBanderasSelect = (Excel.Range)excelApp.InputBox(
                            "Selecciona el RANGO COMPLETO de la(s) formula(s) axuliar(es) ingresada(s) previamente:",
                            "SAVCNG - Origen", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

                        rangoMensajesSelect = (Excel.Range)excelApp.InputBox(
                            "Selecciona el RANGO donde se insertará el MENSAJE DE ERROR:",
                            "SAVCNG - Destino", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

                        excelApp.ScreenUpdating = false;

                        Excel.Range celdaExtBase = null;
                        Excel.Range mergeExt = null;
                        Excel.Range topLeftExt = null;
                        string formViejaM = "";

                        try
                        {
                            // 1. Extracción Segura usando el Ancla de celda combinada (MergeArea)
                            celdaExtBase = (Excel.Range)rangoMensajesSelect.Cells[1, 1];
                            mergeExt = celdaExtBase.MergeArea;
                            topLeftExt = (Excel.Range)mergeExt.Cells[1, 1];

                            if (topLeftExt.HasFormula != null && (bool)topLeftExt.HasFormula)
                            {
                                string fv = topLeftExt.Formula.ToString();
                                if (fv.StartsWith("=")) formViejaM = fv.Substring(1);
                            }
                            else
                            {
                                string tOld = topLeftExt.Value2?.ToString() ?? "";
                                if (!string.IsNullOrEmpty(tOld)) formViejaM = "\"" + tOld.Replace("\"", "\"\"") + "\"";
                            }

                            // 2. Preparamos el Lienzo Gigante
                            rangoMensajesSelect.Merge();
                            rangoMensajesSelect.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            rangoMensajesSelect.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            rangoMensajesSelect.WrapText = true;
                            rangoMensajesSelect.Font.Color = 255;
                            rangoMensajesSelect.Font.Bold = true;

                            // 3. Obtenemos dirección absoluta del rango de banderas y creamos la lógica
                            string dirBanderas = rangoBanderasSelect.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                            string logicaMsjEng = $"IF(SUM({dirBanderas})>0, \"Inconsistencia en fecha(s)\", \"\")";
                            string formMsjEng = "";

                            // 4. Apilamiento Seguro
                            if (!string.IsNullOrEmpty(formViejaM) && !limpiarFormatos)
                            {
                                formMsjEng = $"={formViejaM} & IF({logicaMsjEng}=\"\", \"\", CHAR(10) & {logicaMsjEng})";
                            }
                            else
                            {
                                formMsjEng = $"={logicaMsjEng}";
                            }

                            rangoMensajesSelect.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formMsjEng, rangoMensajesSelect.Row);
                        }
                        finally
                        {
                            ExcelHelper.LiberarCom(topLeftExt); ExcelHelper.LiberarCom(mergeExt); ExcelHelper.LiberarCom(celdaExtBase);
                        }
                    }
                    catch
                    {
                        // Flujo cancelado por usuario de manera segura
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(rangoMensajesSelect); ExcelHelper.LiberarCom(rangoBanderasSelect);
                    }
                }

                // =========================================================================
                // FASE 5: REGISTRO DE AUDITORÍA Y RETORNO (DTO)
                // =========================================================================
                string estadoAuditoria = limpiarFormatos ? " " : "• Se conservaron formulas, formatos y bloqueos.";

                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    "Validación Fechas (Formula Auxiliar)",
                    rangoCapturado.Address.Replace("$", ""),
                    $"Data Validation + Columna Auxiliar ({estadoAuditoria})" //[cite: 2]
                );

                return new ResultadoValidacion //[cite: 4]
                {
                    Exito = true,
                    Mensaje = $"Validación de fechas aplicada con éxito.\n\n• Evaluación binaria (1 o 0) fila por fila insertada.\n{estadoAuditoria}",
                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = $"Error crítico: {ex.Message}", AlertaInyectada = false };
            }
            finally
            {
                ExcelHelper.LiberarCom(wsActual); //[cite: 3]
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}