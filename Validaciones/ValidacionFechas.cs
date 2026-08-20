using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Consumo obligatorio de la Fachada y el DTO

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
                    return new ResultadoValidacion { Exito = false, Mensaje = "El rango seleccionado debe contener exactamente 3 variables contiguas (Día, Mes, Año), incluso si están formadas por celdas combinadas.", AlertaInyectada = false };
                }
                finally
                {
                    ExcelHelper.LiberarCom(pMAno); ExcelHelper.LiberarCom(pMMes); ExcelHelper.LiberarCom(pMDia);
                    ExcelHelper.LiberarCom(pTopLeftAno); ExcelHelper.LiberarCom(pTopLeftMes); ExcelHelper.LiberarCom(pTopLeftDia);
                }

                // =========================================================================
                // FASE 2: UX DE CONFIGURACIÓN DE LÍMITES Y ÁREA DE ALERTA
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
                        "Selecciona el RANGO donde se inyectará el mensaje de error global unificado:",
                        "SAVCNG - Ubicación de Alerta Visual", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                }
                catch
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Selección de alerta cancelada.", AlertaInyectada = false };
                }

                // =========================================================================
                // FASE 3: AUDITORÍA DE COEXISTENCIA AVANZADA (ZERO LEAKS)
                // =========================================================================
                bool tieneValidacionPrevia = false;
                bool tieneFormatosPrevios = false;
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

                if (tieneValidacionPrevia || tieneFormatosPrevios)
                {
                    DialogResult resp = MessageBox.Show(
                        "Se detectaron reglas de captura o colores previos en el rango seleccionado.\n\n" +
                        "NOTA: Las reglas de validación estricta (restricción de escritura) se sobreescribirán obligatoriamente.\n\n" +
                        "¿Deseas MANTENER los colores/alertas (Formatos Condicionales) aplicados previamente para que se apilen?\n\n" +
                        "SÍ = Apilar los nuevos colores sobre los existentes.\n" +
                        "NO = Borrar todo el historial visual y limpiar el lienzo antes de aplicar.",
                        "SAVCNG - Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

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
                // FASE 4: INYECCIÓN MATRICIAL (NUEVA REGLA NS ATÍPICO)
                // =========================================================================
                try
                {
                    rangoAlertaDefinido.Merge();
                    rangoAlertaDefinido.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rangoAlertaDefinido.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rangoAlertaDefinido.Font.Color = 255;
                    rangoAlertaDefinido.Font.Bold = true;

                    string condAno = ""; string condMes = ""; string condDia = ""; string condAtipicoNS = "";

                    int colorFondoRojo = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 199, 206));
                    int colorTextoRojo = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));

                    for (int j = 1; j <= totalAreas; j++)
                    {
                        Excel.Range area = null;
                        Excel.Range celdaDia = null, celdaMes = null, celdaAno = null;
                        Excel.Range mDia = null, mMes = null, mAno = null;
                        Excel.Range rangoDia = null, rangoMes = null, rangoAno = null;

                        Excel.FormatCondition fcDia = null, fcMes = null, fcAno = null;

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

                            rangoDia.NumberFormat = "00";
                            rangoMes.NumberFormat = "00";
                            rangoAno.NumberFormat = "0000";

                            string dirDia = rangoDia.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                            string dirMes = rangoMes.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                            string dirAno = rangoAno.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                            string relDia = celdaDia.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                            string relMes = celdaMes.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                            string relAno = celdaAno.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                            // =================================================================
                            // INYECCIÓN 1: DATA VALIDATION (Incorpora bloqueo por NS atípico)
                            // =================================================================
                            string formDiaEng = $"=OR(TRIM({relDia})=\"NS\", TRIM({relDia})=\"NA\", AND(ISNUMBER({relDia}), {relDia}>0, {relDia}<=IF(AND(ISNUMBER({relMes}), ISNUMBER({relAno})), DAY(DATE({relAno}, {relMes}+1, 0)), 31), NOT(OR(TRIM({relMes})=\"NS\", TRIM({relMes})=\"NA\"))))";
                            string formMesEng = $"=OR(TRIM({relMes})=\"NS\", TRIM({relMes})=\"NA\", AND(ISNUMBER({relMes}), {relMes}>=1, {relMes}<=12, NOT(OR(TRIM({relAno})=\"NS\", TRIM({relAno})=\"NA\"))))";
                            string formAnoEng = $"=OR(TRIM({relAno})=\"NS\", TRIM({relAno})=\"NA\", AND(ISNUMBER({relAno}), {relAno}>={minYear}, {relAno}<={maxYear}))";

                            rangoDia.Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, ExcelHelper.TraducirFormulaLocal(wsActual, formDiaEng, area.Row), Type.Missing);
                            rangoDia.Validation.IgnoreBlank = true; rangoDia.Validation.ShowError = true;

                            rangoMes.Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, ExcelHelper.TraducirFormulaLocal(wsActual, formMesEng, area.Row), Type.Missing);
                            rangoMes.Validation.IgnoreBlank = true; rangoMes.Validation.ShowError = true;

                            rangoAno.Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, ExcelHelper.TraducirFormulaLocal(wsActual, formAnoEng, area.Row), Type.Missing);
                            rangoAno.Validation.IgnoreBlank = true; rangoAno.Validation.ShowError = true;

                            // =================================================================
                            // INYECCIÓN 2: FORMAT CONDITIONS (Detecta exactamente quién falló el NS)
                            // =================================================================
                            string fcDiaEng = $"=AND(ISNUMBER({relDia}), OR({relDia}>IF(AND(ISNUMBER({relMes}), ISNUMBER({relAno})), DAY(DATE({relAno}, {relMes}+1, 0)), 31), TRIM({relMes})=\"NS\", TRIM({relMes})=\"NA\"))";
                            string fcMesEng = $"=AND(ISNUMBER({relMes}), OR({relMes}<1, {relMes}>12, TRIM({relAno})=\"NS\", TRIM({relAno})=\"NA\"))";
                            string fcAnoEng = $"=AND(ISNUMBER({relAno}), OR({relAno}<{minYear}, {relAno}>{maxYear}))";

                            fcDia = (Excel.FormatCondition)rangoDia.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, ExcelHelper.TraducirFormulaLocal(wsActual, fcDiaEng, area.Row));
                            fcDia.Interior.Color = colorFondoRojo; fcDia.Font.Color = colorTextoRojo;

                            fcMes = (Excel.FormatCondition)rangoMes.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, ExcelHelper.TraducirFormulaLocal(wsActual, fcMesEng, area.Row));
                            fcMes.Interior.Color = colorFondoRojo; fcMes.Font.Color = colorTextoRojo;

                            fcAno = (Excel.FormatCondition)rangoAno.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, ExcelHelper.TraducirFormulaLocal(wsActual, fcAnoEng, area.Row));
                            fcAno.Interior.Color = colorFondoRojo; fcAno.Font.Color = colorTextoRojo;

                            // =================================================================
                            // ACUMULADORES GLOBALES (+ condAtipicoNS)
                            // =================================================================
                            condAno += $"SUMPRODUCT(--ISNUMBER({dirAno}), --(({dirAno}>{maxYear})+({dirAno}<{minYear})))+";
                            condMes += $"SUMPRODUCT(--ISNUMBER({dirMes}), --(({dirMes}>12)+({dirMes}<1)))+";
                            condDia += $"SUMPRODUCT(--ISNUMBER({dirDia}), --ISNUMBER({dirMes}), --ISNUMBER({dirAno}), --({dirDia}>DAY(DATE({dirAno}, {dirMes}+1, 0))))+";

                            // Matemática de vectores: Usa '+' en lugar de OR() para compatibilidad absoluta en sumatorias booleanas
                            condAtipicoNS += $"SUMPRODUCT(--ISNUMBER({dirDia}), --((TRIM({dirMes})=\"NS\")+(TRIM({dirMes})=\"NA\")>0)) + SUMPRODUCT(--ISNUMBER({dirMes}), --((TRIM({dirAno})=\"NS\")+(TRIM({dirAno})=\"NA\")>0))+";
                        }
                        finally
                        {
                            ExcelHelper.LiberarCom(fcAno); ExcelHelper.LiberarCom(fcMes); ExcelHelper.LiberarCom(fcDia);
                            ExcelHelper.LiberarCom(rangoAno); ExcelHelper.LiberarCom(rangoMes); ExcelHelper.LiberarCom(rangoDia);
                            ExcelHelper.LiberarCom(mAno); ExcelHelper.LiberarCom(mMes); ExcelHelper.LiberarCom(mDia);
                            ExcelHelper.LiberarCom(celdaAno); ExcelHelper.LiberarCom(celdaMes); ExcelHelper.LiberarCom(celdaDia);
                            ExcelHelper.LiberarCom(area);
                        }
                    }

                    // ENSAMBLAJE FINAL DE ALERTAS
                    condAno += "0"; condMes += "0"; condDia += "0"; condAtipicoNS += "0";
                    string formulaAlertaEng = $"=IF(({condAno})>0, \"Error: Año fuera de límite establecido.\", " +
                                              $"IF(({condMes})>0, \"Error: Mes inválido detectado.\", " +
                                              $"IF(({condAtipicoNS})>0, \"Error: Inconsistencia de fecha.\", " +
                                              $"IF(({condDia})>0, \"Error: Día excede el límite del mes/año.\", \"\"))))";

                    rangoAlertaDefinido.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaAlertaEng, rangoAlertaDefinido.Row);
                }
                finally
                {
                    ExcelHelper.LiberarCom(rangoAlertaDefinido);
                    ExcelHelper.LiberarCom(primerArea);
                }

                // =========================================================================
                // FASE 5: REGISTRO DE AUDITORÍA Y RETORNO (DTO)
                // =========================================================================
                string estadoAuditoria = limpiarFormatos ? "Lienzo limpio" : "Formatos apilados";

                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    "Validación Fechas",
                    rangoCapturado.Address.Replace("$", ""),
                    $"Data Validation + SUMPRODUCT ({estadoAuditoria})"
                );

                return new ResultadoValidacion
                {
                    Exito = true,
                    Mensaje = $"Blindaje de fechas aplicado con éxito.\n\n• Alerta unificada activa.\n• Se inyectó la regla de contención de NS atípicos.\n• {estadoAuditoria}.",
                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = $"Error crítico: {ex.Message}", AlertaInyectada = false };
            }
            finally
            {
                ExcelHelper.LiberarCom(wsActual);
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}