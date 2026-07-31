using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core;

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionFormatoTexto : IValidacionExcel
    {
        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                // =========================================================================
                // FASE 1: AUDITORÍA DE COEXISTENCIA (DETECCIÓN DE REGLAS PREVIAS)
                // =========================================================================
                bool tieneFormatosPrevios = rangoCapturado.FormatConditions.Count > 0;
                bool tieneValidacionPrevia = false;

                try { var tipo = rangoCapturado.Validation.Type; tieneValidacionPrevia = true; } catch { /* Silenciado: No hay DataValidation previa */ }

                bool limpiarFormatos = false;

                if (tieneFormatosPrevios || tieneValidacionPrevia)
                {
                    DialogResult respLimpieza = MessageBox.Show(
                        "Se detectaron configuraciones previas en el rango seleccionado.\n\n" +
                        "NOTA: Al ser una restricción de captura estricta, la regla de celdas se sobreescribirá, pero...\n\n" +
                        "¿Deseas CONSERVAR los colores, alertas o bloqueos (Formatos Condicionales) aplicados previamente?\n\n" +
                        "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                        "NO = BORRAR todo el historial y limpiar el lienzo.",
                        "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                    if (respLimpieza == DialogResult.Cancel)
                    {
                        MessageBox.Show("Proceso cancelado.\nNo se alteró la plantilla.", "Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (respLimpieza == DialogResult.No)
                    {
                        limpiarFormatos = true;
                    }
                }

                // =========================================================================
                // FASE 1.5: RECOLECCIÓN LIGERA DE PREGUNTAS (STATE MANAGEMENT)
                // =========================================================================
                System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();

                for (int i = 1; i <= rangoCapturado.Areas.Count; i++)
                {
                    Excel.Range areaIndividual = (Excel.Range)rangoCapturado.Areas[i];
                    string idPregunta = ExcelHelper.ObtenerNumeroPregunta(areaIndividual);

                    if (!string.IsNullOrEmpty(idPregunta))
                    {
                        preguntasUnicas.Add(idPregunta);
                    }

                    ExcelHelper.LiberarCom(areaIndividual);
                }

                // =========================================================================
                // FASE 2: UX - TÉCNICA DEL RANGO AUXILIAR DINÁMICO
                // =========================================================================
                object resAuxiliar = excelApp.InputBox(
                    "Indica en qué columna libre deseas colocar la validación oculta (AF en adelante).\n\nConsidera que el sistema requerirá espacio libre hacia abajo proporcional al número de filas que seleccionaste.",
                    "SAVCNG - Selección de Fórmula Auxiliar", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                if (resAuxiliar is bool && (bool)resAuxiliar == false) { return; }

                Excel.Range seleccionAuxiliar = null;
                Excel.Range celdaInicioAux = null;
                string colAuxLetra = "";

                try
                {
                    seleccionAuxiliar = (Excel.Range)resAuxiliar;
                    celdaInicioAux = (Excel.Range)seleccionAuxiliar.Cells[1, 1];
                    colAuxLetra = celdaInicioAux.Address.Split('$')[1];
                }
                finally
                {
                    ExcelHelper.LiberarCom(celdaInicioAux);
                    ExcelHelper.LiberarCom(seleccionAuxiliar);
                }

                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 3: PROCESAMIENTO MASIVO POR ÁREAS (SOPORTE MULTI-RANGO)
                // =========================================================================
                foreach (Excel.Range area in rangoCapturado.Areas)
                {
                    Excel.Range primeraCeldaCap = null;
                    Excel.Range rangoAuxiliarArea = null;
                    Excel.Range primeraCeldaAux = null;

                    try
                    {
                        area.Validation.Delete();
                        if (limpiarFormatos) { area.FormatConditions.Delete(); }

                        int filaInicio = area.Row;
                        int filaFin = filaInicio + area.Rows.Count - 1;
                        rangoAuxiliarArea = wsActual.Range[$"{colAuxLetra}{filaInicio}:{colAuxLetra}{filaFin}"];

                        primeraCeldaCap = (Excel.Range)area.Cells[1, 1];
                        string dirCapRel = primeraCeldaCap.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                        primeraCeldaAux = (Excel.Range)rangoAuxiliarArea.Cells[1, 1];
                        string dirAuxRel = primeraCeldaAux.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                        string permitidos = "0123456789ABCDEFGHIJKLMNÑOPQRSTUVWXYZÁÉÍÓÚÜ ";
                        string formulaAuxiliar = $"=IF(OR(ISBLANK({dirCapRel}), AND(EXACT({dirCapRel},UPPER({dirCapRel})), LEN({dirCapRel})=LEN(TRIM({dirCapRel})), SUMPRODUCT(--ISNUMBER(FIND(MID({dirCapRel},ROW(INDIRECT(\"1:\"&MAX(1,LEN({dirCapRel})))),1),\"{permitidos}\")))=LEN({dirCapRel}))), 1, 0)";

                        rangoAuxiliarArea.Formula = formulaAuxiliar;

                        // TRUCO ARQUITECTÓNICO: BYPASS DEL ERROR 0x800A03EC
                        object valorOriginal = primeraCeldaCap.Value2;
                        bool estabaVacia = (valorOriginal == null || string.IsNullOrWhiteSpace(valorOriginal.ToString()));

                        if (estabaVacia)
                        {
                            primeraCeldaCap.Value2 = "A";
                        }

                        // VALIDACIÓN DE DATOS (ULTRA LIGERA)
                        // Aquí usamos ExcelHelper para la traducción de la regla de Data Validation,
                        // previniendo problemas en Excels regionales.
                        string formulaDV = ExcelHelper.TraducirFormulaLocal(wsActual, $"={dirAuxRel}=1", filaInicio);

                        area.Validation.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween, // En este caso particular, la fórmula auxiliar simple evalúa a TRUE/FALSE, pero para consistencia de tu código lo dejamos así. Type.Missing es preferible para Custom Validation.
                            formulaDV,
                            Type.Missing);

                        area.Validation.IgnoreBlank = true;
                        area.Validation.ShowError = true;
                        area.Validation.ErrorTitle = "Formato de texto inválido";
                        area.Validation.ErrorMessage = "El texto debe cumplir estas reglas:\n\n" +
                                                       "• Solo se permite texto en MAYÚSCULAS y NÚMEROS.\n" +
                                                       "• Sin espacios dobles o sobrantes.\n" +
                                                       "• Sin comillas ni signos de puntuación, paréntesis ni caracteres especiales.";

                        if (estabaVacia)
                        {
                            primeraCeldaCap.Value2 = null;
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(primeraCeldaCap);
                        ExcelHelper.LiberarCom(primeraCeldaAux);
                        ExcelHelper.LiberarCom(rangoAuxiliarArea);
                    }
                }

                // =========================================================================
                // FASE 4: CONCLUSIÓN Y REGISTRO EN BITÁCORA
                // =========================================================================
                string estadoAuditoria = limpiarFormatos ? "Se borraron configuraciones visuales previas." : "Se conservaron colores y bloqueos anteriores.";
                string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    "Formato de Texto (Alfanumérico)",
                    rangoCapturado.Address.Replace("$", ""),
                    "Data Validation"
                );

                MessageBox.Show(
                    $"Validación de formato de texto aplicada con éxito en {rangoCapturado.Areas.Count} bloque(s).\n\n" +
                    $"• Columna Auxiliar Inyectada: {colAuxLetra}\n" +
                    $"• Auditoría: {estadoAuditoria}",
                    "SAVCNG - Validación Completada", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error crítico al configurar Formato Texto: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ExcelHelper.LiberarCom(wsActual);
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}