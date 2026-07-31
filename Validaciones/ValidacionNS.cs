using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Consumimos nuestra Fachada

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionNS : IValidacionExcel
    {
        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                // =========================================================================
                // FASE 1: UX DE CONFIGURACIÓN GLOBAL Y CONTROL DE ABORTO
                // =========================================================================
                string textoSugerido = "Alerta: justificar el uso de NS.";
                object resTexto = excelApp.InputBox(
                    "Escribe el texto del mensaje de alerta que se aplicará a las celdas seleccionadas:",
                    "SAVCNG - Configuración Masiva NS (Escucha Pasiva)", textoSugerido, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                // ABORTO TOTAL 1: Si cancela el texto global
                if (resTexto is bool && (bool)resTexto == false)
                {
                    MessageBox.Show("Proceso cancelado por el usuario.\n\nNo se incluyó ninguna validación en la plantilla.", "SAVCNG - Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string textoAlerta = resTexto.ToString().Trim();

                // =========================================================================
                // FASE 2: AUDITORÍA DE COEXISTENCIA (DETECCIÓN DE REGLAS PREVIAS)
                // =========================================================================
                bool tieneValidacionesPrevias = false;

                if (rangoCapturado.FormatConditions.Count > 0)
                {
                    tieneValidacionesPrevias = true;
                }

                if (!tieneValidacionesPrevias)
                {
                    try
                    {
                        var tipoValidacion = rangoCapturado.Validation.Type;
                        tieneValidacionesPrevias = true;
                    }
                    catch { /* Silenciado intencionalmente: Significa que la celda está limpia */ }
                }

                bool limpiarPrevias = false;

                if (tieneValidacionesPrevias)
                {
                    DialogResult respLimpieza = MessageBox.Show(
                        "Se han detectado validaciones de datos o formatos condicionales PREVIOS en el rango capturado.\n\n" +
                        "¿Deseas CONSERVAR lo anterior e integrar la validación NS por debajo?\n\n" +
                        "SÍ = MANTENER reglas previas (Recomendado para apilar reglas).\n" +
                        "NO = BORRAR TODO lo anterior y dejar las celdas limpias.",
                        "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                    // ABORTO TOTAL 2
                    if (respLimpieza == DialogResult.Cancel)
                    {
                        MessageBox.Show("Proceso cancelado.\n\nNo se alteró la plantilla.", "SAVCNG - Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (respLimpieza == DialogResult.No)
                    {
                        limpiarPrevias = true;
                    }
                }

                // =========================================================================
                // FASE 3: ALGORITMO DE AGRUPACIÓN POR PREGUNTA (MATRICES PARALELAS)
                // =========================================================================
                var areasPorPregunta = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<Excel.Range>>();

                for (int i = 1; i <= rangoCapturado.Areas.Count; i++)
                {
                    Excel.Range areaIndividual = (Excel.Range)rangoCapturado.Areas[i];
                    string idPregunta = ExcelHelper.ObtenerNumeroPregunta(areaIndividual); // Uso de Façade

                    if (!areasPorPregunta.ContainsKey(idPregunta))
                    {
                        areasPorPregunta[idPregunta] = new System.Collections.Generic.List<Excel.Range>();
                    }
                    areasPorPregunta[idPregunta].Add(areaIndividual);
                }

                int preguntasProcesadas = 0;
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 4: PROCESAMIENTO INTELIGENTE (ÚNICO VS INDIVIDUAL)
                // =========================================================================
                foreach (var grupo in areasPorPregunta)
                {
                    string preguntaActual = grupo.Key;
                    System.Collections.Generic.List<Excel.Range> listaAreas = grupo.Value;

                    bool usarUnicaAlerta = true;

                    if (listaAreas.Count > 1)
                    {
                        excelApp.ScreenUpdating = true;
                        DialogResult respArquitectura = MessageBox.Show(
                            $"Se han detectado {listaAreas.Count} rangos seleccionados para la PREGUNTA {preguntaActual}.\n\n" +
                            $"¿Deseas configurar una SOLA ALERTA CENTRAL para todos estos rangos?\n\n" +
                            $"SÍ = Se pedirá 1 sola celda de alerta que evaluará todos los rangos juntos.\n" +
                            $"NO = Se pedirán {listaAreas.Count} celdas de alerta (una por cada rango de forma independiente).",
                            $"SAVCNG - Arquitectura P.{preguntaActual}", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                        excelApp.ScreenUpdating = false;

                        // ABORTO TOTAL 3
                        if (respArquitectura == DialogResult.Cancel)
                        {
                            MessageBox.Show("Proceso cancelado durante la configuración.\n\nNo se alteró la plantilla.", "SAVCNG - Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        usarUnicaAlerta = (respArquitectura == DialogResult.Yes);
                    }

                    // -------------------------------------------------------------------------
                    // RUTA A: ALERTA ÚNICA CENTRALIZADA
                    // -------------------------------------------------------------------------
                    if (usarUnicaAlerta)
                    {
                        Excel.Range rangoAlertaCentral = null;
                        try
                        {
                            excelApp.ScreenUpdating = true;
                            object resDestino = excelApp.InputBox(
                                $"[PREGUNTA DETECTADA: {preguntaActual}] - MODO CENTRALIZADO\n\n" +
                                $"Selecciona la celda destino donde aparecerá el mensaje de ALERTA para los {listaAreas.Count} rangos:\n" +
                                $"(Selección Estándar para mensajes: Columna B hasta AD)",
                                $"SAVCNG - Alerta Central P.{preguntaActual}", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                            excelApp.ScreenUpdating = false;

                            // ABORTO TOTAL 4
                            if (resDestino is bool && (bool)resDestino == false)
                            {
                                MessageBox.Show("Mapeo cancelado por el usuario.\n\nEl proceso ha sido abortado por completo.", "SAVCNG - Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }

                            rangoAlertaCentral = (Excel.Range)resDestino;
                            System.Collections.Generic.List<string> fragmentosCountIf = new System.Collections.Generic.List<string>();

                            foreach (Excel.Range area in listaAreas)
                            {
                                if (limpiarPrevias)
                                {
                                    area.Validation.Delete();
                                    area.FormatConditions.Delete();
                                }

                                Excel.Range primeraCeldaArea = null;
                                Excel.FormatCondition fcNS = null;

                                try
                                {
                                    primeraCeldaArea = (Excel.Range)area.Cells[1, 1];
                                    string dirRelativaArea = primeraCeldaArea.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                                    // Traducción delegada a la Fachada (Sustituye el Dummy Cell manual)
                                    string formulaNSLocal = ExcelHelper.TraducirFormulaLocal(wsActual, $"=TRIM({dirRelativaArea})=\"NS\"", area.Row);

                                    fcNS = (Excel.FormatCondition)area.FormatConditions.Add(
                                        Excel.XlFormatConditionType.xlExpression,
                                        Type.Missing,
                                        formulaNSLocal);

                                    fcNS.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 143, 0));
                                    fcNS.Font.Bold = true;
                                    fcNS.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204));
                                    fcNS.StopIfTrue = false;
                                }
                                finally
                                {
                                    ExcelHelper.LiberarCom(fcNS);
                                    ExcelHelper.LiberarCom(primeraCeldaArea);
                                }

                                string addrAbsoluta = area.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                                fragmentosCountIf.Add($"COUNTIF({addrAbsoluta},\"NS\")");
                            }

                            if (rangoAlertaCentral.Count > 1) { rangoAlertaCentral.Merge(); }
                            rangoAlertaCentral.Font.Name = "Arial";
                            rangoAlertaCentral.Font.Size = 9;
                            rangoAlertaCentral.Font.Bold = true;
                            rangoAlertaCentral.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 143, 0));
                            rangoAlertaCentral.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            rangoAlertaCentral.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                            string sumaInner = string.Join(",", fragmentosCountIf);
                            string formulaFinalAlerta = $"=IF(SUM({sumaInner})>0, \"{textoAlerta}\", \"\")";
                            rangoAlertaCentral.Formula = formulaFinalAlerta;
                        }
                        finally
                        {
                            ExcelHelper.LiberarCom(rangoAlertaCentral);
                        }
                    }
                    // -------------------------------------------------------------------------
                    // RUTA B: ALERTAS INDIVIDUALES
                    // -------------------------------------------------------------------------
                    else
                    {
                        for (int j = 0; j < listaAreas.Count; j++)
                        {
                            Excel.Range area = listaAreas[j];
                            Excel.Range rangoAlertaIndiv = null;

                            try
                            {
                                excelApp.ScreenUpdating = true;
                                object resDestino = excelApp.InputBox(
                                    $"[PREGUNTA DETECTADA: {preguntaActual}] - Rango {j + 1} de {listaAreas.Count}\n\n" +
                                    $"Selecciona la celda destino donde aparecerá el mensaje EXCLUSIVO para este rango:",
                                    $"SAVCNG - Alerta Individual P.{preguntaActual} (Área {j + 1})", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                                excelApp.ScreenUpdating = false;

                                // ABORTO TOTAL 5
                                if (resDestino is bool && (bool)resDestino == false)
                                {
                                    MessageBox.Show("Mapeo cancelado por el usuario.\n\nEl proceso ha sido abortado por completo.", "SAVCNG - Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    return;
                                }

                                rangoAlertaIndiv = (Excel.Range)resDestino;

                                if (limpiarPrevias)
                                {
                                    area.Validation.Delete();
                                    area.FormatConditions.Delete();
                                }

                                Excel.Range primeraCeldaArea = null;
                                Excel.FormatCondition fcNS = null;

                                try
                                {
                                    primeraCeldaArea = (Excel.Range)area.Cells[1, 1];
                                    string dirRelativaArea = primeraCeldaArea.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                                    // Traducción delegada a la Fachada
                                    string formulaNSLocal = ExcelHelper.TraducirFormulaLocal(wsActual, $"=TRIM({dirRelativaArea})=\"NS\"", area.Row);

                                    fcNS = (Excel.FormatCondition)area.FormatConditions.Add(
                                        Excel.XlFormatConditionType.xlExpression,
                                        Type.Missing,
                                        formulaNSLocal);

                                    fcNS.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 143, 0));
                                    fcNS.Font.Bold = true;
                                    fcNS.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204));
                                    fcNS.StopIfTrue = false;
                                }
                                finally
                                {
                                    ExcelHelper.LiberarCom(fcNS);
                                    ExcelHelper.LiberarCom(primeraCeldaArea);
                                }

                                if (rangoAlertaIndiv.Count > 1) { rangoAlertaIndiv.Merge(); }
                                rangoAlertaIndiv.Font.Name = "Arial";
                                rangoAlertaIndiv.Font.Size = 9;
                                rangoAlertaIndiv.Font.Bold = true;
                                rangoAlertaIndiv.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 143, 0));
                                rangoAlertaIndiv.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                rangoAlertaIndiv.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                                string addrAbsoluta = area.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                                string formulaFinalAlerta = $"=IF(COUNTIF({addrAbsoluta},\"NS\")>0, \"{textoAlerta}\", \"\")";
                                rangoAlertaIndiv.Formula = formulaFinalAlerta;
                            }
                            finally
                            {
                                ExcelHelper.LiberarCom(rangoAlertaIndiv);
                            }
                        }
                    }

                    preguntasProcesadas++;
                }

                // =========================================================================
                // FASE 5: CONCLUSIÓN Y NOTIFICACIÓN UX
                // =========================================================================
                string estadoPrevias = limpiarPrevias ? "Se borraron validaciones anteriores." : "Se conservaron validaciones anteriores.";
                string preguntasDetectadas = string.Join(", ", areasPorPregunta.Keys);

                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    "Validación NS",
                    rangoCapturado.Address.Replace("$", ""),
                    "Fórmulas / FormatCondition"
                );

                MessageBox.Show(
                    $"Procesamiento masivo finalizado con éxito.\n\n" +
                    $"• Preguntas procesadas: {preguntasProcesadas}\n" +
                    $"• Total de rangos mapeados: {rangoCapturado.Areas.Count}\n" +
                    $"• Auditoría: {estadoPrevias}\n\n" +
                    $"Nota: Las celdas admiten cualquier tipo de dato y se resaltarán en color ORO si detectan 'NS'.",
                    "SAVCNG Automatización Masiva", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error crítico en el motor de masificación: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ExcelHelper.LiberarCom(wsActual);
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}