using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Consumimos nuestra Fachada y el DTO

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionDecimales : IValidacionExcel
    {
        // 1. EL CONTRATO AHORA EXIGE DEVOLVER EL DTO
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                // =========================================================================
                // FASE 1: UX DE DECISIÓN (Bucle de captura segura)
                // =========================================================================
                string opcionEscogida = "";
                while (true)
                {
                    object seleccionTipo = excelApp.InputBox(
                        "Ingresa el NÚMERO de la regla que deseas aplicar:\n\n" +
                        "1 = Solo números ENTEROS.\n" +
                        "2 = Números DECIMALES (Hasta 10 posiciones).",
                        "SAVCNG - Configuración Numérica",
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1);

                    // ABORTO TOTAL 1: Retornamos el DTO en lugar del return vacío
                    if (seleccionTipo is bool && (bool)seleccionTipo == false)
                    {
                        return new ResultadoValidacion
                        {
                            Exito = false,
                            Mensaje = "Configuración cancelada por el usuario.",
                            AlertaInyectada = false
                        };
                    }

                    opcionEscogida = seleccionTipo.ToString().Trim();

                    if (opcionEscogida == "1" || opcionEscogida == "2")
                    {
                        break;
                    }
                    else
                    {
                        // Este MessageBox se permite por ser interactivo (Validación de entrada)
                        MessageBox.Show("Opción no válida. Por favor ingresa el número 1 o 2.", "Dato Incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                bool esValidacionEnteros = (opcionEscogida == "1");
                string etiquetaBitacora = $"Validación {(esValidacionEnteros ? "Enteros" : "Decimales")} (Positivos)";
                string tituloError = esValidacionEnteros ? "Solo números enteros" : "Formato decimal inválido";
                string mensajeError = esValidacionEnteros
                    ? "El formato de esta celda no admite texto, decimales, ni números negativos.\n\nPor favor, introduce únicamente un NÚMERO ENTERO POSITIVO (incluyendo el 0) o las claves 'NS' y 'NA'."
                    : "El formato de esta celda exige NÚMEROS POSITIVOS (incluyendo el 0) con un máximo de 10 posiciones decimales.\n\nPor favor, introduce un número válido o las claves 'NS' y 'NA'.";

                // =========================================================================
                // FASE 2: AUDITORÍA DE COEXISTENCIA
                // =========================================================================
                bool tieneFormatosPrevios = rangoCapturado.FormatConditions.Count > 0;
                bool tieneValidacionPrevia = false;

                try { var tipo = rangoCapturado.Validation.Type; tieneValidacionPrevia = true; } catch { /* Silenciado */ }

                bool limpiarFormatos = false;

                if (tieneFormatosPrevios || tieneValidacionPrevia)
                {
                    excelApp.ScreenUpdating = true; // Restaurar UX para el diálogo
                    DialogResult respLimpieza = MessageBox.Show(
                        "Se detectaron configuraciones previas en el rango seleccionado.\n\n" +
                        "NOTA: Al ser una restricción de captura estricta, la regla se sobreescribirá, pero...\n\n" +
                        "¿Deseas CONSERVAR los colores, alertas o bloqueos (Formatos Condicionales) aplicados previamente?\n\n" +
                        "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                        "NO = BORRAR todo el historial y limpiar el lienzo.",
                        "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                    excelApp.ScreenUpdating = false;

                    // ABORTO TOTAL 2: Retorno vía DTO
                    if (respLimpieza == DialogResult.Cancel)
                    {
                        return new ResultadoValidacion
                        {
                            Exito = false,
                            Mensaje = "Proceso cancelado. No se alteró la plantilla."
                        };
                    }
                    if (respLimpieza == DialogResult.No) limpiarFormatos = true;
                }

                // =========================================================================
                // FASE 3: RECOLECCIÓN LIGERA DE PREGUNTAS (Usando ExcelHelper)
                // =========================================================================
                System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();
                int totalAreas = rangoCapturado.Areas.Count;

                for (int i = 1; i <= totalAreas; i++)
                {
                    Excel.Range areaIndividual = null;
                    try
                    {
                        areaIndividual = (Excel.Range)rangoCapturado.Areas[i];
                        string idPregunta = ExcelHelper.ObtenerNumeroPregunta(areaIndividual);

                        if (!string.IsNullOrEmpty(idPregunta))
                        {
                            preguntasUnicas.Add(idPregunta);
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(areaIndividual);
                    }
                }

                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 4: PROCESAMIENTO MASIVO POR ÁREAS Y TRADUCCIÓN (FAÇADE)
                // =========================================================================
                // RASTREO ZERO LEAKS: Convertido de foreach a for para blindar la memoria
                for (int j = 1; j <= totalAreas; j++)
                {
                    Excel.Range area = null;
                    Excel.Range primeraCelda = null;
                    Excel.FormatCondition fcSecundario = null;

                    try
                    {
                        area = (Excel.Range)rangoCapturado.Areas[j];

                        area.Validation.Delete();
                        if (limpiarFormatos) { area.FormatConditions.Delete(); }

                        if (esValidacionEnteros)
                        {
                            area.NumberFormat = "#,##0";
                        }
                        else
                        {
                            area.NumberFormat = "General";
                        }

                        primeraCelda = (Excel.Range)area.Cells[1, 1];
                        string dirRel = primeraCelda.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                        int filaBase = area.Row;

                        // Construcción Universal
                        string formulaBaseIngles = esValidacionEnteros
                            ? $"IF(ISNUMBER({dirRel}), AND(TRUNC({dirRel})={dirRel}, {dirRel}>=0), OR(TRIM({dirRel})=\"NS\", TRIM({dirRel})=\"NA\"))"
                            : $"IF(ISNUMBER({dirRel}), AND(ROUND({dirRel}, 10)={dirRel}, {dirRel}>=0), OR(TRIM({dirRel})=\"NS\", TRIM({dirRel})=\"NA\"))";

                        // Delegamos la traducción al ExcelHelper
                        string formulaValidacionLocal = ExcelHelper.TraducirFormulaLocal(wsActual, $"={formulaBaseIngles}", filaBase);
                        string formulaCondicionLocal = ExcelHelper.TraducirFormulaLocal(wsActual, $"=AND({dirRel}<>\"\", NOT({formulaBaseIngles}))", filaBase);

                        // Inyección Data Validation
                        area.Validation.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Type.Missing,
                            formulaValidacionLocal,
                            Type.Missing);

                        area.Validation.IgnoreBlank = true;
                        area.Validation.ShowError = true;
                        area.Validation.ErrorTitle = tituloError;
                        area.Validation.ErrorMessage = mensajeError;

                        // Inyección Formato Condicional
                        fcSecundario = (Excel.FormatCondition)area.FormatConditions.Add(
                            Excel.XlFormatConditionType.xlExpression,
                            Type.Missing,
                            formulaCondicionLocal);

                        fcSecundario.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 199, 206));
                        fcSecundario.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));
                        fcSecundario.StopIfTrue = false;
                    }
                    finally
                    {
                        // Limpieza estricta COM iterativa
                        ExcelHelper.LiberarCom(fcSecundario);
                        ExcelHelper.LiberarCom(primeraCelda);
                        ExcelHelper.LiberarCom(area);
                    }
                }

                // =========================================================================
                // FASE 5: CONCLUSIÓN Y REGISTRO EN BITÁCORA (EMPACADO EN DTO)
                // =========================================================================
                string estadoAuditoria = limpiarFormatos ? "Se borraron configuraciones visuales previas." : "Se conservaron colores y bloqueos anteriores.";
                string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    etiquetaBitacora,
                    rangoCapturado.Address.Replace("$", ""),
                    "Validation + FormatCondition"
                );

                string mensajeExito = $"{etiquetaBitacora} aplicada con éxito en {totalAreas} bloque(s).\n\n" +
                                      $"• Auditoría: {estadoAuditoria}\n" +
                                      $"• Regla Activa: Bloqueo de captura inválida (Data Validation).\n" +
                                      $"• Regla Pasiva: Resalte rojo en caso de alteración externa (Format Conditions).";

                // ÉXITO TOTAL: Retorno del DTO a la Vista
                return new ResultadoValidacion
                {
                    Exito = true,
                    Mensaje = mensajeExito,
                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                // ERROR CRÍTICO: Retorno del DTO
                return new ResultadoValidacion
                {
                    Exito = false,
                    Mensaje = "Error crítico al aplicar la validación: " + ex.Message,
                    AlertaInyectada = false
                };
            }
            finally
            {
                ExcelHelper.LiberarCom(wsActual);
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}