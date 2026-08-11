using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Consumimos nuestra Fachada y el DTO

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionFechas : IValidacionExcel
    {
        // 1. EL CONTRATO AHORA EXIGE DEVOLVER EL DTO (ResultadoValidacion)
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                // =========================================================================
                // FASE 1: UX DE CONFIGURACIÓN DE LÍMITES Y CONTROL DE ABORTO
                // =========================================================================
                object resultadoInferior = excelApp.InputBox(
                    "Indica el valor MÍNIMO aceptado para esta validación:\n\n(Ej. 1 para días/meses, o 1821 para años).",
                    "SAVCNG - Límite Inferior", "1", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                // ABORTO 1: DTO
                if (resultadoInferior is bool && (bool)resultadoInferior == false)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Configuración de límite mínimo cancelada.", AlertaInyectada = false };
                }

                string inputInferior = resultadoInferior.ToString().Trim();
                if (string.IsNullOrWhiteSpace(inputInferior))
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Límite inferior vacío. Operación abortada.", AlertaInyectada = false };
                }

                object resultadoSuperior = excelApp.InputBox(
                    "Indica el valor MÁXIMO aceptado para esta validación:\n\n(Ej. 31 para días, 12 para meses, o 2026 para años).",
                    "SAVCNG - Límite Superior", "2026", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                // ABORTO 2: DTO
                if (resultadoSuperior is bool && (bool)resultadoSuperior == false)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Configuración de límite máximo cancelada.", AlertaInyectada = false };
                }

                string inputSuperior = resultadoSuperior.ToString().Trim();
                if (string.IsNullOrWhiteSpace(inputSuperior))
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Límite superior vacío. Operación abortada.", AlertaInyectada = false };
                }

                // Validación estricta de las variables C#
                if (!int.TryParse(inputInferior, out int limiteInferior) || !int.TryParse(inputSuperior, out int limiteSuperior))
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Por favor, asegúrate de escribir únicamente números enteros.\nNo se permiten letras, decimales, ni espacios en blanco.", AlertaInyectada = false };
                }

                if (limiteInferior > limiteSuperior)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = $"Error de Lógica:\nEl límite mínimo ({limiteInferior}) no puede ser mayor que el límite máximo ({limiteSuperior}).", AlertaInyectada = false };
                }

                // Forzamos un mínimo de 2 dígitos por si el límite superior fuera de 1 dígito (ej. 9).
                string mascaraFormato = new String('0', Math.Max(2, inputSuperior.Length));

                // =========================================================================
                // FASE 2: AUDITORÍA DE COEXISTENCIA (ADAPTADA A DATA VALIDATION)
                // =========================================================================
                bool tieneFormatosPrevios = rangoCapturado.FormatConditions.Count > 0;
                bool tieneValidacionPrevia = false;

                try { var tipo = rangoCapturado.Validation.Type; tieneValidacionPrevia = true; } catch { /* Silenciado: No hay DataValidation previa */ }

                bool limpiarFormatos = false;

                if (tieneFormatosPrevios || tieneValidacionPrevia)
                {
                    // Interacción permitida
                    DialogResult respLimpieza = MessageBox.Show(
                        "Se detectaron configuraciones previas en el rango seleccionado.\n\n" +
                        "NOTA: Al ser una restricción de captura estricta, la regla de celdas se sobreescribirá, pero...\n\n" +
                        "¿Deseas CONSERVAR los colores, alertas o bloqueos (Formatos Condicionales) aplicados previamente?\n\n" +
                        "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                        "NO = BORRAR todo el historial y limpiar el lienzo.",
                        "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                    // ABORTO 3: DTO
                    if (respLimpieza == DialogResult.Cancel)
                    {
                        return new ResultadoValidacion { Exito = false, Mensaje = "Proceso cancelado.\nNo se alteró la plantilla.", AlertaInyectada = false };
                    }

                    if (respLimpieza == DialogResult.No)
                    {
                        limpiarFormatos = true;
                    }
                }

                // =========================================================================
                // FASE 2.5: RECOLECCIÓN LIGERA DE PREGUNTAS (STATE MANAGEMENT)
                // =========================================================================
                System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();
                int totalAreas = rangoCapturado.Areas.Count;

                // RASTREO ZERO LEAKS: Envolviendo la iteración con bloque finally
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
                        ExcelHelper.LiberarCom(areaIndividual); // Limpieza estricta COM
                    }
                }

                string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 3: PROCESAMIENTO MASIVO POR ÁREAS Y TRADUCCIÓN UNIVERSAL
                // =========================================================================

                // RASTREO ZERO LEAKS: Reemplazo del foreach por ciclo for indexado
                for (int j = 1; j <= totalAreas; j++)
                {
                    Excel.Range area = null;
                    Excel.Range primeraCelda = null;

                    try
                    {
                        area = (Excel.Range)rangoCapturado.Areas[j];

                        area.Validation.Delete();
                        if (limpiarFormatos) { area.FormatConditions.Delete(); }

                        // ---NUEVO AJUSTE: Aplicar el formato numérico ---
                       area.NumberFormat = mascaraFormato;

                        primeraCelda = (Excel.Range)area.Cells[1, 1];
                        string direccionRelativa = primeraCelda.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                        string formulaValidacionIngles = $"=IF(ISNUMBER({direccionRelativa}), AND({direccionRelativa}>={limiteInferior}, {direccionRelativa}<={limiteSuperior}, TRUNC({direccionRelativa})={direccionRelativa}), OR(TRIM({direccionRelativa})=\"NS\", TRIM({direccionRelativa})=\"NA\"))";

                        // Traducción vía Façade (Elude el Error 0x800A03EC de manera nativa sin el DummyCell manual)
                        string formulaValidacionLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaValidacionIngles, area.Row);

                        area.Validation.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            formulaValidacionLocal,
                            Type.Missing);

                        area.Validation.IgnoreBlank = true;
                        area.Validation.ShowError = true;
                        area.Validation.ErrorTitle = "Captura Inválida (Solo Enteros)";
                        area.Validation.ErrorMessage = $"El formato de esta celda no admite números con decimales.\n\nEl número entero debe estar entre {limiteInferior} y {limiteSuperior}.\n\nTambién puedes usar las claves permitidas 'NS' o 'NA'.";
                    }
                    finally
                    {
                        // Liberación estricta del proxy de área y sus dependencias en cada vuelta
                        ExcelHelper.LiberarCom(primeraCelda);
                        ExcelHelper.LiberarCom(area);
                    }
                }

                // =========================================================================
                // FASE 4: CONCLUSIÓN Y NOTIFICACIÓN UX (EMPACADA EN DTO)
                // =========================================================================
                string estadoAuditoria = limpiarFormatos ? "Se borraron configuraciones visuales previas." : "Se conservaron colores y bloqueos anteriores.";

                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    "Rango Numérico (Fechas)",
                    rangoCapturado.Address.Replace("$", ""),
                    "Data Validation (Custom)"
                );

                string mensajeExito = $"Validación de rango aplicada con éxito en {totalAreas} bloque(s).\n\n" +
                                      $"• Criterio: Números enteros del {limiteInferior} al {limiteSuperior} (o claves NS/NA).\n" +
                                      $"• Auditoría: {estadoAuditoria}";

                // ÉXITO TOTAL: DTO
                return new ResultadoValidacion
                {
                    Exito = true,
                    Mensaje = mensajeExito,
                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                // ERROR CRÍTICO: DTO
                return new ResultadoValidacion
                {
                    Exito = false,
                    Mensaje = $"Error crítico en el motor de fechas: {ex.Message}",
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