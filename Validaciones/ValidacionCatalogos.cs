using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Consumimos la Fachada y el DTO

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionCatalogos : IValidacionExcel
    {
        // 1. EL CONTRATO AHORA EXIGE DEVOLVER EL DTO
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            // PARCHE ZERO LEAKS: Envolvemos TODO en un Try-Catch-Finally maestro
            try
            {
                string formulaOpciones = "";
                string separador = excelApp.International[Excel.XlApplicationInternational.xlListSeparator].ToString();

                // Interacción permitida: Pregunta de decisión al usuario
                DialogResult tipoEntrada = MessageBox.Show(
                    "¿Deseas escribir el(los) valor(es) del catálogo manualmente (ej: un solo valor como 'X' o varios como '1,2,9')?\n\n" +
                    "SÍ: Escribir el(los) valor(es) directamente.\n" +
                    "NO: Seleccionar celdas de Excel.",
                    "Origen del Catálogo",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (tipoEntrada == DialogResult.Yes)
                {
                    // ==========================================================
                    // CASO 1: ENTRADA MANUAL (Ideal para "X")
                    // ==========================================================
                    object resultadoTexto = excelApp.InputBox(
                        "Escribe el(los) valor(es) para tu lista desplegable.\nNOTA: Si son varias deberan estar separadas por comas):",
                        "Escribir Opciones",
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                    // ABORTO: Empaquetado en el DTO
                    if (resultadoTexto is bool && (bool)resultadoTexto == false)
                    {
                        return new ResultadoValidacion { Exito = false, Mensaje = "Configuración manual cancelada por el usuario.", AlertaInyectada = false };
                    }

                    string textoEscrito = resultadoTexto.ToString().Trim();
                    if (string.IsNullOrEmpty(textoEscrito))
                    {
                        return new ResultadoValidacion { Exito = false, Mensaje = "No se ingresó ningún valor. Proceso abortado.", AlertaInyectada = false };
                    }

                    // Reemplaza comas por el separador correcto de la PC
                    formulaOpciones = textoEscrito.Replace(",", separador);
                }
                else
                {
                    // ==========================================================
                    // CASO 2: SELECCIÓN DE CELDAS
                    // ==========================================================
                    object resultadoInput = excelApp.InputBox(
                        "Selecciona el rango de opciones o la celda que contiene el catálogo (ej: 1,2,9):",
                        "Seleccionar Origen del Catálogo",
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

                    if (resultadoInput is bool && (bool)resultadoInput == false)
                    {
                        return new ResultadoValidacion { Exito = false, Mensaje = "Selección de rango cancelada por el usuario.", AlertaInyectada = false };
                    }

                    Excel.Range rangoOrigen = null;
                    try
                    {
                        rangoOrigen = (Excel.Range)resultadoInput;
                        bool esCeldaUnicaOCombinada = (rangoOrigen.Count == 1) || (bool)rangoOrigen.MergeCells;

                        if (esCeldaUnicaOCombinada)
                        {
                            DialogResult respuesta = MessageBox.Show(
                                "Se ha detectado un único bloque de texto como catálogo.\n" +
                                "¿Deseas que el sistema extraiga automáticamente SOLO LOS NÚMEROS (ej. 1, 2, 9) para crear las opciones?\n\n" +
                                "• [Aceptar]: Extraer números y continuar.\n" +
                                "• [Cancelar]: Abortar esta operación.",
                                "Configuración de Catálogo",
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Question);

                            if (respuesta == DialogResult.Cancel)
                            {
                                return new ResultadoValidacion { Exito = false, Mensaje = "Extracción automática cancelada.", AlertaInyectada = false };
                            }

                            Excel.Range primeraCeldaOrigen = null;
                            string textoCelda = "";

                            try
                            {
                                primeraCeldaOrigen = (Excel.Range)rangoOrigen.Cells[1, 1];
                                textoCelda = primeraCeldaOrigen.Text?.ToString() ?? "";
                            }
                            finally
                            {
                                ExcelHelper.LiberarCom(primeraCeldaOrigen);
                            }

                            if (string.IsNullOrWhiteSpace(textoCelda))
                            {
                                return new ResultadoValidacion { Exito = false, Mensaje = "La celda origen está vacía. No se puede extraer el catálogo.", AlertaInyectada = false };
                            }

                            string[] pedacitos = textoCelda.Split(new char[] { '.', ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                            System.Collections.Generic.List<string> listaLimpios = new System.Collections.Generic.List<string>();

                            foreach (string pedazo in pedacitos)
                            {
                                string soloNumeros = "";
                                foreach (char letra in pedazo)
                                {
                                    if (char.IsDigit(letra)) soloNumeros += letra;
                                }
                                if (!string.IsNullOrEmpty(soloNumeros))
                                {
                                    listaLimpios.Add(soloNumeros);
                                }
                            }

                            formulaOpciones = string.Join(separador, listaLimpios);
                        }
                        else
                        {
                            DialogResult respuestaRango = MessageBox.Show(
                                "Se han detectado varias celdas seleccionadas como catálogo.\n" +
                                "¿Deseas que el sistema extraiga automáticamente SOLO LOS NÚMEROS (ej. 1, 2, 9) de cada celda para crear las opciones?\n\n" +
                                "• [Aceptar]: Extraer números y continuar.\n" +
                                "• [Cancelar]: Abortar esta operación.",
                                "Configuración de Catálogo",
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Question);

                            if (respuestaRango == DialogResult.Cancel)
                            {
                                return new ResultadoValidacion { Exito = false, Mensaje = "Extracción múltiple cancelada.", AlertaInyectada = false };
                            }

                            System.Collections.Generic.List<string> listaLimpios = new System.Collections.Generic.List<string>();
                            int totalCeldasOrigen = rangoOrigen.Cells.Count;

                            // RASTREO ZERO LEAKS: Convertido de foreach a for para proteger la RAM
                            for (int k = 1; k <= totalCeldasOrigen; k++)
                            {
                                Excel.Range celda = null;
                                try
                                {
                                    celda = (Excel.Range)rangoOrigen.Cells[k];
                                    string textoCelda = celda.Text?.ToString() ?? "";
                                    string soloNumeros = "";

                                    foreach (char letra in textoCelda)
                                    {
                                        if (char.IsDigit(letra)) soloNumeros += letra;
                                    }

                                    if (!string.IsNullOrEmpty(soloNumeros))
                                    {
                                        listaLimpios.Add(soloNumeros);
                                    }
                                }
                                finally
                                {
                                    ExcelHelper.LiberarCom(celda);
                                }
                            }

                            if (listaLimpios.Count == 0)
                            {
                                return new ResultadoValidacion { Exito = false, Mensaje = "No se encontraron valores numéricos en el rango seleccionado.\nNo se puede crear el catálogo.", AlertaInyectada = false };
                            }

                            formulaOpciones = string.Join(separador, listaLimpios);
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(rangoOrigen);
                    }
                }

                // ==========================================================
                // PREVENCIÓN DE SOBRESCRITURA DE VALIDACIÓN
                // ==========================================================
                bool tieneValidacionPrevia = false;
                Excel.Range primeraCeldaRango = null;

                try
                {
                    primeraCeldaRango = (Excel.Range)rangoCapturado.Cells[1, 1];
                    int tipoValidacion = primeraCeldaRango.Validation.Type;
                    tieneValidacionPrevia = true;
                }
                catch
                {
                    tieneValidacionPrevia = false;
                }
                finally
                {
                    ExcelHelper.LiberarCom(primeraCeldaRango);
                }

                if (tieneValidacionPrevia)
                {
                    DialogResult sobrescribir = MessageBox.Show(
                        "Las celdas que seleccionaste ya tienen una validación de datos o lista desplegable asignada.\n\n" +
                        "¿Estás seguro de que deseas borrarla y aplicar este nuevo catálogo en su lugar?",
                        "Validación existente detectada",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (sobrescribir == DialogResult.No)
                    {
                        return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada para preservar la validación existente.", AlertaInyectada = false };
                    }
                }

                // =========================================================================
                // RECOLECCIÓN LIGERA DE PREGUNTAS (STATE MANAGEMENT)
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

                string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                // ==========================================================
                // APLICAR LA VALIDACIÓN Y REGISTRO EN BITÁCORA
                // ==========================================================
                rangoCapturado.Validation.Delete();

                rangoCapturado.Validation.Add(
                    Excel.XlDVType.xlValidateList,
                    Excel.XlDVAlertStyle.xlValidAlertStop,
                    Excel.XlFormatConditionOperator.xlBetween,
                    formulaOpciones,
                    Type.Missing);

                rangoCapturado.Validation.InCellDropdown = true;
                rangoCapturado.Validation.IgnoreBlank = true;

                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    "Catálogo (Lista Desplegable)",
                    rangoCapturado.Address.Replace("$", ""),
                    "Data Validation"
                );

                // ÉXITO TOTAL: Retorno del DTO
                return new ResultadoValidacion
                {
                    Exito = true,
                    Mensaje = "¡Validación de catálogo aplicada con éxito!",
                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                // ERROR CRÍTICO: Retorno del DTO
                return new ResultadoValidacion
                {
                    Exito = false,
                    Mensaje = "Error crítico al aplicar la validación de catálogos: " + ex.Message,
                    AlertaInyectada = false
                };
            }
            finally
            {
                // Devolvemos el control visual pase lo que pase
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}