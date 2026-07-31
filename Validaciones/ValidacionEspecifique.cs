using System;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core;

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionEspecifique : IValidacionExcel
    {
        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet ws = null;
            Excel.Range rangoCatalogo = null;
            Excel.Range rangoMensaje = null;
            Excel.Range celdaMotor = null;

            try
            {
                ws = (Excel.Worksheet)rangoCapturado.Worksheet;

                // =========================================================================
                // FASE 1: RECOPILACIÓN DE ESPACIOS DE TRABAJO (UX BLINDADA)
                // =========================================================================
                object resCatalogo = excelApp.InputBox(
                   "1. Selecciona las OPCIONES DEL CATÁLOGO.\nNOTA: Debera omitir de la selección las opciones 'Otro(Especifique)' y/o 'No identificado' del catálogo correspondiente. ",
                    "Mapeo de Catálogo", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resCatalogo is bool && (bool)resCatalogo == false) return;
                rangoCatalogo = (Excel.Range)resCatalogo;

                object resMensaje = excelApp.InputBox(
                    "2. Selecciona el rango o celda donde se mostrará el MENSAJE DE ALERTA.",
                    "Destino de Alerta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resMensaje is bool && (bool)resMensaje == false) return;
                rangoMensaje = (Excel.Range)resMensaje;

                object resMotor = excelApp.InputBox(
                    "3. Selecciona UNA CELDA VACÍA (columna AF en adelante) para construir el Diccionario Auxiliar de Busqueda.\nADVERTENCIA: Considere un espacio libre de dos columnas por N filas. (N = numero de opciones del catalogo)",
                    "Generación del Motor Oculto", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resMotor is bool && (bool)resMotor == false) return;
                celdaMotor = (Excel.Range)resMotor;

                // =========================================================================
                // FASE 1.5: AUDITORÍA DE COEXISTENCIA Y STATE MANAGEMENT
                // =========================================================================
                bool limpiarFormatos = false;
                if (rangoCatalogo.FormatConditions.Count > 0)
                {
                    DialogResult respLimpieza = MessageBox.Show(
                        "Se detectaron formatos condicionales (colores/alertas) previos en el CATÁLOGO seleccionado.\n\n" +
                        "¿Deseas CONSERVAR los colores anteriores además del resaltado amarillo de esta búsqueda?\n\n" +
                        "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                        "NO = BORRAR todo el historial y limpiar el lienzo.",
                        "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                    if (respLimpieza == DialogResult.Cancel) return;
                    if (respLimpieza == DialogResult.No) { limpiarFormatos = true; }
                }

                System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();
                for (int i = 1; i <= rangoCapturado.Areas.Count; i++)
                {
                    Excel.Range areaIndividual = (Excel.Range)rangoCapturado.Areas[i];
                    string idPregunta = ExcelHelper.ObtenerNumeroPregunta(areaIndividual);
                    if (!string.IsNullOrEmpty(idPregunta)) { preguntasUnicas.Add(idPregunta); }
                    ExcelHelper.LiberarCom(areaIndividual);
                }
                string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 2: CONSTRUCCIÓN DEL MOTOR AUXILIAR MATRICIAL
                // =========================================================================
                Excel.Range primeraCeldaAzul = null;
                string dirAzulAbsoluta = "";

                try
                {
                    primeraCeldaAzul = (Excel.Range)rangoCapturado.Cells[1, 1];
                    dirAzulAbsoluta = primeraCeldaAzul.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                }
                finally
                {
                    ExcelHelper.LiberarCom(primeraCeldaAzul);
                }

                Excel.Range celdaLimpiaAzul = null;
                string dirLimpiaAbsoluta = "";

                try
                {
                    celdaLimpiaAzul = celdaMotor.Offset[0, 0];
                    string formulaLimpieza = $"=LOWER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({dirAzulAbsoluta},\"á\",\"a\"),\"é\",\"e\"),\"í\",\"i\"),\"ó\",\"o\"),\"ú\",\"u\"))";
                    celdaLimpiaAzul.Formula = formulaLimpieza;
                    celdaLimpiaAzul.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    celdaLimpiaAzul.Font.Bold = true;
                    dirLimpiaAbsoluta = celdaLimpiaAzul.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                }
                finally
                {
                    ExcelHelper.LiberarCom(celdaLimpiaAzul);
                }

                int filaMotor = 1;
                System.Collections.Generic.List<string> celdasResultadosBooleanos = new System.Collections.Generic.List<string>();

                if (limpiarFormatos) { rangoCatalogo.FormatConditions.Delete(); }

                foreach (Excel.Range celdaCat in rangoCatalogo.Cells)
                {
                    try
                    {
                        string textoOriginal = celdaCat.Text != null ? celdaCat.Text.ToString() : "";

                        if (!string.IsNullOrWhiteSpace(textoOriginal) && textoOriginal.Trim() != "")
                        {
                            string textoProcesado = Regex.Replace(textoOriginal, @"\(.*?\)", "").Trim();
                            textoProcesado = textoProcesado.ToLower()
                                .Replace("á", "a").Replace("é", "e").Replace("í", "i")
                                .Replace("ó", "o").Replace("ú", "u");

                            string[] separadores = { " y/o ", " y ", " o ", " e ", ",", "/", " a ", " ante ", " bajo ", " cabe ", " con ", " contra ", " de ", " del ", " desde ", " durante ", " en ", " entre ", " hacia ", " hasta ", " mediante ", " para ", " por ", " según ", " sin ", " sobre ", " tras " };
                            string[] palabrasClave = textoProcesado.Split(separadores, StringSplitOptions.RemoveEmptyEntries);

                            System.Collections.Generic.List<string> fragmentosSearch = new System.Collections.Generic.List<string>();
                            foreach (string palabra in palabrasClave)
                            {
                                string palabraLimpia = palabra.Trim();
                                if (palabraLimpia.Length > 2)
                                {
                                    fragmentosSearch.Add($"ISNUMBER(SEARCH(\"{palabraLimpia}\", {dirLimpiaAbsoluta}))");
                                }
                            }

                            if (fragmentosSearch.Count > 0)
                            {
                                Excel.Range filaValidacion = null;
                                Excel.Range celdaMatch = null;
                                Excel.FormatCondition fc = null;

                                try
                                {
                                    filaValidacion = celdaMotor.Offset[filaMotor, 0];
                                    filaValidacion.Value2 = textoOriginal;

                                    celdaMatch = celdaMotor.Offset[filaMotor, 1];
                                    string formulaMatch = $"=OR({string.Join(",", fragmentosSearch)})";
                                    celdaMatch.Formula = formulaMatch;

                                    string dirMatchAbsoluta = celdaMatch.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                                    celdasResultadosBooleanos.Add(dirMatchAbsoluta);

                                    // =========================================================================
                                    // FASE 3: TRUCO ARQUITECTÓNICO DE TRADUCCIÓN VÍA FAÇADE
                                    // =========================================================================
                                    string celdaCatRelativa = celdaCat.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                                    string formulaFormatCondIngles = $"=AND({celdaCatRelativa}<>\"\", {dirMatchAbsoluta}=TRUE)";

                                    // Utilizamos ExcelHelper para traducir y evitar fugas
                                    string formulaFormatCondLocal = ExcelHelper.TraducirFormulaLocal(ws, formulaFormatCondIngles, celdaCat.Row);

                                    fc = (Excel.FormatCondition)celdaCat.FormatConditions.Add(
                                        Excel.XlFormatConditionType.xlExpression,
                                        Type.Missing,
                                        formulaFormatCondLocal);

                                    fc.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                    fc.Font.Bold = true;

                                    filaMotor++;
                                }
                                finally
                                {
                                    ExcelHelper.LiberarCom(fc);
                                    ExcelHelper.LiberarCom(celdaMatch);
                                    ExcelHelper.LiberarCom(filaValidacion);
                                }
                            }
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(celdaCat); // Crucial para catálogos largos
                    }
                }

                // =========================================================================
                // FASE 4: VINCULACIÓN DEL MENSAJE DE ALERTA 
                // =========================================================================
                if (celdasResultadosBooleanos.Count > 0)
                {
                    Excel.Range celdaInicioRango = null;
                    Excel.Range celdaFinRango = null;

                    try
                    {
                        celdaInicioRango = celdaMotor.Offset[1, 1];
                        celdaFinRango = celdaMotor.Offset[filaMotor - 1, 1];

                        string addressInicio = celdaInicioRango.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                        string addressFin = celdaFinRango.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                        string formulaMensajeFinal = $"=IF(COUNTIF({addressInicio}:{addressFin}, TRUE)>0, \"Alerta: Revise el texto ingresado en el Especifique ya que podría existir en las opciones resaltadas en amarillo\", \"\")";

                        if (rangoMensaje.Count > 1) { rangoMensaje.Merge(); }
                        rangoMensaje.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        rangoMensaje.Font.Name = "Arial";
                        rangoMensaje.Font.Size = 9;
                        rangoMensaje.Font.Bold = true;
                        rangoMensaje.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 144, 0));

                        rangoMensaje.Formula = formulaMensajeFinal;
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(celdaInicioRango);
                        ExcelHelper.LiberarCom(celdaFinRango);
                    }
                }

                // =========================================================================
                // FASE 5: REGISTRO DE AUDITORÍA Y NOTIFICACIÓN
                // =========================================================================
                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    "Motor de Búsqueda (Especifique)",
                    rangoCapturado.Address.Replace("$", ""),
                    "Formato Condicional - Fórmulas"
                );

                MessageBox.Show("El Diccionario de Palabras Clave y la Alerta Inteligente fueron construidos con éxito.", "SAVCNG - ExcelDNA", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error crítico al configurar el motor de búsqueda: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ExcelHelper.LiberarCom(rangoCatalogo);
                ExcelHelper.LiberarCom(rangoMensaje);
                ExcelHelper.LiberarCom(celdaMotor);
                ExcelHelper.LiberarCom(ws);

                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}