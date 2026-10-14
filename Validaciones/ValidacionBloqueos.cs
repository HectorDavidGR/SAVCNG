using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Uso obligatorio de la Fachada y el DTO

namespace SAVCNG_ExcelDNA.Validaciones
{
    /// <summary>
    /// Módulo Principal de Bloqueos para Censos Nacionales.
    /// Arquitectura SAVCNG 2.0 - Zero Leaks & Fórmulas Universales.
    /// Modificación: Menú optimizado para Fase 1 (Código X y Numeral).
    /// </summary>
    public class ValidacionBloqueos : IValidacionExcel
    {
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            if (rangoCapturado == null)
            {
                return new ResultadoValidacion
                {
                    Exito = false,
                    Mensaje = "¡Espera! Primero debes capturar un rango de celdas en Excel.",
                    AlertaInyectada = false
                };
            }

            Excel.Worksheet wsActual = null;
            int totalBloqueosAplicados = 0;
            bool continuarBloqueos = true;

            // Lista para recolectar las celdas que contienen la suma de cada bloqueo
            List<string> celdasSumaTotales = new List<string>();

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                // =========================================================================
                // BUCLE PRINCIPAL DE BLOQUEOS (NIVEL 1 RECURRENTE)
                // =========================================================================
                while (continuarBloqueos)
                {
                    excelApp.ScreenUpdating = true;

                    object resNivel1 = excelApp.InputBox(
                        "SELECCIONE EL TIPO DE BLOQUEO A APLICAR (1 o 2):\n\n" +
                        "1 = BLOQUEO POR CÓDIGO X\n" +
                        "2 = BLOQUEO POR NUMERAL POR FILA\n",
                        "SAVCNG 2.0 - Menú Principal de Bloqueos",
                        "1", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1);

                    if (resNivel1 is bool && (bool)resNivel1 == false)
                    {
                        continuarBloqueos = false;
                        continue; // Salimos del ciclo para ir a la creación del mensaje global
                    }

                    int opcionNivel1;
                    if (!int.TryParse(resNivel1.ToString(), out opcionNivel1))
                    {
                        MessageBox.Show("Por favor ingresa una opción numérica válida (1 o 2).", "Opción Inválida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }

                    ResultadoValidacion resBloqueo = null;
                    string dirSumaParcial = string.Empty;

                    switch (opcionNivel1)
                    {
                        case 1:
                            resBloqueo = ProcesarSubmenuCodigoX(excelApp, libroCenso, wsActual, rangoCapturado, out dirSumaParcial);
                            break;
                        case 2:
                            resBloqueo = EjecutarNumeral_FilaPorFila(excelApp, libroCenso, wsActual, rangoCapturado, out dirSumaParcial);
                            break;
                        default:
                            MessageBox.Show($"Categoría [{opcionNivel1}] no válida. Selecciona 1 o 2.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            continue;
                    }

                    if (resBloqueo != null && resBloqueo.Exito)
                    {
                        totalBloqueosAplicados++;
                        if (!string.IsNullOrEmpty(dirSumaParcial))
                        {
                            celdasSumaTotales.Add(dirSumaParcial);
                        }

                        excelApp.ScreenUpdating = true;
                        DialogResult resPregunta = MessageBox.Show(
                            $"{resBloqueo.Mensaje}\n\n¿Deseas agregar otro bloqueo para esta misma pregunta?",
                            "SAVCNG 2.0 - Agregar Otro Bloqueo",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        continuarBloqueos = (resPregunta == DialogResult.Yes);
                    }
                    else
                    {
                        // Si se canceló el bloqueo actual y hay bloqueos previos
                        if (totalBloqueosAplicados > 0)
                        {
                            DialogResult resReintentar = MessageBox.Show(
                                $"{resBloqueo?.Mensaje ?? "Operación cancelada."}\n\nYa tienes {totalBloqueosAplicados} bloqueo(s) preparado(s).\n¿Deseas volver al menú principal para seleccionar otro bloqueo?",
                                "SAVCNG 2.0 - Aviso",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);
                            continuarBloqueos = (resReintentar == DialogResult.Yes);
                        }
                        else
                        {
                            return resBloqueo ?? new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada.", AlertaInyectada = false };
                        }
                    }
                } // Fin Bucle Principal

                // =========================================================================
                // CREACIÓN DEL MENSAJE GLOBAL ÚNICO AL FINALIZAR LA RECOLECCIÓN
                // =========================================================================
                if (totalBloqueosAplicados > 0 && celdasSumaTotales.Count > 0)
                {
                    excelApp.ScreenUpdating = true;

                    string mensajeDefault = "Favor de verificar que no tenga información en celdas sombreadas";
                    object resTextoMsg = excelApp.InputBox(
                        $"¡Se prepararon {totalBloqueosAplicados} bloqueos!\n\nAhora, escribe el MENSAJE DE ALERTA GLOBAL que aparecerá si cualquiera de estos bloqueos se activa:",
                        "Mensaje de alerta global", mensajeDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                    if (resTextoMsg is bool && (bool)resTextoMsg == false)
                    {
                        return new ResultadoValidacion { Exito = true, Mensaje = $"Se aplicaron {totalBloqueosAplicados} bloqueos, pero cancelaste la inyección del mensaje visual.", AlertaInyectada = false };
                    }
                    string textoMensaje = resTextoMsg.ToString().Trim();

                    object resRangoMsg = excelApp.InputBox("Selecciona la celda o rango destino donde aparecerá este MENSAJE DE ALERTA:", "Ubicación del Mensaje Único", Type: 8);
                    if (resRangoMsg is bool && (bool)resRangoMsg == false)
                    {
                        return new ResultadoValidacion { Exito = true, Mensaje = $"Se aplicaron {totalBloqueosAplicados} bloqueos, pero cancelaste la ubicación del mensaje visual.", AlertaInyectada = false };
                    }

                    Excel.Range blk_rangoMensaje = null;
                    try
                    {
                        blk_rangoMensaje = (Excel.Range)resRangoMsg;
                        excelApp.ScreenUpdating = false;

                        if (blk_rangoMensaje.Count > 1) blk_rangoMensaje.Merge();

                        string sumaGlobal = string.Join("+", celdasSumaTotales);
                        string formulaMsgUniversal = $"=IF(({sumaGlobal})>0, \"{textoMensaje}\", \"\")";

                        blk_rangoMensaje.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaMsgUniversal, blk_rangoMensaje.Row);

                        blk_rangoMensaje.Font.Name = "Arial";
                        blk_rangoMensaje.Font.Size = 9;
                        blk_rangoMensaje.Font.Bold = true;
                        blk_rangoMensaje.Font.Color = ColorTranslator.ToOle(Color.Red);

                        blk_rangoMensaje.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        blk_rangoMensaje.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        blk_rangoMensaje.WrapText = false;

                        return new ResultadoValidacion
                        {
                            Exito = true,
                            Mensaje = $"¡Proceso finalizado con éxito! Se completaron {totalBloqueosAplicados} bloqueos en un único mensaje de alerta.",
                            AlertaInyectada = true
                        };
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(blk_rangoMensaje);
                    }
                }

                return new ResultadoValidacion
                {
                    Exito = totalBloqueosAplicados > 0,
                    Mensaje = totalBloqueosAplicados > 0 ? $"Se completaron {totalBloqueosAplicados} bloqueos sin mensaje final." : "No se aplicaron bloqueos.",
                    AlertaInyectada = false
                };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion
                {
                    Exito = false,
                    Mensaje = "Error crítico durante el procesamiento: " + ex.Message,
                    AlertaInyectada = false
                };
            }
            finally
            {
                ExcelHelper.LiberarCom(wsActual);
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }

        // =========================================================================
        // MÉTODO AUXILIAR: AUDITORÍA DE COEXISTENCIA
        // =========================================================================
        private bool VerificarYLimpiarFormatosPrevios(Excel.Range rango)
        {
            if (rango == null) return true;

            bool tieneFormatos = false;
            try
            {
                // Buscamos si alguna área del rango ya tiene Formatos Condicionales
                for (int i = 1; i <= rango.Areas.Count; i++)
                {
                    Excel.Range area = rango.Areas[i];
                    if (area.FormatConditions.Count > 0)
                    {
                        tieneFormatos = true;
                        break;
                    }
                }
            }
            catch { }

            if (tieneFormatos)
            {
                DialogResult res = MessageBox.Show(
                    "Se detectaron configuraciones previas en el rango seleccionado.\n\n" +
                    "NOTA: Al ser una restricción de captura estricta, la regla de celdas se sobreescribirá, pero...\n\n" +
                    "¿Deseas CONSERVAR los colores, alertas o bloqueos (Formatos Condicionales) aplicados previamente?\n\n" +
                    "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                    "NO = BORRAR todo el historial y limpiar el lienzo.",
                    "SAVCNG - Auditoría de Coexistencia",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);

                if (res == DialogResult.Cancel)
                    return false; // El usuario abortó la operación por completo

                if (res == DialogResult.No)
                {
                    try
                    {
                        rango.FormatConditions.Delete(); // Limpiamos el lienzo
                    }
                    catch { }
                }
            }

            return true; // Continuamos con la inyección
        }

        #region SUB-SWITCH CATEGORÍA 1: BLOQUEO CÓDIGO X
        private ResultadoValidacion ProcesarSubmenuCodigoX(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Worksheet wsActual, Excel.Range rangoCapturado, out string direccionSuma)
        {
            direccionSuma = string.Empty;

            object resNivel2 = excelApp.InputBox(
                "CATEGORÍA: BLOQUEO CÓDIGO X\n" +
                "Selecciona la variante requerida:\n\n" +
                "1 = BLOQUEO MUTUO (un solo código x)\n" +
                "2 = BLOQUEO POR DISPARADOR (Simple o Matricial)\n" +
                "3 = BLOQUEO TABLA A TABLA\n",
                "Sub-Menú: Bloqueo Código X",
                "2", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1);

            if (resNivel2 is bool && (bool)resNivel2 == false)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada por el usuario.", AlertaInyectada = false };
            }

            int opcionNivel2 = Convert.ToInt32(resNivel2);

            switch (opcionNivel2)
            {
                case 1:
                    return EjecutarCodigoX_BloqueoCruzado(excelApp, libroCenso, wsActual, rangoCapturado, out direccionSuma);
                case 2:
                    return EjecutarCodigoX_DisparadorInteligente(excelApp, libroCenso, wsActual, rangoCapturado, out direccionSuma);
                case 3:
                    return EjecutarCodigoX_BloqueoRojoYGris(excelApp, libroCenso, wsActual, rangoCapturado, out direccionSuma);
                default:
                    return new ResultadoValidacion { Exito = false, Mensaje = $"Variante [{opcionNivel2}] no válida.", AlertaInyectada = false };
            }
        }

        // ----------------------------------------------------------------------------------
        // CASO 1.1: BLOQUEO CRUZADO AUTOMÁTICO (MUTUO)
        // ----------------------------------------------------------------------------------
        private ResultadoValidacion EjecutarCodigoX_BloqueoCruzado(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Worksheet wsActual, Excel.Range rangoCapturado, out string direccionSuma)
        {
            direccionSuma = string.Empty;

            if (!VerificarYLimpiarFormatosPrevios(rangoCapturado))
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada desde la Auditoría de Coexistencia.", AlertaInyectada = false };
            }

            object resAux = excelApp.InputBox("Selecciona la CELDA AUXILIAR:", "SAVCNG-CELDA AUXILIAR", Type: 8);
            if (resAux is bool && (bool)resAux == false)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada.", AlertaInyectada = false };
            }

            Excel.Range rangeInputAux = (Excel.Range)resAux;
            Excel.Range blk_celdaAuxiliar = null;
            Excel.Range celdaTituloAux = null;

            try
            {
                blk_celdaAuxiliar = (Excel.Range)rangeInputAux.Cells[1, 1];
                excelApp.ScreenUpdating = false;

                celdaTituloAux = blk_celdaAuxiliar.Offset[-1, 0];
                celdaTituloAux.Value = "SAV-BLOQUEO UN COD X";
                celdaTituloAux.Font.Name = "Arial";
                celdaTituloAux.Font.Size = 4.5;
                celdaTituloAux.Font.Bold = true;
                celdaTituloAux.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                celdaTituloAux.WrapText = true; // Ajustar texto
                celdaTituloAux.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(166, 166, 166));

                List<string> direccionesAreas = new List<string>();
                int totalAreasRango = rangoCapturado.Areas.Count;
                for (int a = 1; a <= totalAreasRango; a++)
                {
                    Excel.Range areaTemp = (Excel.Range)rangoCapturado.Areas[a];
                    direccionesAreas.Add(areaTemp.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false));
                    ExcelHelper.LiberarCom(areaTemp);
                }

                string listaCeldasSeparadas = string.Join(",", direccionesAreas);
                string formulaAuxUniversal = $"=IF(COUNTA({listaCeldasSeparadas})>1, 1, 0)";
                blk_celdaAuxiliar.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaAuxUniversal, blk_celdaAuxiliar.Row);

                string idPregunta = ExcelHelper.ObtenerNumeroPregunta(rangoCapturado);

                List<Excel.Range> listaCeldas = new List<Excel.Range>();
                for (int a = 1; a <= totalAreasRango; a++)
                {
                    Excel.Range area = (Excel.Range)rangoCapturado.Areas[a];
                    int totalFilas = area.Rows.Count;
                    int totalCols = area.Columns.Count;

                    for (int f = 1; f <= totalFilas; f++)
                    {
                        for (int c = 1; c <= totalCols; c++)
                        {
                            listaCeldas.Add((Excel.Range)area.Cells[f, c]);
                        }
                    }
                    ExcelHelper.LiberarCom(area);
                }

                int totalCeldas = listaCeldas.Count;
                for (int i = 0; i < totalCeldas; i++)
                {
                    Excel.Range celdaDisp = listaCeldas[i];
                    string dirDispAbs = celdaDisp.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                    List<string> dirsTarget = new List<string>();
                    for (int j = 0; j < totalCeldas; j++)
                    {
                        if (i != j) dirsTarget.Add(listaCeldas[j].get_Address(true, true, Excel.XlReferenceStyle.xlA1, false));
                    }

                    Excel.Range rangeObj = wsActual.get_Range(string.Join(",", dirsTarget), Type.Missing);
                    Excel.FormatConditions conds = rangeObj.FormatConditions;

                    string formulaFCUniversal = $"={dirDispAbs}=\"X\"";
                    string formulaFCLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaFCUniversal, celdaDisp.Row);

                    Excel.FormatCondition fc = (Excel.FormatCondition)conds.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaFCLocal);
                    fc.Interior.Pattern = Excel.XlPattern.xlPatternGray50;
                    fc.Interior.PatternColorIndex = (int)Excel.XlColorIndex.xlColorIndexAutomatic;
                    fc.StopIfTrue = false;

                    ExcelHelper.LiberarCom(fc);
                    ExcelHelper.LiberarCom(conds);
                    ExcelHelper.LiberarCom(rangeObj);
                }

                for (int k = 0; k < listaCeldas.Count; k++) ExcelHelper.LiberarCom(listaCeldas[k]);
                AuditoriaCenso.RegistrarAccion(libroCenso, idPregunta, "Bloqueo Código X - Cruzado ", rangoCapturado.Address.Replace("$", ""), "Formato Condicional");

                direccionSuma = blk_celdaAuxiliar.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                return new ResultadoValidacion { Exito = true, Mensaje = "Bloqueo cruzado automático aplicado exitosamente.", AlertaInyectada = true };
            }
            finally
            {
                ExcelHelper.LiberarCom(celdaTituloAux);
                ExcelHelper.LiberarCom(blk_celdaAuxiliar);
                ExcelHelper.LiberarCom(rangeInputAux);
            }
        }

        // ----------------------------------------------------------------------------------
        // CASO 1.2: BLOQUEO POR DISPARADOR (INTELIGENTE: SIMPLE O MATRICIAL)
        // ----------------------------------------------------------------------------------
        private ResultadoValidacion EjecutarCodigoX_DisparadorInteligente(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Worksheet wsActual, Excel.Range rangoCapturado, out string direccionSuma)
        {
            direccionSuma = string.Empty;
            List<string> formulasAuxiliares = new List<string>();
            bool agregarOtro = false;
            int contadorBloqueos = 0;

            do
            {
                contadorBloqueos++;
                excelApp.ScreenUpdating = true;

                object resDisp = excelApp.InputBox($"Selecciona la CELDA (o COLUMNA) DISPARADORA:", "SAVCNG- RANGO DISPARADOR", Type: 8);
                if (resDisp is bool && (bool)resDisp == false) break;
                Excel.Range rangoDisp = (Excel.Range)resDisp;

                object resBloq = excelApp.InputBox($"[Bloqueo Disparador #{contadorBloqueos}]\nSelecciona la MATRIZ o CELDAS OBJETIVO a bloquear:", "Celdas Objetivo", Type: 8);
                if (resBloq is bool && (bool)resBloq == false)
                {
                    ExcelHelper.LiberarCom(rangoDisp);
                    break;
                }
                Excel.Range rangoObjetivo = (Excel.Range)resBloq;

                // Auditoría de Coexistencia sobre el objetivo
                if (!VerificarYLimpiarFormatosPrevios(rangoObjetivo))
                {
                    ExcelHelper.LiberarCom(rangoObjetivo);
                    ExcelHelper.LiberarCom(rangoDisp);
                    return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada desde la Auditoría de Coexistencia.", AlertaInyectada = false };
                }

                try
                {
                    excelApp.ScreenUpdating = false;
                    int totalFilasDisp = rangoDisp.Rows.Count;
                    int totalFilasObj = rangoObjetivo.Rows.Count;

                    // ==========================================
                    // MODO 1: SIMPLE (1 celda disparadora)
                    // ==========================================
                    if (totalFilasDisp == 1)
                    {
                        Excel.Range celdaDisp = (Excel.Range)rangoDisp.Cells[1, 1];
                        string dirDispAbs = celdaDisp.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                        string dirBloqRel = rangoObjetivo.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                        string formulaFCUniversal = $"={dirDispAbs}=\"X\"";
                        string formulaLocalFC = ExcelHelper.TraducirFormulaLocal(wsActual, formulaFCUniversal, celdaDisp.Row);

                        Excel.FormatConditions conds = rangoObjetivo.FormatConditions;
                        Excel.FormatCondition fc = (Excel.FormatCondition)conds.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaLocalFC);
                        fc.Interior.Pattern = Excel.XlPattern.xlPatternGray50;
                        fc.Interior.PatternColorIndex = (int)Excel.XlColorIndex.xlColorIndexAutomatic;
                        fc.StopIfTrue = false;

                        formulasAuxiliares.Add($"=IF(AND({dirDispAbs}=\"X\", COUNTA({dirBloqRel})>0), 1, 0)");

                        ExcelHelper.LiberarCom(fc);
                        ExcelHelper.LiberarCom(conds);
                        ExcelHelper.LiberarCom(celdaDisp);
                    }
                    // ==========================================
                    // MODO 2: MATRICIAL (Columna fila por fila)
                    // ==========================================
                    else
                    {
                        if (totalFilasDisp != totalFilasObj)
                        {
                            MessageBox.Show($"Discrepancia detectada:\nEl disparador tiene {totalFilasDisp} filas y el objetivo {totalFilasObj}.\n\nPara aplicar el modo matricial, ambas tablas deben tener la misma cantidad de filas.", "Error de Selección", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            agregarOtro = true; // Forzamos para que vuelva a intentar
                            continue;
                        }

                        Excel.Range primeraCeldaDisp = (Excel.Range)rangoDisp.Cells[1, 1];
                        string colLetterDisp = Regex.Replace(primeraCeldaDisp.get_Address(false, false), @"[\d]", "");
                        string dirDispFC = $"${colLetterDisp}{primeraCeldaDisp.Row}";

                        string formulaFCUniversal = $"={dirDispFC}=\"X\"";
                        string formulaFCLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaFCUniversal, primeraCeldaDisp.Row);

                        Excel.FormatConditions conds = rangoObjetivo.FormatConditions;
                        Excel.FormatCondition fc = (Excel.FormatCondition)conds.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaFCLocal);
                        fc.Interior.Pattern = Excel.XlPattern.xlPatternGray50;
                        fc.Interior.PatternColorIndex = (int)Excel.XlColorIndex.xlColorIndexAutomatic;
                        fc.StopIfTrue = false;

                        for (int i = 1; i <= totalFilasObj; i++)
                        {
                            Excel.Range celdaDispFila = (Excel.Range)rangoDisp.Rows[i];
                            Excel.Range celdaObjFila = (Excel.Range)rangoObjetivo.Rows[i];
                            Excel.Range celdaDispUnica = (Excel.Range)celdaDispFila.Cells[1, 1];

                            string dirDispFilaAbsCol = $"${colLetterDisp}{celdaDispUnica.Row}";
                            string dirObjFilaRel = celdaObjFila.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                            formulasAuxiliares.Add($"=IF(AND({dirDispFilaAbsCol}=\"X\", COUNTA({dirObjFilaRel})>0), 1, 0)");

                            ExcelHelper.LiberarCom(celdaDispUnica);
                            ExcelHelper.LiberarCom(celdaObjFila);
                            ExcelHelper.LiberarCom(celdaDispFila);
                        }

                        ExcelHelper.LiberarCom(fc);
                        ExcelHelper.LiberarCom(conds);
                        ExcelHelper.LiberarCom(primeraCeldaDisp);
                    }

                    excelApp.ScreenUpdating = true;
                    DialogResult resOtro = MessageBox.Show(
                        $"Regla configurada exitosamente.\n\n¿Deseas agregar OTRA celda/columna disparadora con código X para esta MISMA validación?",
                        "Agregar Otra Regla", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    agregarOtro = (resOtro == DialogResult.Yes);
                }
                finally
                {
                    ExcelHelper.LiberarCom(rangoObjetivo);
                    ExcelHelper.LiberarCom(rangoDisp);
                }

            } while (agregarOtro);

            if (formulasAuxiliares.Count == 0)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "No se configuró ningún bloqueo interno.", AlertaInyectada = false };
            }

            excelApp.ScreenUpdating = true;

            object resAux = excelApp.InputBox("Selecciona la CELDA AUXILIAR donde irán las formulas :", "Celda Auxiliar Base", Type: 8);
            if (resAux is bool && (bool)resAux == false)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Se canceló la creación de las formulas auxiliares.", AlertaInyectada = false };
            }
            Excel.Range rangeInputAux = (Excel.Range)resAux;
            Excel.Range blk_celdaAuxBase = (Excel.Range)rangeInputAux.Cells[1, 1];
            ExcelHelper.LiberarCom(rangeInputAux);

            Excel.Range celdaTituloAux = null;
            Excel.Range celdaAuxSuma = null;
            List<string> direccionesFormulas = new List<string>();

            try
            {
                excelApp.ScreenUpdating = false;

                celdaTituloAux = blk_celdaAuxBase.Offset[-1, 0];
                celdaTituloAux.Value = "SAV-BLOQUEOS COD X";
                celdaTituloAux.Font.Name = "Arial";
                celdaTituloAux.Font.Size = 4.5;
                celdaTituloAux.Font.Bold = true;
                celdaTituloAux.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                celdaTituloAux.WrapText = true; // Ajustar texto
                celdaTituloAux.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(166, 166, 166));

                for (int i = 0; i < formulasAuxiliares.Count; i++)
                {
                    Excel.Range celdaFormulaActual = blk_celdaAuxBase.Offset[i, 0];
                    celdaFormulaActual.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulasAuxiliares[i], celdaFormulaActual.Row);
                    direccionesFormulas.Add(celdaFormulaActual.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false));
                    ExcelHelper.LiberarCom(celdaFormulaActual);
                }

                string primeraCeldaFormula = direccionesFormulas[0];
                string ultimaCeldaFormula = direccionesFormulas[direccionesFormulas.Count - 1];
                string rangoSumaVertical = direccionesFormulas.Count == 1 ? primeraCeldaFormula : $"{primeraCeldaFormula}:{ultimaCeldaFormula}";

                celdaAuxSuma = blk_celdaAuxBase.Offset[formulasAuxiliares.Count, 0];
                string formulaSumaUniversal = $"=SUM({rangoSumaVertical})";
                celdaAuxSuma.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaSumaUniversal, celdaAuxSuma.Row);

                string idPregunta = ExcelHelper.ObtenerNumeroPregunta(rangoCapturado);
                AuditoriaCenso.RegistrarAccion(libroCenso, idPregunta, "Bloqueo Código X - Disparador", rangoCapturado.Address.Replace("$", ""), $"Formato Condicional");

                direccionSuma = celdaAuxSuma.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                return new ResultadoValidacion
                {
                    Exito = true,
                    Mensaje = $"¡Logrado! Se inyectaron {formulasAuxiliares.Count} bloqueos con sus formulas auxiliares.",
                    AlertaInyectada = true
                };
            }
            finally
            {
                ExcelHelper.LiberarCom(celdaAuxSuma);
                ExcelHelper.LiberarCom(celdaTituloAux);
                ExcelHelper.LiberarCom(blk_celdaAuxBase);
            }
        }

        // ----------------------------------------------------------------------------------
        // CASO 1.3: BLOQUEO ROJO Y GRIS (Pintado por FC + Alerta por Celda Auxiliar)
        // ----------------------------------------------------------------------------------
        private ResultadoValidacion EjecutarCodigoX_BloqueoRojoYGris(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Worksheet wsActual, Excel.Range rangoCapturado, out string direccionSuma)
        {
            direccionSuma = string.Empty;
            excelApp.ScreenUpdating = true;

            Excel.Range rangoDisparador = null;
            Excel.Range rangoObjetivo = null;
            Excel.Range rangeInputAux = null;
            Excel.Range blk_celdaAuxiliar = null;
            Excel.Range primeraCeldaDisp = null;
            Excel.Range primeraCeldaObj = null;
            Excel.Range celdaTituloAux = null;
            Excel.FormatConditions conds = null;
            Excel.FormatCondition fcGris = null;
            Excel.FormatCondition fcRojo = null;

            try
            {
                object resDisp = excelApp.InputBox("[Bloqueo Tabla a Tabla]\nSelecciona la MATRIZ DISPARADORA (Tabla 1 con las 'X'):", "Matriz Disparadora", Type: 8);
                if (resDisp is bool && (bool)resDisp == false) return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada por el usuario.", AlertaInyectada = false };
                rangoDisparador = (Excel.Range)resDisp;

                object resBloq = excelApp.InputBox("[Bloqueo Tabla a Tabla]\nSelecciona la MATRIZ OBJETIVO (Tabla 2 a bloquear/validar):", "Matriz Objetivo", Type: 8);
                if (resBloq is bool && (bool)resBloq == false) return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada por el usuario.", AlertaInyectada = false };
                rangoObjetivo = (Excel.Range)resBloq;

                if (!VerificarYLimpiarFormatosPrevios(rangoObjetivo))
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada desde la Auditoría de Coexistencia.", AlertaInyectada = false };
                }

                object resAux = excelApp.InputBox("Selecciona la CELDA AUXILIAR para las formulas auxiliares:", "Celda Auxiliar", Type: 8);
                if (resAux is bool && (bool)resAux == false) return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada por el usuario.", AlertaInyectada = false };

                rangeInputAux = (Excel.Range)resAux;
                blk_celdaAuxiliar = (Excel.Range)rangeInputAux.Cells[1, 1];

                excelApp.ScreenUpdating = false;

                primeraCeldaDisp = (Excel.Range)rangoDisparador.Cells[1, 1];
                primeraCeldaObj = (Excel.Range)rangoObjetivo.Cells[1, 1];

                string dirDispRelativa = primeraCeldaDisp.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                string dirObjRelativa = primeraCeldaObj.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                string dirDispAbs = rangoDisparador.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                string dirObjAbs = rangoObjetivo.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                celdaTituloAux = blk_celdaAuxiliar.Offset[-1, 0];
                celdaTituloAux.Value = "SAV-BLOQUEO COD X MATRICIAL";
                celdaTituloAux.Font.Name = "Arial";
                celdaTituloAux.Font.Size = 4.5;
                celdaTituloAux.Font.Bold = true;
                celdaTituloAux.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                celdaTituloAux.WrapText = true; // Ajustar texto
                celdaTituloAux.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(166, 166, 166));

                string formulaAuxUniversal = $"=SUMPRODUCT(({dirDispAbs}<>\"X\")*({dirObjAbs}<>\"\"))";

                blk_celdaAuxiliar.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaAuxUniversal, blk_celdaAuxiliar.Row);

                conds = rangoObjetivo.FormatConditions;

                string formulaGrisUniversal = $"={dirDispRelativa}<>\"X\"";
                string formulaGrisLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaGrisUniversal, primeraCeldaObj.Row);

                fcGris = (Excel.FormatCondition)conds.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaGrisLocal);
                fcGris.Interior.Pattern = Excel.XlPattern.xlPatternGray50;
                fcGris.Interior.PatternColorIndex = (int)Excel.XlColorIndex.xlColorIndexAutomatic;
                fcGris.StopIfTrue = false;

                string formulaRojoUniversal = $"=AND({dirDispRelativa}=\"X\", {dirObjRelativa}=\"\")";
                string formulaRojoLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaRojoUniversal, primeraCeldaObj.Row);

                fcRojo = (Excel.FormatCondition)conds.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaRojoLocal);
                fcRojo.Interior.Color = ColorTranslator.ToOle(Color.Red);
                fcRojo.Font.Color = ColorTranslator.ToOle(Color.White);
                fcRojo.Font.Bold = true;
                fcRojo.StopIfTrue = false;

                string idPregunta = ExcelHelper.ObtenerNumeroPregunta(rangoCapturado);
                AuditoriaCenso.RegistrarAccion(libroCenso, idPregunta, "Bloqueo Código X - Tabla a Tabla", rangoObjetivo.Address.Replace("$", ""), "Formato Condicional.");

                direccionSuma = blk_celdaAuxiliar.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                return new ResultadoValidacion
                {
                    Exito = true,
                    Mensaje = "Bloqueo Tabla a Tabla inyectado correctamente.",
                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion
                {
                    Exito = false,
                    Mensaje = "Error en Bloqueo Tabla a Tabla " + ex.Message,
                    AlertaInyectada = false
                };
            }
            finally
            {
                ExcelHelper.LiberarCom(fcRojo);
                ExcelHelper.LiberarCom(fcGris);
                ExcelHelper.LiberarCom(conds);
                ExcelHelper.LiberarCom(celdaTituloAux);
                ExcelHelper.LiberarCom(primeraCeldaObj);
                ExcelHelper.LiberarCom(primeraCeldaDisp);
                ExcelHelper.LiberarCom(blk_celdaAuxiliar);
                ExcelHelper.LiberarCom(rangeInputAux);
                ExcelHelper.LiberarCom(rangoObjetivo);
                ExcelHelper.LiberarCom(rangoDisparador);
            }
        }
        #endregion

        #region MÉTODOS Y CATEGORÍAS DIRECTAS
        // ----------------------------------------------------------------------------------
        // LÓGICA DIRECTA: BLOQUEO NUMERAL POR FILA (DINÁMICO CON LISTA DE CÓDIGOS Y OR)
        // ----------------------------------------------------------------------------------
        private ResultadoValidacion EjecutarNumeral_FilaPorFila(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Worksheet wsActual, Excel.Range rangoCapturado, out string direccionSuma)
        {
            direccionSuma = string.Empty;
            excelApp.ScreenUpdating = true;

            Excel.Range rangoDisparador = null;
            Excel.Range rangoObjetivo = null;
            Excel.Range rangeInputAux = null;
            Excel.Range blk_celdaAuxBase = null;
            Excel.Range primeraCeldaDisp = null;
            Excel.Range celdaTituloAux = null;
            Excel.Range celdaAuxSuma = null;

            try
            {
                object resDisp = excelApp.InputBox("Selecciona la COLUMNA O RANGO DISPARADOR con el código excluyente:", "Rango Disparador", Type: 8);
                if (resDisp is bool && (bool)resDisp == false) return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada.", AlertaInyectada = false };
                rangoDisparador = (Excel.Range)resDisp;

                object resCodigos = excelApp.InputBox("Ingresa los códigos que bloquean la pregunta separados por coma (ej. 2,9):", "Códigos Excluyentes", "2,9", Type: 2);
                if (resCodigos is bool && (bool)resCodigos == false) return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada.", AlertaInyectada = false };

                string inputCodigos = resCodigos.ToString().Trim();
                string[] arrayCodigos = inputCodigos.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                if (arrayCodigos.Length == 0)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Debe especificar al menos un código excluyente.", AlertaInyectada = false };
                }

                object resBloq = excelApp.InputBox("Selecciona la(s) FILA(S) ó MATRIZ(CES) OBJETIVO.\nSi son varias partes, mantén presionado CTRL:", "Matriz Objetivo", Type: 8);
                if (resBloq is bool && (bool)resBloq == false) return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada.", AlertaInyectada = false };
                rangoObjetivo = (Excel.Range)resBloq;

                if (!VerificarYLimpiarFormatosPrevios(rangoObjetivo))
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada desde la Auditoría de Coexistencia.", AlertaInyectada = false };
                }

                int totalFilasDisp = rangoDisparador.Rows.Count;
                int totalAreasObj = rangoObjetivo.Areas.Count;

                for (int a = 1; a <= totalAreasObj; a++)
                {
                    Excel.Range areaTemp = null;
                    try
                    {
                        areaTemp = (Excel.Range)rangoObjetivo.Areas[a];
                        if (areaTemp.Rows.Count != totalFilasDisp)
                        {
                            return new ResultadoValidacion { Exito = false, Mensaje = $"Discrepancia matricial: El disparador tiene {totalFilasDisp} filas, pero el Área {a} objetivo tiene {areaTemp.Rows.Count} filas.", AlertaInyectada = false };
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(areaTemp);
                    }
                }

                object resAux = excelApp.InputBox("Selecciona la CELDA AUXILIAR para las fórmulas auxiliares:", "Celda Auxiliar", Type: 8);
                if (resAux is bool && (bool)resAux == false) return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada.", AlertaInyectada = false };
                rangeInputAux = (Excel.Range)resAux;
                blk_celdaAuxBase = (Excel.Range)rangeInputAux.Cells[1, 1];

                excelApp.ScreenUpdating = false;

                primeraCeldaDisp = (Excel.Range)rangoDisparador.Cells[1, 1];
                string colLetterDisp = Regex.Replace(primeraCeldaDisp.get_Address(false, false), @"[\d]", "");
                string dirDispFC = $"${colLetterDisp}{primeraCeldaDisp.Row}";

                List<string> partesOr = new List<string>();
                for (int c = 0; c < arrayCodigos.Length; c++)
                {
                    string codigoLimpio = arrayCodigos[c].Trim();
                    if (double.TryParse(codigoLimpio, out double numVal))
                    {
                        partesOr.Add($"{dirDispFC}={numVal}");
                    }
                    else
                    {
                        partesOr.Add($"{dirDispFC}=\"{codigoLimpio}\"");
                    }
                }
                string cadenaOrUniversal = string.Join(",", partesOr);
                string formulaFCUniversal = $"=AND({dirDispFC}<>\"\", OR({cadenaOrUniversal}))";

                for (int a = 1; a <= totalAreasObj; a++)
                {
                    Excel.Range areaActual = null;
                    Excel.FormatConditions conds = null;
                    Excel.FormatCondition fcGris = null;
                    try
                    {
                        areaActual = (Excel.Range)rangoObjetivo.Areas[a];
                        conds = areaActual.FormatConditions;
                        string formulaFCLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaFCUniversal, primeraCeldaDisp.Row);

                        fcGris = (Excel.FormatCondition)conds.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaFCLocal);
                        fcGris.Interior.Pattern = Excel.XlPattern.xlPatternGray50;
                        fcGris.Interior.PatternColorIndex = (int)Excel.XlColorIndex.xlColorIndexAutomatic;
                        fcGris.StopIfTrue = false;
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(fcGris);
                        ExcelHelper.LiberarCom(conds);
                        ExcelHelper.LiberarCom(areaActual);
                    }
                }

                List<string> direccionesAuxiliares = new List<string>();

                for (int i = 1; i <= totalFilasDisp; i++)
                {
                    Excel.Range celdaDispFila = null;
                    Excel.Range celdaDispUnica = null;
                    Excel.Range celdaAuxFila = null;

                    try
                    {
                        celdaDispFila = (Excel.Range)rangoDisparador.Rows[i];
                        celdaDispUnica = (Excel.Range)celdaDispFila.Cells[1, 1];
                        string dirDispFilaAbs = $"${colLetterDisp}${celdaDispUnica.Row}";

                        List<string> partesOrFila = new List<string>();
                        for (int c = 0; c < arrayCodigos.Length; c++)
                        {
                            string codigoLimpio = arrayCodigos[c].Trim();
                            if (double.TryParse(codigoLimpio, out double numVal))
                            {
                                partesOrFila.Add($"{dirDispFilaAbs}={numVal}");
                            }
                            else
                            {
                                partesOrFila.Add($"{dirDispFilaAbs}=\"{codigoLimpio}\"");
                            }
                        }
                        string cadenaOrFilaUniversal = string.Join(",", partesOrFila);

                        List<string> rangosCounta = new List<string>();
                        for (int a = 1; a <= totalAreasObj; a++)
                        {
                            Excel.Range areaActual = null;
                            Excel.Range filaEnArea = null;
                            try
                            {
                                areaActual = (Excel.Range)rangoObjetivo.Areas[a];
                                filaEnArea = (Excel.Range)areaActual.Rows[i];
                                rangosCounta.Add(filaEnArea.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false));
                            }
                            finally
                            {
                                ExcelHelper.LiberarCom(filaEnArea);
                                ExcelHelper.LiberarCom(areaActual);
                            }
                        }

                        string argumentosCounta = string.Join(",", rangosCounta);
                        string formulaAuxFilaUniv = $"=IF(AND({dirDispFilaAbs}<>\"\", OR({cadenaOrFilaUniversal}), COUNTA({argumentosCounta})>0), 1, 0)";

                        celdaAuxFila = blk_celdaAuxBase.Offset[i - 1, 0];
                        celdaAuxFila.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaAuxFilaUniv, celdaAuxFila.Row);
                        direccionesAuxiliares.Add(celdaAuxFila.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false));
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(celdaAuxFila);
                        ExcelHelper.LiberarCom(celdaDispUnica);
                        ExcelHelper.LiberarCom(celdaDispFila);
                    }
                }

                celdaTituloAux = blk_celdaAuxBase.Offset[-1, 0];
                celdaTituloAux.Value = "SAV-BLOQUEO POR NUMERAL";
                celdaTituloAux.Font.Name = "Arial";
                celdaTituloAux.Font.Size = 4.5;
                celdaTituloAux.Font.Bold = true;
                celdaTituloAux.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                celdaTituloAux.WrapText = true; // Ajustar texto
                celdaTituloAux.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(166, 166, 166));

                celdaAuxSuma = blk_celdaAuxBase.Offset[totalFilasDisp, 0];
                string primeraCeldaAuxDir = direccionesAuxiliares[0];
                string ultimaCeldaAuxDir = direccionesAuxiliares[direccionesAuxiliares.Count - 1];
                string rangoSumaVertical = totalFilasDisp == 1 ? primeraCeldaAuxDir : $"{primeraCeldaAuxDir}:{ultimaCeldaAuxDir}";

                string formulaSumaUniv = $"=SUM({rangoSumaVertical})";
                celdaAuxSuma.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaSumaUniv, celdaAuxSuma.Row);

                string idPregunta = ExcelHelper.ObtenerNumeroPregunta(rangoCapturado);
                AuditoriaCenso.RegistrarAccion(libroCenso, idPregunta, "Bloqueo por Numeral", rangoObjetivo.Address.Replace("$", ""), $"Formato Condicional");

                direccionSuma = celdaAuxSuma.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                return new ResultadoValidacion
                {
                    Exito = true,
                    Mensaje = $"Bloqueo excluyente aplicado exitosamente evaluando {arrayCodigos.Length} código(s).",
                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion
                {
                    Exito = false,
                    Mensaje = "Error en Bloqueo Excluyente: " + ex.Message,
                    AlertaInyectada = false
                };
            }
            finally
            {
                ExcelHelper.LiberarCom(celdaAuxSuma);
                ExcelHelper.LiberarCom(celdaTituloAux);
                ExcelHelper.LiberarCom(primeraCeldaDisp);
                ExcelHelper.LiberarCom(blk_celdaAuxBase);
                ExcelHelper.LiberarCom(rangeInputAux);
                ExcelHelper.LiberarCom(rangoObjetivo);
                ExcelHelper.LiberarCom(rangoDisparador);
            }
        }
        #endregion
    }
}