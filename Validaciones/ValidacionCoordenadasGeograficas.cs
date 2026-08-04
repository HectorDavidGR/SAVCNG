using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core;

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionCoordenadasGeograficas : IValidacionExcel
    {
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                // =========================================================================
                // 1. FASE 1: UX DE DECISIÓN (Bucle de captura segura con InputBox Nativo)
                // =========================================================================
                string opcionEscogida = "";
                while (true)
                {
                    // Type = 1 obliga a Excel a aceptar solo valores numéricos
                    object seleccionTipo = excelApp.InputBox(
                        "Ingresa el NÚMERO de la regla que deseas aplicar:\n\n" +
                        "1 = Latitud (11 a 33, exactamente 8 dígitos numéricos).\n" +
                        "2 = Longitud (-123 a -83, exactamente 8 dígitos numéricos).",
                        "SAVCNG - Tipo de Coordenada Geográfica",
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1);

                    // Si el usuario presiona "Cancelar", excelApp.InputBox devuelve el booleano 'false'
                    if (seleccionTipo is bool && (bool)seleccionTipo == false)
                    {
                        return new ResultadoValidacion { Exito = false, Mensaje = "Validación abortada por el usuario.", AlertaInyectada = false };
                    }

                    // Parseo de la opción
                    string valorIngresado = seleccionTipo.ToString();
                    if (valorIngresado == "1" || valorIngresado == "2")
                    {
                        opcionEscogida = valorIngresado;
                        break; // Salimos del bucle si la opción es válida
                    }
                    else
                    {
                        MessageBox.Show("Opción inválida. Por favor, ingresa 1 o 2.", "SAVCNG - Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                bool esLatitud = (opcionEscogida == "1");
                string formulaDV_Ingles = "";
                string formulaFC_Rango_Ingles = "";
                string formulaFC_Formato_Ingles = "";
                string mensajeError = "";

                // =========================================================================
                // 2. Recolección de Estado (Para Auditoría)
                // =========================================================================
                string numeroPregunta = ExcelHelper.ObtenerNumeroPregunta(rangoCapturado);
                string direccionRango = rangoCapturado.Address[false, false];

                excelApp.ScreenUpdating = false;

                // =========================================================================
                // 4. Inyección Matricial (Zero Leaks: Ciclos FOR estrictos)
                // =========================================================================
                int totalAreas = rangoCapturado.Areas.Count;
                for (int i = 1; i <= totalAreas; i++)
                {
                    Excel.Range area = null;
                    Excel.Range primeraCelda = null;
                    Excel.Validation validacion = null;
                    Excel.FormatConditions fcs = null;

                    // Punteros para la Regla 1 (Rango - Rojo)
                    Excel.FormatCondition fcRango = null;
                    Excel.Interior interiorRango = null;
                    Excel.Font fontRango = null;

                    // Punteros para la Regla 2 (Formato/Caracteres - Rojo)
                    Excel.FormatCondition fcFormato = null;
                    Excel.Interior interiorFormato = null;
                    Excel.Font fontFormato = null;

                    try
                    {
                        area = (Excel.Range)rangoCapturado.Areas[i];
                        primeraCelda = (Excel.Range)area.Cells[1, 1];

                        string celdaRef = primeraCelda.Address[false, false];

                        // Ingeniería de Fórmulas Cortocircuito Divididas
                        // Se usa SUBSTITUTE(SUBSTITUTE(..., ".", ""), "-", "") para contar solo los dígitos reales
                        string validadorDigitos = $"LEN(SUBSTITUTE(SUBSTITUTE(TRIM({celdaRef}), \".\", \"\"), \"-\", \"\"))";

                        if (esLatitud)
                        {
                            // Data Validation General (Engloba todo)
                            formulaDV_Ingles = $"=IF(ISBLANK({celdaRef}), TRUE, IF(ISERROR({celdaRef}+0), FALSE, AND(({celdaRef}+0)>=11, ({celdaRef}+0)<=33, {validadorDigitos}=8)))";

                            // FC Regla 1 (Rojo): Solo falla matemáticamente por rango
                            formulaFC_Rango_Ingles = $"=IF(ISBLANK({celdaRef}), FALSE, IF(ISERROR({celdaRef}+0), FALSE, OR(({celdaRef}+0)<11, ({celdaRef}+0)>33)))";

                            // FC Regla 2 (Rojo): Falla porque es texto, o rompe la longitud de 8 dígitos
                            formulaFC_Formato_Ingles = $"=IF(ISBLANK({celdaRef}), FALSE, IF(ISERROR({celdaRef}+0), TRUE, {validadorDigitos}<>8))";

                            mensajeError = "Latitud Inválida.\n- Debe estar entre 11 y 33.\n- Debe tener exactamente 8 dígitos (sin contar el punto decimal).";
                        }
                        else
                        {
                            formulaDV_Ingles = $"=IF(ISBLANK({celdaRef}), TRUE, IF(ISERROR({celdaRef}+0), FALSE, AND(({celdaRef}+0)>=-123, ({celdaRef}+0)<=-83, {validadorDigitos}=8)))";

                            formulaFC_Rango_Ingles = $"=IF(ISBLANK({celdaRef}), FALSE, IF(ISERROR({celdaRef}+0), FALSE, OR(({celdaRef}+0)<-123, ({celdaRef}+0)>-83)))";

                            formulaFC_Formato_Ingles = $"=IF(ISBLANK({celdaRef}), FALSE, IF(ISERROR({celdaRef}+0), TRUE, {validadorDigitos}<>8))";

                            mensajeError = "Longitud Inválida.\n- Debe estar entre -83 y -123.\n- Debe tener exactamente 8 dígitos (sin contar el punto decimal ni el signo negativo).";
                        }

                        // Traducción vía Fachada
                        string formulaDV_Trad = ExcelHelper.TraducirFormulaLocal(wsActual, formulaDV_Ingles, primeraCelda.Row);
                        string formulaFCRango_Trad = ExcelHelper.TraducirFormulaLocal(wsActual, formulaFC_Rango_Ingles, primeraCelda.Row);
                        string formulaFCFormato_Trad = ExcelHelper.TraducirFormulaLocal(wsActual, formulaFC_Formato_Ingles, primeraCelda.Row);

                        // Aplicar Data Validation
                        validacion = area.Validation;
                        validacion.Delete();
                        validacion.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Type.Missing, formulaDV_Trad, Type.Missing);
                        validacion.ErrorTitle = "Regla Geográfica Incumplida";
                        validacion.ErrorMessage = mensajeError;
                        validacion.ShowError = true;

                        // Aplicar Format Conditions Multi-Regla
                        fcs = area.FormatConditions;
                        fcs.Delete();

                        // -> Inyectar Regla 1: Rango Geográfico (Relleno rojo claro con texto rojo oscuro)
                        fcRango = (Excel.FormatCondition)fcs.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaFCRango_Trad, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        interiorRango = fcRango.Interior;
                        interiorRango.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 199, 206));
                        fontRango = fcRango.Font;
                        fontRango.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));

                        // -> Inyectar Regla 2: Longitud / Decimales / Texto (MISMO COLOR: Relleno rojo claro con texto rojo oscuro)
                        fcFormato = (Excel.FormatCondition)fcs.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaFCFormato_Trad, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        interiorFormato = fcFormato.Interior;
                        interiorFormato.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 199, 206));
                        fontFormato = fcFormato.Font;
                        fontFormato.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));
                    }
                    finally
                    {
                        // Destrucción Masiva y Quirúrgica (Zero Leaks absoluto)
                        ExcelHelper.LiberarCom(fontFormato);
                        ExcelHelper.LiberarCom(interiorFormato);
                        ExcelHelper.LiberarCom(fcFormato);
                        ExcelHelper.LiberarCom(fontRango);
                        ExcelHelper.LiberarCom(interiorRango);
                        ExcelHelper.LiberarCom(fcRango);
                        ExcelHelper.LiberarCom(fcs);
                        ExcelHelper.LiberarCom(validacion);
                        ExcelHelper.LiberarCom(primeraCelda);
                        ExcelHelper.LiberarCom(area);
                    }
                }

                // =========================================================================
                // 5. Auditoría
                // =========================================================================
                string tipoLog = esLatitud ? "Latitud" : "Longitud";
                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    numeroPregunta,
                    $"Validación de Coordenadas ({tipoLog})",
                    direccionRango,
                    "DV y Formatos Multicondicionales"
                );

                // =========================================================================
                // 6. Retorno de DTO
                // =========================================================================
                return new ResultadoValidacion { Exito = true, Mensaje = $"Validación de {tipoLog} aplicada con formato condicional rojo.", AlertaInyectada = true };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Error crítico en Coordenadas: " + ex.Message, AlertaInyectada = false };
            }
            finally
            {
                ExcelHelper.LiberarCom(wsActual);
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}