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
                // 1. FASE 1: UX DE DECISIÓN - TIPO DE REGLA (InputBox Nativo)
                // =========================================================================
                string opcionEscogida = "";
                while (true)
                {
                    object seleccionTipo = excelApp.InputBox(
                        "Ingresa el NÚMERO de la regla que deseas aplicar:\n\n" +
                        "1 = Latitud (11 a 33, hasta 8 dígitos numéricos).\n" +
                        "2 = Longitud (-123 a -83, hasta 9 dígitos numéricos).",
                        "SAVCNG - Tipo de Coordenada Geográfica",
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1);

                    if (seleccionTipo is bool && (bool)seleccionTipo == false)
                    {
                        return new ResultadoValidacion { Exito = false, Mensaje = "Validación abortada por el usuario.", AlertaInyectada = false };
                    }

                    string valorIngresado = seleccionTipo.ToString();
                    if (valorIngresado == "1" || valorIngresado == "2")
                    {
                        opcionEscogida = valorIngresado;
                        break;
                    }
                    else
                    {
                        MessageBox.Show("Opción inválida. Por favor, ingresa 1 o 2.", "SAVCNG - Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                // =========================================================================
                // 2. FASE 1.5: UX DE DECISIÓN - DETECCIÓN Y APILAMIENTO DE FORMATOS
                // =========================================================================
                bool tieneFormatosPrevios = false;
                int totalAreasCheck = rangoCapturado.Areas.Count;

                // Detección táctica de formatos sin fugas de memoria
                for (int i = 1; i <= totalAreasCheck; i++)
                {
                    Excel.Range areaCheck = null;
                    Excel.FormatConditions fcsCheck = null;
                    try
                    {
                        areaCheck = (Excel.Range)rangoCapturado.Areas[i];
                        fcsCheck = areaCheck.FormatConditions;
                        if (fcsCheck.Count > 0)
                        {
                            tieneFormatosPrevios = true;
                            break; // Si encontramos uno, no necesitamos seguir buscando
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(fcsCheck);
                        ExcelHelper.LiberarCom(areaCheck);
                    }
                }

                bool eliminarPrevios = true; // Valor por defecto si el rango está limpio

                if (tieneFormatosPrevios)
                {
                    while (true)
                    {
                        object seleccionFormatos = excelApp.InputBox(
                            "Se detectaron FORMATOS CONDICIONALES previos en el rango seleccionado.\n" +
                            "¿Qué deseas hacer con ellos?\n\n" +
                            "1 = ELIMINAR los formatos previos (Inyectar SOLO la nueva validación).\n" +
                            "2 = CONSERVAR los formatos previos (APILAR la nueva validación).",
                            "SAVCNG - Gestión de Validaciones Previas",
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1);

                        if (seleccionFormatos is bool && (bool)seleccionFormatos == false)
                        {
                            return new ResultadoValidacion { Exito = false, Mensaje = "Validación abortada por el usuario.", AlertaInyectada = false };
                        }

                        string valorIngresadoFmt = seleccionFormatos.ToString();
                        if (valorIngresadoFmt == "1")
                        {
                            eliminarPrevios = true;
                            break;
                        }
                        else if (valorIngresadoFmt == "2")
                        {
                            eliminarPrevios = false;
                            break;
                        }
                        else
                        {
                            MessageBox.Show("Opción inválida. Por favor, ingresa 1 o 2.", "SAVCNG - Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }

                bool esLatitud = (opcionEscogida == "1");
                string formulaDV_Ingles = "";
                string formulaFC_Rango_Ingles = "";
                string formulaFC_Formato_Ingles = "";
                string mensajeError = "";

                // =========================================================================
                // 3. Recolección de Estado (Para Auditoría)
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

                    Excel.FormatCondition fcRango = null;
                    Excel.Interior interiorRango = null;
                    Excel.Font fontRango = null;

                    Excel.FormatCondition fcFormato = null;
                    Excel.Interior interiorFormato = null;
                    Excel.Font fontFormato = null;

                    try
                    {
                        area = (Excel.Range)rangoCapturado.Areas[i];
                        primeraCelda = (Excel.Range)area.Cells[1, 1];

                        string celdaRef = primeraCelda.Address[false, false];
                        string validadorDigitos = $"LEN(SUBSTITUTE(SUBSTITUTE(TRIM({celdaRef}), \".\", \"\"), \"-\", \"\"))";

                        if (esLatitud)
                        {
                            formulaDV_Ingles = $"=IF(ISBLANK({celdaRef}), TRUE, IF(ISERROR({celdaRef}+0), FALSE, AND(({celdaRef}+0)>=11, ({celdaRef}+0)<=33, {validadorDigitos}<=8)))";
                            formulaFC_Rango_Ingles = $"=IF(ISBLANK({celdaRef}), FALSE, IF(ISERROR({celdaRef}+0), FALSE, OR(({celdaRef}+0)<11, ({celdaRef}+0)>33)))";
                            formulaFC_Formato_Ingles = $"=IF(ISBLANK({celdaRef}), FALSE, IF(ISERROR({celdaRef}+0), TRUE, {validadorDigitos}>8))";
                            mensajeError = "Latitud Inválida.\n- Debe estar entre 11 y 33.\n- Debe tener hasta 8 dígitos (sin contar el punto decimal).";
                        }
                        else
                        {
                            formulaDV_Ingles = $"=IF(ISBLANK({celdaRef}), TRUE, IF(ISERROR({celdaRef}+0), FALSE, AND(({celdaRef}+0)>=-123, ({celdaRef}+0)<=-83, {validadorDigitos}<=9)))";
                            formulaFC_Rango_Ingles = $"=IF(ISBLANK({celdaRef}), FALSE, IF(ISERROR({celdaRef}+0), FALSE, OR(({celdaRef}+0)<-123, ({celdaRef}+0)>-83)))";
                            formulaFC_Formato_Ingles = $"=IF(ISBLANK({celdaRef}), FALSE, IF(ISERROR({celdaRef}+0), TRUE, {validadorDigitos}>9))";
                            mensajeError = "Longitud Inválida.\n- Debe estar entre -83 y -123.\n- Debe tener hasta 9 dígitos (sin contar el punto decimal ni el signo negativo).";
                        }

                        // Traducción vía Fachada
                        string formulaDV_Trad = ExcelHelper.TraducirFormulaLocal(wsActual, formulaDV_Ingles, primeraCelda.Row);
                        string formulaFCRango_Trad = ExcelHelper.TraducirFormulaLocal(wsActual, formulaFC_Rango_Ingles, primeraCelda.Row);
                        string formulaFCFormato_Trad = ExcelHelper.TraducirFormulaLocal(wsActual, formulaFC_Formato_Ingles, primeraCelda.Row);

                        // -----------------------------------------------------------------
                        // DATA VALIDATION (Obligatorio borrar porque Excel no permite apilar)
                        // -----------------------------------------------------------------
                        validacion = area.Validation;
                        validacion.Delete();
                        validacion.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Type.Missing, formulaDV_Trad, Type.Missing);
                        validacion.ErrorTitle = "Regla Geográfica Incumplida";
                        validacion.ErrorMessage = mensajeError;
                        validacion.ShowError = true;

                        // -----------------------------------------------------------------
                        // FORMAT CONDITIONS (Soporta apilamiento condicionado)
                        // -----------------------------------------------------------------
                        fcs = area.FormatConditions;
                        if (eliminarPrevios)
                        {
                            fcs.Delete(); // Solo borramos si el usuario escogió "1" o el rango estaba limpio
                        }

                        // -> Inyectar Regla 1 (Rango Geográfico)
                        fcRango = (Excel.FormatCondition)fcs.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaFCRango_Trad, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        interiorRango = fcRango.Interior;
                        interiorRango.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 199, 206));
                        fontRango = fcRango.Font;
                        fontRango.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));

                        // -> Inyectar Regla 2 (Longitud / Decimales)
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
                return new ResultadoValidacion { Exito = true, Mensaje = $"Validación de {tipoLog} aplicada correctamente.", AlertaInyectada = true };
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