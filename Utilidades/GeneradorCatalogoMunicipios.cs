using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Consumo obligatorio de la Fachada y el DTO[cite: 3, 4]

namespace SAVCNG_ExcelDNA.Utilidades
{
    public class GeneradorCatalogoMunicipios
    {
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, string rutaArchivoCatalogo)
        {
            Excel.Workbook libroOrigen = null;
            Excel.Worksheet wsOrigen = null;
            Excel.Range rangoUsado = null;

            Excel.Worksheet wsNuevo = null;
            Excel.Range celdaInicio = null;
            Excel.Range celdaFin = null;
            Excel.Range rangoDestino = null;
            Excel.Range columnasDestino = null;
            Excel.Worksheet wsActiva = null;

            Excel.Range celdaInicioUniq = null;
            Excel.Range celdaFinUniq = null;
            Excel.Range rangoUniq = null;

            // Variables de UI
            Excel.Range celdaEntidad = null;
            Excel.Range celdaClaveEntidad = null;
            Excel.Range celdaMunicipio = null;
            Excel.Range celdaClaveMunicipio = null;

            try
            {
                // =====================================================================
                // FASE 1: UX - RECOPILACIÓN DE CELDAS DE TRABAJO (HACER ESTO PRIMERO)
                // =====================================================================
                wsActiva = (Excel.Worksheet)libroCenso.Worksheets[2];
                wsActiva.Activate();

                object resEntidad = excelApp.InputBox("1. Selecciona la CELDA donde irá la LISTA DESPLEGABLE DE ENTIDAD:", "SAVCNG - Entidad", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resEntidad is bool && (bool)resEntidad == false) return new ResultadoValidacion { Exito = false, Mensaje = "Configuración cancelada.", AlertaInyectada = false };
                celdaEntidad = (Excel.Range)resEntidad;

                object resCveEnt = excelApp.InputBox("2. Selecciona la CELDA donde aparecerá la CLAVE DE LA ENTIDAD:", "SAVCNG - Clave Entidad", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resCveEnt is bool && (bool)resCveEnt == false) return new ResultadoValidacion { Exito = false, Mensaje = "Configuración cancelada.", AlertaInyectada = false };
                celdaClaveEntidad = (Excel.Range)resCveEnt;

                bool requiereMunicipio = false;
                bool incluirOtroMunicipio = false;
                bool incluirNoIdentificado = false;

                DialogResult respMun = MessageBox.Show("¿Deseas agregar también la lista desplegable DEPENDIENTE para el MUNICIPIO?", "SAVCNG - Municipio", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (respMun == DialogResult.Yes)
                {
                    requiereMunicipio = true;

                    object resMun = excelApp.InputBox("3. Selecciona la CELDA donde irá la LISTA DESPLEGABLE DE MUNICIPIO:", "SAVCNG - Municipio", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                    if (resMun is bool && (bool)resMun == false) return new ResultadoValidacion { Exito = false, Mensaje = "Configuración cancelada.", AlertaInyectada = false };
                    celdaMunicipio = (Excel.Range)resMun;

                    object resCveMun = excelApp.InputBox("4. Selecciona la CELDA donde aparecerá la CLAVE DEL MUNICIPIO:", "SAVCNG - Clave Municipio", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                    if (resCveMun is bool && (bool)resCveMun == false) return new ResultadoValidacion { Exito = false, Mensaje = "Configuración cancelada.", AlertaInyectada = false };
                    celdaClaveMunicipio = (Excel.Range)resCveMun;

                    DialogResult respOtro = MessageBox.Show("¿Deseas incluir la opción 'Otro municipio o demarcación territorial' en el catálogo?", "Opciones Adicionales", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    incluirOtroMunicipio = (respOtro == DialogResult.Yes);

                    DialogResult respNoId = MessageBox.Show("¿Deseas incluir la opción 'No identificado' en el catálogo?", "Opciones Adicionales", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    incluirNoIdentificado = (respNoId == DialogResult.Yes);
                }

                // =====================================================================
                // FASE 2: LECTURA AISLADA DEL CATÁLOGO EXTERNO
                // =====================================================================
                excelApp.ScreenUpdating = false;

                libroOrigen = excelApp.Workbooks.Open(rutaArchivoCatalogo);
                wsOrigen = (Excel.Worksheet)libroOrigen.Worksheets[1];
                rangoUsado = wsOrigen.UsedRange;

                object[,] matrizDatos = (object[,])rangoUsado.Value2;
                libroOrigen.Close(false);

                if (matrizDatos == null || matrizDatos.GetLength(0) < 5)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "El archivo seleccionado no cumple con la estructura de INEGI.", AlertaInyectada = false };
                }

                // =====================================================================
                // FASE 3: PROCESAMIENTO MÁQUINA DE ESTADOS (Corte de Control Inteligente)
                // =====================================================================
                int totalFilas = matrizDatos.GetLength(0);
                List<object[]> datosProcesados = new List<object[]>();

                List<string> entidadesUnicasNombres = new List<string>();
                List<string> entidadesUnicasClaves = new List<string>();

                datosProcesados.Add(new object[] { "CVEGEO", "CVE_ENT", "NOM_ENT", "CVE_MUN", "NOM_MUN", "HELPER_INDEX" });

                string cveEntAnterior = "";
                string cveEntOriAnterior = "";
                string nomEntAnterior = "";
                int conteoMunicipiosEntidad = 0; // NUEVO: Rastreador de volumen de municipios

                for (int i = 5; i <= totalFilas; i++)
                {
                    string cveGeo = matrizDatos[i, 1]?.ToString().Trim();
                    if (string.IsNullOrEmpty(cveGeo)) continue;

                    string cveEntOri = matrizDatos[i, 2]?.ToString().Trim();
                    string nomEnt = matrizDatos[i, 3]?.ToString().Trim();
                    string nomMun = matrizDatos[i, 6]?.ToString().Trim();

                    string cveEnt = string.IsNullOrEmpty(cveEntOri) ? "" : "2" + cveEntOri;
                    string cveMun = cveGeo;

                    if (!entidadesUnicasClaves.Contains(cveEnt))
                    {
                        entidadesUnicasClaves.Add(cveEnt);
                        entidadesUnicasNombres.Add(nomEnt);
                    }

                    // Corte de estado: Inyectar opciones dinámicas al cambiar de entidad
                    if (cveEnt != cveEntAnterior && !string.IsNullOrEmpty(cveEntAnterior))
                    {
                        // LÓGICA DINÁMICA: Si tiene 100 o más, usa la serie 990, si no, usa la serie 090
                        string sufijoOtro = conteoMunicipiosEntidad >= 100 ? "998" : "098";
                        string sufijoNoId = conteoMunicipiosEntidad >= 100 ? "999" : "099";

                        if (incluirOtroMunicipio)
                        {
                            string cveGeoOtro = $"{cveEntOriAnterior}{sufijoOtro}";
                            datosProcesados.Add(new object[] { cveGeoOtro, cveEntAnterior, nomEntAnterior, cveGeoOtro, "Otro municipio o demarcación territorial", $"{nomEntAnterior}|Otro municipio o demarcación territorial" });
                        }

                        if (incluirNoIdentificado)
                        {
                            string cveGeoNoId = $"{cveEntOriAnterior}{sufijoNoId}";
                            datosProcesados.Add(new object[] { cveGeoNoId, cveEntAnterior, nomEntAnterior, cveGeoNoId, "No identificado", $"{nomEntAnterior}|No identificado" });
                        }

                        // Reiniciamos el contador para la nueva entidad que apenas va a procesarse
                        conteoMunicipiosEntidad = 0;
                    }

                    datosProcesados.Add(new object[] { cveGeo, cveEnt, nomEnt, cveMun, nomMun, $"{nomEnt}|{nomMun}" });

                    // Incrementamos el contador por la fila válida que acabamos de agregar
                    conteoMunicipiosEntidad++;

                    cveEntAnterior = cveEnt;
                    cveEntOriAnterior = cveEntOri;
                    nomEntAnterior = nomEnt;
                }

                // Cierre del ciclo: Inyectar opciones dinámicas a la ÚLTIMA entidad procesada
                if (!string.IsNullOrEmpty(cveEntAnterior))
                {
                    string sufijoOtro = conteoMunicipiosEntidad >= 100 ? "998" : "098";
                    string sufijoNoId = conteoMunicipiosEntidad >= 100 ? "999" : "099";

                    if (incluirOtroMunicipio)
                    {
                        string cveGeoOtro = $"{cveEntOriAnterior}{sufijoOtro}";
                        datosProcesados.Add(new object[] { cveGeoOtro, cveEntAnterior, nomEntAnterior, cveGeoOtro, "Otro municipio o demarcación territorial", $"{nomEntAnterior}|Otro municipio o demarcación territorial" });
                    }

                    if (incluirNoIdentificado)
                    {
                        string cveGeoNoId = $"{cveEntOriAnterior}{sufijoNoId}";
                        datosProcesados.Add(new object[] { cveGeoNoId, cveEntAnterior, nomEntAnterior, cveGeoNoId, "No identificado", $"{nomEntAnterior}|No identificado" });
                    }
                }

                // Convertir listas a matrices de inyección
                int filasSalida = datosProcesados.Count;
                object[,] matrizSalida = new object[filasSalida, 6];
                for (int f = 0; f < filasSalida; f++)
                {
                    for (int c = 0; c < 6; c++) matrizSalida[f, c] = datosProcesados[f][c];
                }

                int totalUnicas = entidadesUnicasNombres.Count;
                object[,] matrizUnicas = new object[totalUnicas + 1, 2];
                matrizUnicas[0, 0] = "UNIQUE_ENT"; matrizUnicas[0, 1] = "UNIQUE_CVE";
                for (int u = 0; u < totalUnicas; u++)
                {
                    matrizUnicas[u + 1, 0] = entidadesUnicasNombres[u];
                    matrizUnicas[u + 1, 1] = entidadesUnicasClaves[u];
                }

                // =====================================================================
                // FASE 4: INYECCIÓN EN HOJA OCULTA (CENSO)
                // =====================================================================
                string nombreHojaCatalogo = "catalogo_ent_mun";
                excelApp.DisplayAlerts = false;

                for (int h = 1; h <= libroCenso.Worksheets.Count; h++)
                {
                    Excel.Worksheet hojaEvaluar = null;
                    try
                    {
                        hojaEvaluar = (Excel.Worksheet)libroCenso.Worksheets[h];
                        if (hojaEvaluar.Name == nombreHojaCatalogo) { hojaEvaluar.Delete(); break; }
                    }
                    finally { ExcelHelper.LiberarCom(hojaEvaluar); }
                }
                excelApp.DisplayAlerts = true;

                wsNuevo = (Excel.Worksheet)libroCenso.Worksheets.Add(After: libroCenso.Worksheets[libroCenso.Worksheets.Count]);
                wsNuevo.Name = nombreHojaCatalogo;

                celdaInicio = (Excel.Range)wsNuevo.Cells[1, 1];
                celdaFin = (Excel.Range)wsNuevo.Cells[filasSalida, 6];
                rangoDestino = wsNuevo.Range[celdaInicio, celdaFin];
                rangoDestino.NumberFormat = "@";
                rangoDestino.Value2 = matrizSalida;

                celdaInicioUniq = (Excel.Range)wsNuevo.Cells[1, 8];
                celdaFinUniq = (Excel.Range)wsNuevo.Cells[totalUnicas + 1, 9];
                rangoUniq = wsNuevo.Range[celdaInicioUniq, celdaFinUniq];
                rangoUniq.NumberFormat = "@";
                rangoUniq.Value2 = matrizUnicas;

                columnasDestino = wsNuevo.UsedRange.EntireColumn;
                columnasDestino.AutoFit();
                wsNuevo.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                // =====================================================================
                // FASE 5: INYECCIÓN DE FÓRMULAS UNIVERSALES (BLINDAJE 0x800A03EC)
                // =====================================================================
                try { libroCenso.Names.Item("SAVCNG_ColEnt").Delete(); } catch { }
                try { libroCenso.Names.Item("SAVCNG_ColMun").Delete(); } catch { }
                try { libroCenso.Names.Item("SAVCNG_ListaUnicas").Delete(); } catch { }

                libroCenso.Names.Add("SAVCNG_ColEnt", $"='{nombreHojaCatalogo}'!$C:$C");
                libroCenso.Names.Add("SAVCNG_ColMun", $"='{nombreHojaCatalogo}'!$E:$E");
                libroCenso.Names.Add("SAVCNG_ListaUnicas", $"='{nombreHojaCatalogo}'!$H$2:$H${totalUnicas + 1}");

                Excel.Range tlEnt = null;
                Excel.Range maEnt = null;
                Excel.Range maCveEnt = null;
                string addrEnt = "";
                try
                {
                    tlEnt = (Excel.Range)celdaEntidad.Cells[1, 1];
                    addrEnt = tlEnt.get_Address(false, true, Excel.XlReferenceStyle.xlA1, false, Type.Missing);

                    maEnt = celdaEntidad.MergeArea;
                    maCveEnt = celdaClaveEntidad.MergeArea;

                    maEnt.Validation.Delete();
                    maEnt.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Type.Missing, "=SAVCNG_ListaUnicas", Type.Missing);
                    maEnt.Validation.InCellDropdown = true;

                    string formulaClaveEntIngles = $"=IF(ISBLANK({addrEnt}), \"\", VLOOKUP({addrEnt}, '{nombreHojaCatalogo}'!$H:$I, 2, FALSE))";
                    maCveEnt.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActiva, formulaClaveEntIngles, celdaClaveEntidad.Row);
                }
                finally
                {
                    ExcelHelper.LiberarCom(tlEnt);
                    ExcelHelper.LiberarCom(maEnt);
                    ExcelHelper.LiberarCom(maCveEnt);
                }

                if (requiereMunicipio)
                {
                    Excel.Range tlMun = null;
                    Excel.Range maMun = null;
                    Excel.Range maCveMun = null;
                    string addrMun = "";
                    try
                    {
                        tlMun = (Excel.Range)celdaMunicipio.Cells[1, 1];
                        addrMun = tlMun.get_Address(false, true, Excel.XlReferenceStyle.xlA1, false, Type.Missing);

                        maMun = celdaMunicipio.MergeArea;
                        maCveMun = celdaClaveMunicipio.MergeArea;

                        maMun.Validation.Delete();
                        string formulaValMunIngles = $"=OFFSET(SAVCNG_ColMun, MATCH(IF(ISBLANK({addrEnt}), \"NOM_ENT\", {addrEnt}), SAVCNG_ColEnt, 0)-1, 0, MAX(1, COUNTIF(SAVCNG_ColEnt, {addrEnt})), 1)";
                        string valMunLocal = ExcelHelper.TraducirFormulaLocal(wsActiva, formulaValMunIngles, celdaMunicipio.Row);

                        maMun.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Type.Missing, valMunLocal, Type.Missing);
                        maMun.Validation.InCellDropdown = true;

                        string formClaveMunIngles = $"=IF(OR(ISBLANK({addrEnt}), ISBLANK({addrMun})), \"\", INDEX('{nombreHojaCatalogo}'!$D:$D, MATCH({addrEnt}&\"|\"&{addrMun}, '{nombreHojaCatalogo}'!$F:$F, 0)))";
                        maCveMun.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActiva, formClaveMunIngles, celdaClaveMunicipio.Row);
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(tlMun);
                        ExcelHelper.LiberarCom(maMun);
                        ExcelHelper.LiberarCom(maCveMun);
                    }
                }

                wsActiva.Activate();

                // =====================================================================
                // FASE 6: AUDITORÍA Y DTO
                // =====================================================================
                AuditoriaCenso.RegistrarAccion(
                    libroCenso, "N/A", "Inyección de Catálogos Geográficos (XP Compatible)",
                    celdaEntidad.Address.Replace("$", ""), "Data Validation (OFFSET) + Helper Columns" //[cite: 2]
                );

                return new ResultadoValidacion //[cite: 4]
                {
                    Exito = true,
                    Mensaje = $"Catálogo inyectado y automatizado con éxito.\n\nSe blindó la compatibilidad con versiones anteriores de Excel y se automatizaron las claves.",
                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Error crítico durante la inyección del catálogo: " + ex.Message, AlertaInyectada = false };
            }
            finally
            {
                ExcelHelper.LiberarCom(celdaEntidad);
                ExcelHelper.LiberarCom(celdaClaveEntidad);
                ExcelHelper.LiberarCom(celdaMunicipio);
                ExcelHelper.LiberarCom(celdaClaveMunicipio);
                ExcelHelper.LiberarCom(rangoUniq);
                ExcelHelper.LiberarCom(celdaFinUniq);
                ExcelHelper.LiberarCom(celdaInicioUniq);
                ExcelHelper.LiberarCom(columnasDestino);
                ExcelHelper.LiberarCom(rangoDestino);
                ExcelHelper.LiberarCom(celdaFin);
                ExcelHelper.LiberarCom(celdaInicio);
                ExcelHelper.LiberarCom(wsNuevo);
                ExcelHelper.LiberarCom(rangoUsado);
                ExcelHelper.LiberarCom(wsActiva);
                ExcelHelper.LiberarCom(wsOrigen);
                ExcelHelper.LiberarCom(libroOrigen);

                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}