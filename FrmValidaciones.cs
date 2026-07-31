using SAVCNG_ExcelDNA.Core;
using SAVCNG_ExcelDNA.Validaciones;
using System;
using System.Drawing;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; // Importante para entenderse con Excel

namespace SAVCNG_ExcelDNA
{
    public partial class FrmValidaciones : Form
    {
        // Esta variable guardará el archivo de Excel que abrimos
        private Excel.Workbook _libroCenso;

        private Excel.Range _rangoCapturado;

        // Variable para guardar el rango de opciones del catálogo
        private Excel.Range _rangoCatalogo;

        // Aquí guardaremos los rangos que se van acumulando para validar
        private System.Collections.Generic.List<Excel.Range> _pendientesCatalogo = new System.Collections.Generic.List<Excel.Range>();

        // Aquí guardaremos el mapeo de qué rango pertenece a qué pregunta
        private System.Collections.Generic.Dictionary<string, string> _preguntaPorRango = new System.Collections.Generic.Dictionary<string, string>();

        // Carga inicial del formulario
        private void FrmValidaciones_Load(object sender, EventArgs e)
        {

            CargarEstadoDelCenso();

        }


        // Este es el constructor. Le agregamos "Excel.Workbook libroAbierto" para que reciba el censo
        public FrmValidaciones(Excel.Workbook libroAbierto)
        {
            InitializeComponent();

            // Guardamos el libro en nuestra variable para usarlo después
            _libroCenso = libroAbierto;

            // Actualizamos la etiqueta con el nombre del archivo
            lblCenso.Text = "Censo cargado: " + _libroCenso.Name;

            // Truco de experto: Hacemos que la ventana siempre esté por encima de Excel
            this.TopMost = true;
            // Que aparezca en el centro de la pantalla
            this.StartPosition = FormStartPosition.CenterScreen;

            // =========================================================================
            // INYECCIÓN DE UX: Llamamos al configurador de ToolTips al cargar la ventana
            // =========================================================================
            ConfigurarToolTips();

        }

        // --- INICIO DEL MÓDULO DE EXPERIENCIA DE USUARIO (UX) ---
        // Mensajes emergentes para los botones check
        private void ConfigurarToolTips()
        {
            // 1. Instanciamos el componente nativo de WinForms de manera global para la ventana
            ToolTip toolTipValidaciones = new ToolTip();

            // 2. Configuración de tiempos del motor (Valores en milisegundos)
            toolTipValidaciones.AutoPopDelay = 10000;  // UX: Tiempo que el mensaje se queda visible (10 segundos para leer tranquilos)
            toolTipValidaciones.InitialDelay = 200;    // UX: Tiempo de espera al poner el cursor encima antes de mostrar el mensaje (0.4 seg)
            toolTipValidaciones.ReshowDelay = 100;     // UX: Tiempo de transición al mover el cursor rápidamente a otro botón
            toolTipValidaciones.ShowAlways = true;     // Asegura que el tooltip se vea incluso si Excel tiene el foco principal por debajo

            // 3. Diseño Estético de la ventana flotante
            toolTipValidaciones.ToolTipIcon = ToolTipIcon.Info; // Muestra un pequeño icono azul de "i"
            toolTipValidaciones.ToolTipTitle = "Descripción de la validación";

            // 4. Mapeo del Diccionario de Descripciones por CheckBox
            // Nota: Reemplaza "chkDecimales", "chkFechas", etc. si el nombre interno de tus controles difiere un poco.

            toolTipValidaciones.SetToolTip(this.chkBlancos,
                "Inyecta una matriz de álgebra booleana que alerta en azul si el informante deja celdas vacías en una fila iniciada.");

            toolTipValidaciones.SetToolTip(this.chkBloqueo,
                "Bloquea celdas restrictivamente basándose en una condición lógica (Ej. 'Sí').\nPermite apilar reglas previas.");

            toolTipValidaciones.SetToolTip(this.chkCatalogos,
                "Crea una lista de opciones predefinidas. Es ideal para campos de opción múltiple o catálogos, limitando lo que se puede escribir en la celda.");
            
            toolTipValidaciones.SetToolTip(this.chkCoordenadas,
                "Restringe las celdas para aceptar únicamente coordenadas geográficas dentro del territorio nacional mexicano. Valida que la latitud se encuentre entre 11 y 33, o la longitud entre -83 y -123.");

            toolTipValidaciones.SetToolTip(this.chkDecimales,
                "Asegura que solo se puedan escribir números enteros (sin punto decimal). Además, permite registrar los códigos 'NS' y 'NA'");

            toolTipValidaciones.SetToolTip(this.chkEspClave,
                "Ideal para los campos de 'Especifique'. Si el texto que el usuario intenta capturar ya estaba en el catálogo principal, te avisa en color amarillo para no duplicar información.");

            toolTipValidaciones.SetToolTip(this.chkFechas,
                "Restringe la(s) celda(s) para aceptar solo números dentro de un rango que tú definas (ideal para días, meses o años). También acepta las claves 'NS' y 'NA'.");

            toolTipValidaciones.SetToolTip(this.chkFormatoTexto,
                "Ideal para preguntas abiertas como nombres o direcciones. Estandariza las respuestas transformando todo a mayúsculas y eliminando errores de datos como simbolos o dobles espacios.");

            toolTipValidaciones.SetToolTip(this.chkNS,
                "Busca el código 'NS' en el(los) rango(s) establecido(s). Si lo encuentra, te avisa con una alerta en color amarillo para que puedas revisarlas.");

            toolTipValidaciones.SetToolTip(this.chkSumas,
                "Motor de sumas cruzadas horizontales (FormatConditions) y generación de fórmulas automáticas (Σ) verticales.\nPermite apilamiento jerárquico multinivel.");
            
        }
        // Evento para seleccionar exclusivamente 1 checkbox a la vez
        private void CheckBox_Exclusivo_CheckedChanged(object sender, EventArgs e)
        {
            // 1. Identificamos qué CheckBox disparó el evento
            CheckBox chkActivo = sender as CheckBox;

            // 2. Si el CheckBox se está encendiendo, apagamos los demás
            if (chkActivo != null && chkActivo.Checked)
            {
                // NOTA: Cambia "panelOpciones" por el nombre del GroupBox, Panel o TableLayoutPanel 
                // donde tengas metidos tus 6 CheckBox.
                foreach (Control ctrl in tLP_Paso2_chkboxes.Controls)
                {
                    // Si el control es un CheckBox y NO es el que acaban de presionar...
                    if (ctrl is CheckBox && ctrl != chkActivo)
                    {
                        // Lo desmarcamos
                        ((CheckBox)ctrl).Checked = false;
                    }
                }
            }
        }

        // --- AQUÍ IRÁN LOS EVENTOS DE LOS BOTONES QUE SE USAN EN LA INTERFAZ ---
        // Evento para caprturar el rango seleccionado en el excel mediante boton Capturar Rango
        private void btnCapturarRango_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. Instanciamos la aplicación de Excel para interactuar con la interfaz activa
                Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

                // 2. Extraemos el objeto seleccionado actualmente por el usuario
                object seleccion = excelApp.Selection;

                // 3. Verificamos que se trate de un rango de celdas válido
                if (seleccion is Excel.Range)
                {
                    // =========================================================================
                    // PARCHE ZERO LEAKS: Liberamos el puntero anterior antes de sobreescribirlo
                    // =========================================================================
                    if (_rangoCapturado != null)
                    {
                        ExcelHelper.LiberarCom(_rangoCapturado);
                    }

                    // 4. Guardamos la nueva selección en la memoria global del Add-In
                    _rangoCapturado = (Excel.Range)seleccion;

                    // =========================================================================
                    // 5. ALGORITMO DE DETECCIÓN MULTI-PREGUNTA
                    // =========================================================================
                    System.Collections.Generic.List<string> preguntasDetectadas = new System.Collections.Generic.List<string>();

                    // Iteramos sobre cada sub-bloque (área) seleccionado con la tecla CTRL
                    foreach (Excel.Range area in _rangoCapturado.Areas)
                    {
                        try
                        {
                            // Reutilizamos el método de la Fachada enviando área por área
                            string pregunta = ExcelHelper.ObtenerNumeroPregunta(area);

                            // Evitamos duplicados en la interfaz
                            if (!preguntasDetectadas.Contains(pregunta))
                            {
                                preguntasDetectadas.Add(pregunta);
                            }
                        }
                        finally
                        {
                            // PARCHE ZERO LEAKS: Destrucción inmediata del área iterada
                            ExcelHelper.LiberarCom(area);
                        }
                    }

                    // Unimos todas las preguntas detectadas separadas por comas (Ej: "1.1, 1.2, 1.5")
                    string textoPreguntas = string.Join(", ", preguntasDetectadas);

                    // =========================================================================
                    // 6. ACTUALIZACIÓN DE LA EXPERIENCIA DE USUARIO (UX/UI)
                    // =========================================================================
                    string direccionLimpia = _rangoCapturado.Address.Replace("$", "");

                    lblPregunta.Text = "Pregunta(s) detectada(s): " + textoPreguntas;
                    lblRangoSeleccionado.Text = "Rango seleccionado: " + direccionLimpia;

                    MessageBox.Show(this,
                        $"Se capturó correctamente la selección.\n\n" +
                        $"• Coordenadas: {direccionLimpia}\n" +
                        $"• Bloques (Áreas): {_rangoCapturado.Areas.Count}\n" +
                        $"• Pregunta(s): {textoPreguntas}",
                        "SAVCNG - Captura Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this, "Por favor, selecciona celdas de Excel, no imágenes ni gráficos.",
                                    "Aviso de Captura", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Ocurrió un error al capturar en memoria: " + ex.Message,
                                "Error Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // Evento para aplicar validaciones mediante boton Aplicar Validación en pestaña de validaciones
        private void btnAplicar_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            string separador = excelApp.International[Excel.XlApplicationInternational.xlListSeparator].ToString();

            try
            {
                // Primero verificamos que el usuario no haya olvidado capturar un rango
                if (_rangoCapturado == null)
                {
                    MessageBox.Show(this,"¡Espera! Primero debes capturar un rango.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return; // Detenemos el código aquí
                }
                //=====================================================================================================
                // --- VALIDACIÓN DECIMALES ---
                else if (chkDecimales.Checked == true)
                {
                    // Invocamos la Estrategia aislada
                    IValidacionExcel validacion = new SAVCNG_ExcelDNA.Validaciones.ValidacionDecimales();
                    validacion.Ejecutar(excelApp, _libroCenso, _rangoCapturado);

                    chkDecimales.Checked = false; // Desmarcamos en UI tras terminar
                }
                // --- VALIDACIÓN Catalogos ---
                else if (chkCatalogos.Checked == true)
                {
                    IValidacionExcel validacionCatalogos = new ValidacionCatalogos();
                    validacionCatalogos.Ejecutar(excelApp, _libroCenso, _rangoCapturado);

                    chkCatalogos.Checked = false;
                }
                // --- VALIDACIÓN NS ---
                else if (chkNS.Checked == true)
                {
                    IValidacionExcel validacionNS = new ValidacionNS();
                    validacionNS.Ejecutar(excelApp, _libroCenso, _rangoCapturado);

                    chkNS.Checked = false;
                }
                // --- VALIDACIÓN FORMATO TEXTO ---
                else if (chkFormatoTexto.Checked == true)
                {
                    IValidacionExcel validacionFormatoTexto = new ValidacionFormatoTexto();
                    validacionFormatoTexto.Ejecutar(excelApp, _libroCenso, _rangoCapturado);

                    chkFormatoTexto.Checked = false;
                }
                // --- VALIDACIÓN BLOQUEOS ---
                else if (chkBloqueo.Checked == true)
                {
                    IValidacionExcel validacionBloqueos = new ValidacionBloqueos();
                    validacionBloqueos.Ejecutar(excelApp, _libroCenso, _rangoCapturado);

                    chkBloqueo.Checked = false;
                }
                // --- VALIDACION BLANCOS -------
                else if (chkBlancos.Checked == true)
                {
                    IValidacionExcel validacionBlancos = new ValidacionBlancos();
                    validacionBlancos.Ejecutar(excelApp, _libroCenso, _rangoCapturado);

                    chkBlancos.Checked = false; // Desmarcamos el UI tras finalizar o cancelar
                }
                // --- VALIDACIÓN ESPECIFIQUES (PALABRAS CLAVE) ---
                else if (chkEspClave.Checked == true)
                {
                    if (_libroCenso == null || _rangoCapturado == null)
                    {
                        MessageBox.Show(this, "Operación denegada: Captura la celda destino (Especifique) primero.",
                                        "Arquitectura SAVCNG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        chkEspClave.Checked = false;
                        return;
                    }

                    IValidacionExcel validacionEspecifique = new ValidacionEspecifique();
                    validacionEspecifique.Ejecutar(excelApp, _libroCenso, _rangoCapturado);

                    chkEspClave.Checked = false;
                }
                // --- VALIDACIÓN FECHAS ---
                else if (chkFechas.Checked == true)
                {
                    IValidacionExcel validacionFechas = new ValidacionFechas();
                    validacionFechas.Ejecutar(excelApp, _libroCenso, _rangoCapturado);

                    chkFechas.Checked = false;
                }
                // --- VALIDACIÓN SUMAS (NUEVO MOTOR BOOLEANO Y ALERTAS CON SOPORTE MERGE) ---
                else if (chkSumas.Checked == true)
                {
                    IValidacionExcel validacionSumas = new ValidacionSumas();
                    validacionSumas.Ejecutar(excelApp, _libroCenso, _rangoCapturado);

                    chkSumas.Checked = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,"Ocurrió un error al aplicar el formato: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void chkCatalogos_CheckedChanged(object sender, EventArgs e)
        {
            // Solo actuamos si el usuario MARCA la casilla
            if (chkCatalogos.Checked)
            {
                // Revisamos que no se haya saltado el paso 1 (Capturar rango)
                if (_libroCenso == null || _rangoCapturado == null)
                {
                    MessageBox.Show(this,"Primero carga un censo y captura un rango con el botón.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    chkCatalogos.Checked = false; // Desmarcamos la casilla
                }
                // ¡Y listo! No hacemos nada más aquí, dejamos que el botón Aplicar haga el trabajo duro.
            }
        }

        private void chkNS_CheckedChanged(object sender, EventArgs e)
        {
            // Solo actuamos si el usuario MARCA la casilla
            if (chkNS.Checked)
            {
                // Revisamos que no se haya saltado el paso 1 (Capturar rango)
                if (_libroCenso == null || _rangoCapturado == null)
                {
                    MessageBox.Show(this,"Primero carga un censo y captura un rango con el botón.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    chkNS.Checked = false; // Desmarcamos la casilla
                }
            }
        }

        private void chkBlancos_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBlancos.Checked)
            {
                // Validamos la regla de negocio: Nada ocurre si no se ha mapeado el terreno previamente
                if (_libroCenso == null || _rangoCapturado == null)
                {
                    MessageBox.Show(this,"Operación denegada: Carga un censo y define el rango de memoria primero.", "Advertencia Arquitectónica", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    chkBlancos.Checked = false;
                }
            }
        }

        private void chkFechas_CheckedChanged(object sender, EventArgs e)
        {
            // SI ==> MARCA la casilla
            if (chkFechas.Checked) // <-- Corregido: antes decía chkAños.Checked
            {
                // Revisamos que no se haya saltado el paso 1 (Capturar rango)
                if (_libroCenso == null || _rangoCapturado == null)
                {
                    MessageBox.Show(this,"Primero carga un censo y captura un rango con el botón.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    chkFechas.Checked = false; // <-- Corregido: antes decía chkAños.Checked
                }
            }
        }

        private void CargarEstadoDelCenso()
        {
            // Obtenemos el historial directamente de la memoria embebida del archivo Excel
            System.Data.DataTable historial = AuditoriaCenso.ObtenerHistorialCenso(_libroCenso);

            // Lo enlazamos a la grilla para que el usuario pueda ver, filtrar u ordenar
            dgvAuditoria.DataSource = historial;

            // 1.Responsividad del Contenedor(Se estira con la ventana)
            dgvAuditoria.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // 2. Responsividad del Contenido (Las columnas llenan el espacio vacío)
            dgvAuditoria.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // 3. Estética Institucional (Limpieza de ruido visual)
            dgvAuditoria.AllowUserToAddRows = false; // Quita la última fila en blanco editable
            dgvAuditoria.AllowUserToDeleteRows = false;
            dgvAuditoria.ReadOnly = true; // El auditor no debe alterar la bitácora desde aquí

            // 4. Experiencia de Navegación (Mejora la legibilidad)
            dgvAuditoria.SelectionMode = DataGridViewSelectionMode.FullRowSelect; // Selecciona toda la fila al hacer clic
            dgvAuditoria.MultiSelect = false;
            dgvAuditoria.RowHeadersVisible = false; // Oculta la columna gris inútil de la izquierda

            // 5. Estilo de celdas alternadas (Zebra striping para fatiga visual)
            dgvAuditoria.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
            dgvAuditoria.BackgroundColor = System.Drawing.Color.White; // Fondo blanco en lugar de gris oscuro

        }

        private void btnActualizar_Click(object sender, EventArgs e)
        {
            try
            {
                // UX: Cambiamos el cursor a "Cargando" (Reloj de arena/Círculo azul)
                Cursor.Current = Cursors.WaitCursor;

                // Invocamos nuestro motor de lectura que va a la hoja VeryHidden
                CargarEstadoDelCenso();

                // Refrescamos visualmente el control
                dgvAuditoria.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error al actualizar la vista de la bitácora: " + ex.Message, "Error de Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // UX: Devolvemos el cursor a la normalidad
                Cursor.Current = Cursors.Default;
            }
        }

        // Descargar Bitácora
        private void btnDescargarBitacora_Click(object sender, EventArgs e)
        {
            // 1. Validamos que el libro exista antes de pasarlo al servicio
            if (_libroCenso == null) return;

            // 2. Instanciamos el puente de comunicación COM con ExcelDNA
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

            // 3. Inyectamos la dependencia y ejecutamos el servicio aislado
            SAVCNG_ExcelDNA.Utilidades.IOperacionLibro servicioBitacora = new SAVCNG_ExcelDNA.Utilidades.DescargarBitacoraService();
            servicioBitacora.Ejecutar(excelApp, _libroCenso);
        }
        // --- INICIO DE EVENTOS PARA PESTAÑA REVISIÓN/UTILIDADES ---
        //Funcion para bloqueo de hojas con contraseña
        private void btnBloqueo_Click(object sender, EventArgs e)
        {
            // 1. Validamos que haya un censo (libro de Excel) cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria para proteger.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Contabilizamos el estado inicial de protección de las hojas
            int totalHojas = _libroCenso.Worksheets.Count;
            int hojasYaProtegidasAlInicio = 0;

            foreach (Excel.Worksheet hojaCheck in _libroCenso.Worksheets)
            {
                if (hojaCheck.ProtectContents)
                {
                    hojasYaProtegidasAlInicio++;
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(hojaCheck);
            }

            // EXCEPCIÓN A: Si TODAS las hojas ya contaban con protección previa
            if (hojasYaProtegidasAlInicio == totalHojas)
            {
                MessageBox.Show(this, "El libro ya está protegido por lo que no se aplicó protección con la clave del año elegido", "Aviso de Protección", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 3. Validamos y obtenemos el valor del año desde el NumericUpDown (nudPeriodo)
            int anioSeleccionado = decimal.ToInt32(nudPeriodo.Value);

            if (anioSeleccionado <= 0)
            {
                MessageBox.Show(this, "Por favor, selecciona un año válido en el periodo antes de proteger el libro.", "Periodo Requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 4. Estructuramos la contraseña concatenando "njm" con el año
            string passwordEstructurada = "njm" + anioSeleccionado.ToString();

            Excel.Application localExcelApp = null;
            int hojasProtegidasExitosamente = 0;

            try
            {
                localExcelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                localExcelApp.ScreenUpdating = false;

                // 5. Recorremos el libro para aplicar la protección únicamente a las hojas desprotegidas
                foreach (Excel.Worksheet hoja in _libroCenso.Worksheets)
                {
                    try
                    {
                        // Si la hoja no está protegida, aplicamos el bloqueo algorítmico
                        if (!hoja.ProtectContents)
                        {
                            hoja.Protect(
                                Password: passwordEstructurada,
                                DrawingObjects: true,
                                Contents: true,
                                Scenarios: true,
                                UserInterfaceOnly: false
                            );
                            hojasProtegidasExitosamente++;
                        }
                    }
                    catch (Exception exHoja)
                    {
                        System.Diagnostics.Debug.WriteLine($"No se pudo proteger la hoja '{hoja.Name}': {exHoja.Message}");
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(hoja);
                    }
                }

                localExcelApp.ScreenUpdating = true;

                // EXCEPCIÓN B: Si sólo ALGUNAS hojas estaban protegidas previamente
                if (hojasYaProtegidasAlInicio > 0)
                {
                    MessageBox.Show(this, $"Sólo se protegieron {hojasProtegidasExitosamente} hojas debido a que las demás ya estaban protegidas con una clave previa.", "Protección Parcial", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Flujo Normal: Si ninguna estaba protegida y se bloquearon todas
                    MessageBox.Show(this, $"Se han protegido con éxito todas las hojas del libro.\n\nContraseña aplicada: {passwordEstructurada}", "Libro Protegido", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                if (localExcelApp != null) localExcelApp.ScreenUpdating = true;
                MessageBox.Show(this, "Error crítico al intentar proteger las hojas: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Funcion para desbloqueo de hojas con contraseña
        private void btnDesbloqueo_Click(object sender, EventArgs e)
        {
            // 1. Validamos que haya un censo (libro de Excel) cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria para desproteger.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. VALIDACIÓN: Contabilizar cuántas hojas están protegidas actualmente antes de proceder
            int hojasProtegidasAlInicio = 0;
            foreach (Excel.Worksheet hojaCheck in _libroCenso.Worksheets)
            {
                if (hojaCheck.ProtectContents)
                {
                    hojasProtegidasAlInicio++;
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(hojaCheck);
            }

            // Si NINGUNA hoja está protegida en todo el libro, se lanza el Alert y se detiene la ejecución
            if (hojasProtegidasAlInicio == 0)
            {
                MessageBox.Show(this, "No se aplicó la acción de desproteger debido a que las hojas no estaban protegidas.", "Aviso de Desprotección", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 3. Obtenemos de forma segura el año desde el control 'nudPeriodo'
            int anioSeleccionado = decimal.ToInt32(nudPeriodo.Value);

            if (anioSeleccionado <= 0)
            {
                MessageBox.Show(this, "Por favor, verifica que el control de Periodo tenga el año correspondiente al censo.", "Periodo Inválido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 4. Reconstruimos la contraseña dinámica para desproteger (ej. "njm2026")
            string passwordEstructurada = "njm" + anioSeleccionado.ToString();

            Excel.Application localExcelApp = null;
            try
            {
                localExcelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                localExcelApp.ScreenUpdating = false;

                int hojasDesprotegidasExitosamente = 0;
                int hojasFallidas = 0;

                // 5. Recorremos cada una de las hojas de trabajo y desprotegemos únicamente las que estén bloqueadas
                foreach (Excel.Worksheet hoja in _libroCenso.Worksheets)
                {
                    try
                    {
                        if (hoja.ProtectContents)
                        {
                            hoja.Unprotect(passwordEstructurada);
                            hojasDesprotegidasExitosamente++;
                        }
                    }
                    catch (Exception exHoja)
                    {
                        // Si la contraseña dinámica no coincide con la clave real de la hoja
                        hojasFallidas++;
                        System.Diagnostics.Debug.WriteLine($"No se pudo desproteger la hoja '{hoja.Name}': {exHoja.Message}");
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(hoja);
                    }
                }

                localExcelApp.ScreenUpdating = true;

                // 6. Mensajes informativos de finalización
                if (hojasFallidas > 0)
                {
                    MessageBox.Show(this,
                        $"Se desprotegieron {hojasDesprotegidasExitosamente} hojas correctamente.\n\n" +
                        $"Sin embargo, {hojasFallidas} hojas no pudieron ser desprotegidas con la clave '{passwordEstructurada}'. " +
                        "Verifica si corresponden a otro año o clave previa.",
                        "Desprotección Parcial",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show(this, "Se han desprotegido con éxito todas las hojas del libro de trabajo.", "Libro Desprotegido", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                if (localExcelApp != null) localExcelApp.ScreenUpdating = true;
                MessageBox.Show(this, "Error crítico al intentar desproteger las hojas: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Funcion para aplicar un color a todas las celdas de todas las hojas para revisión de bloqueos
        private void btnAplicarFormato_Click(object sender, EventArgs e)
        {
            // 1. Validar que haya un censo cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

                // Apagamos la actualización de pantalla para que el barrido sea instantáneo
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();

                foreach (Excel.Worksheet wsCheck in _libroCenso.Worksheets)
                {
                    if (wsCheck.ProtectContents || wsCheck.ProtectDrawingObjects || wsCheck.ProtectScenarios)
                    {
                        hojasBloqueadas.Add("• " + wsCheck.Name);
                    }
                    // Liberación estricta de memoria COM
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCheck);
                }

                // Si hay hojas bloqueadas, abortamos y le avisamos al usuario amigablemente
                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(this,
                        "No se puede aplicar el formato de color porque el libro contiene hojas protegidas:\n\n" +
                        listaHojas + "\n\n" +
                        "Desprotege estas hojas (desde la pestaña Revisión) y vuelve a intentarlo.",
                        "Validación Interrumpida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // =========================================================================
                // FASE 2: INYECCIÓN DE FORMATO CONDICIONAL POR HOJA
                // =========================================================================
                int hojasModificadas = 0;

                foreach (Excel.Worksheet ws in _libroCenso.Worksheets)
                {
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;
                    Excel.Range celdaDummy = null;

                    try
                    {
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;
                        bool existe = false;

                        // 2.1 Búsqueda Inversa para verificar si la regla ya existe
                        for (int i = formatos.Count; i >= 1; i--)
                        {
                            Excel.FormatCondition fc = null;
                            try
                            {
                                object objFc = formatos[i];

                                // En C#, debemos verificar el tipo del objeto COM antes de leerlo
                                if (objFc is Excel.FormatCondition)
                                {
                                    fc = (Excel.FormatCondition)objFc;

                                    if (fc.Type == (int)Excel.XlFormatConditionType.xlExpression)
                                    {
                                        string formulaCondicion = "";
                                        try
                                        {
                                            formulaCondicion = fc.Formula1;
                                        }
                                        catch { /* Silenciamos si Excel deniega la lectura de esa fórmula específica */ }

                                        if (!string.IsNullOrEmpty(formulaCondicion))
                                        {
                                            // Limpiamos espacios y pasamos a mayúsculas para hacer una comparación universal
                                            string fLimpia = formulaCondicion.Replace(" ", "").ToUpper();

                                            // Cubrimos tanto "CELDA" (Español) como "CELL" (Inglés)
                                            if (fLimpia.Contains("CELDA(\"PROTECT\"") || fLimpia.Contains("CELL(\"PROTECT\""))
                                            {
                                                existe = true;
                                                break; // Ya existe, rompemos el ciclo
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // Ignoramos errores de lectura de reglas aisladas
                            }
                            finally
                            {
                                if (fc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fc);
                            }
                        }

                        // 2.2 Si no la encontró, la inyectamos
                        if (!existe)
                        {
                            // TRUCO ARQUITECTÓNICO: Traducir la fórmula al idioma de la PC local usando una celda oculta.
                            // Esto evita el clásico error de sintaxis al inyectar código rígido en un Excel en español.
                            celdaDummy = ws.Cells[1048576, 16384]; // Última celda XFD1048576
                            celdaDummy.Formula = "=CELL(\"protect\",A1)"; // La metemos en inglés universal
                            string formulaLocal = celdaDummy.FormulaLocal; // La extraemos en español (o el idioma del usuario)
                            celdaDummy.Clear();

                            Excel.FormatCondition nuevaRegla = (Excel.FormatCondition)formatos.Add(
                                Excel.XlFormatConditionType.xlExpression,
                                Type.Missing,
                                formulaLocal);

                            // Aplicamos el color RGB(255, 230, 153) que equivale al amarillo de tu macro original
                            nuevaRegla.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 230, 153));

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(nuevaRegla);
                            hojasModificadas++;
                        }
                    }
                    finally
                    {
                        // Limpieza absoluta del Garbage Collector en cada ciclo de la hoja
                        if (celdaDummy != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                        if (formatos != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(formatos);
                        if (celdasHoja != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdasHoja);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    }
                }

                // =========================================================================
                // FASE 3: CONCLUSIÓN UX
                // =========================================================================
                MessageBox.Show(this,
                    $"Formato de color (Celdas Bloqueadas) aplicado en todo el libro.\n\n" +
                    $"Se aplicó exitosamente en {hojasModificadas} hoja(s) que no contaban con él.",
                    "SAVCNG - Macro de Colores", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Ocurrió un error crítico al aplicar la macro de colores: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null)
                {
                    // Restauramos SIEMPRE la actualización visual al terminar o fallar
                    excelApp.ScreenUpdating = true;
                }
            }
        }
        //Funcion para quitar el formato aplicado por btnAplicarFormato_Click
        private void btnLimpiarFormato_Click(object sender, EventArgs e)
        {
            // 1. Validar que haya un censo cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

                // Apagamos la actualización de pantalla para un borrado instantáneo y silencioso
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();

                foreach (Excel.Worksheet wsCheck in _libroCenso.Worksheets)
                {
                    if (wsCheck.ProtectContents || wsCheck.ProtectDrawingObjects || wsCheck.ProtectScenarios)
                    {
                        hojasBloqueadas.Add("• " + wsCheck.Name);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCheck);
                }

                // Si la hoja está bloqueada, Excel no nos dejará borrar la regla
                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(this,
                        "No se puede limpiar el formato de color porque el libro contiene hojas protegidas:\n\n" +
                        listaHojas + "\n\n" +
                        "Desprotege estas hojas (desde la pestaña Revisión) y vuelve a intentarlo.",
                        "Operación Interrumpida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // =========================================================================
                // FASE 2: BÚSQUEDA Y ELIMINACIÓN DE LA REGLA
                // =========================================================================
                int hojasLimpiadas = 0;

                foreach (Excel.Worksheet ws in _libroCenso.Worksheets)
                {
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;
                    bool hojaModificada = false;

                    try
                    {
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;

                        // RECORRIDO INVERSO OBLIGATORIO: Del último formato al primero
                        for (int i = formatos.Count; i >= 1; i--)
                        {
                            Excel.FormatCondition fc = null;
                            try
                            {
                                object objFc = formatos[i];

                                if (objFc is Excel.FormatCondition)
                                {
                                    fc = (Excel.FormatCondition)objFc;

                                    if (fc.Type == (int)Excel.XlFormatConditionType.xlExpression)
                                    {
                                        string formulaCondicion = "";
                                        try { formulaCondicion = fc.Formula1; } catch { }

                                        if (!string.IsNullOrEmpty(formulaCondicion))
                                        {
                                            string fLimpia = formulaCondicion.Replace(" ", "").ToUpper();

                                            // Evaluamos si es la regla de nuestro botón "Aplicar"
                                            if (fLimpia.Contains("CELDA(\"PROTECT\"") || fLimpia.Contains("CELL(\"PROTECT\""))
                                            {
                                                // ¡La encontramos! La destruimos
                                                fc.Delete();
                                                hojaModificada = true;
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // Ignoramos errores aislados de lectura/borrado
                            }
                            finally
                            {
                                if (fc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fc);
                            }
                        }

                        if (hojaModificada)
                        {
                            hojasLimpiadas++;
                        }
                    }
                    finally
                    {
                        if (formatos != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(formatos);
                        if (celdasHoja != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdasHoja);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    }
                }

                // =========================================================================
                // FASE 3: CONCLUSIÓN UX
                // =========================================================================
                if (hojasLimpiadas > 0)
                {
                    MessageBox.Show(this,
                        $"El formato de color (Celdas Bloqueadas) fue eliminado correctamente.\n\n" +
                        $"Se limpiaron {hojasLimpiadas} hoja(s) del libro.",
                        "SAVCNG - Limpieza Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this,
                        "No se encontró el formato de color en ninguna hoja del libro.\n\nEl censo ya está limpio.",
                        "SAVCNG - Sin Cambios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Ocurrió un error crítico al limpiar los colores: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
        //Funcion para aplicar Colorimetria a las validaciones para identificarlas
        private void btnAplicarColores_Click(object sender, EventArgs e)
        {
            // 1. Validar que haya un censo cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();

                foreach (Excel.Worksheet wsCheck in _libroCenso.Worksheets)
                {
                    if (wsCheck.ProtectContents || wsCheck.ProtectDrawingObjects || wsCheck.ProtectScenarios)
                    {
                        hojasBloqueadas.Add("• " + wsCheck.Name);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCheck);
                }

                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(this,
                        "No se puede aplicar la macro de colores porque el libro contiene hojas protegidas:\n\n" +
                        listaHojas + "\n\n" +
                        "Desprotege estas hojas desde la pestaña Revisión y vuelve a intentarlo.",
                        "Validación Interrumpida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // =========================================================================
                // FASE 2: ESTRUCTURA DE FÓRMULAS UNIVERSALES (MATRIZ DE TRADUCCIÓN)
                // =========================================================================
                // Definimos las 7 reglas en inglés universal para traducirlas nativamente en cada hoja.
                // Formato 1 al 7 tal como venían definidos en tu aplicativo original.
                string[] formulasIngles = new string[]
                {
            "=AND($A$1<>\"\",SEARCH(\"igual o mayor\",A1)>0)",  // Regla 2
            "=AND($A$1<>\"\",SEARCH(\"igual o menor\",A1)>0)",  // Regla 3
            "=AND($A$1<>\"\",SEARCH(\"pase a la pregunta\",A1)>0)", // Regla 4
            "=AND($A$1<>\"\",SEARCH(\"en blanco\",A1)>0)",      // Regla 5
            "=AND($A$1<>\"\",SEARCH(\"no puede registrar\",A1)>0)", // Regla 6
            "=AND($A$1<>\"\",SUM(A1)>0)",                       // Regla 7
            "=AND($A$1<>\"\",SEARCH(\"la pregunta\",A1)>0)"     // Regla 1
                };

                int hojasConfiguradas = 0;

                foreach (Excel.Worksheet ws in _libroCenso.Worksheets)
                {
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;
                    Excel.Range celdaDummy = null;

                    try
                    {
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;

                        // 2.1 BÚSQUEDA PREVIA: Si la hoja ya tiene al menos una de nuestras reglas, asumimos que ya fue procesada
                        bool yaExisteRegra = false;
                        for (int i = formatos.Count; i >= 1; i--)
                        {
                            Excel.FormatCondition fcCheck = null;
                            try
                            {
                                object objFc = formatos[i];
                                if (objFc is Excel.FormatCondition)
                                {
                                    fcCheck = (Excel.FormatCondition)objFc;
                                    if (fcCheck.Type == (int)Excel.XlFormatConditionType.xlExpression)
                                    {
                                        string formulaText = fcCheck.Formula1 ?? "";
                                        if (formulaText.Replace(" ", "").ToUpper().Contains("$A$1<>\"\""))
                                        {
                                            yaExisteRegra = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            catch { }
                            finally
                            {
                                if (fcCheck != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fcCheck);
                            }
                        }

                        if (yaExisteRegra)
                        {
                            continue; // Nos saltamos esta hoja para no duplicar configuraciones
                        }

                        // 2.2 INYECCIÓN DE LAS 7 REGLAS CON TRADUCCIÓN NATIVA SENSIBLE A IDIOMA
                        celdaDummy = ws.Cells[1048576, 16384]; // Celda XFD1048576

                        for (int idx = 0; idx < formulasIngles.Length; idx++)
                        {
                            // Forzar traducción al lenguaje de Excel local
                            celdaDummy.Formula = formulasIngles[idx];
                            string formulaLocal = celdaDummy.FormulaLocal;
                            celdaDummy.Clear();

                            Excel.FormatCondition nuevaRegla = (Excel.FormatCondition)formatos.Add(
                                Excel.XlFormatConditionType.xlExpression,
                                Type.Missing,
                                formulaLocal);

                            // Configuración estética individual según tu código de origen
                            switch (idx)
                            {
                                case 0: // Formato 2: Accent6 con Tinte (0.5999)
                                    nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                    nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent6;
                                    nuevaRegla.Interior.TintAndShade = 0.599963377788629;
                                    break;

                                case 1: // Formato 3: Accent2 con Tinte (0.5999)
                                    nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                    nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent2;
                                    nuevaRegla.Interior.TintAndShade = 0.599963377788629;
                                    break;

                                case 2: // Formato 4: Color Fijo Decimal 16764159
                                case 3: // Formato 5: Mismo comportamiento
                                case 4: // Formato 6: Mismo comportamiento
                                    nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                    nuevaRegla.Interior.Color = 16764159;
                                    nuevaRegla.Interior.TintAndShade = 0;
                                    break;

                                case 5: // Formato 7: Accent5 con Tinte (0.7999)
                                    nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                    nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent5;
                                    nuevaRegla.Interior.TintAndShade = 0.799981688894314;
                                    break;

                                case 6: // Formato 1: Accent4 con Tinte (0.5999)
                                    nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                    nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent4;
                                    nuevaRegla.Interior.TintAndShade = 0.599963377788629;
                                    break;
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(nuevaRegla);
                        }

                        hojasConfiguradas++;
                    }
                    finally
                    {
                        if (celdaDummy != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                        if (formatos != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(formatos);
                        if (celdasHoja != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdasHoja);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    }
                }

                MessageBox.Show(this,
                    $"Funcion de colorimetría aplicada con éxito.\n\nSe configuraron {hojasConfiguradas} hoja(s) del censo de forma segura.",
                    "Reglas de Control Añadidas", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error crítico al inyectar las macros de control A1: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
        //Funcion para revertir Colorimetria aplicada por btnAplicarColores_Click
        private void btnLimpiarColores_Click(object sender, EventArgs e)
        {
            // 1. Validar que haya un censo cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();

                foreach (Excel.Worksheet wsCheck in _libroCenso.Worksheets)
                {
                    if (wsCheck.ProtectContents || wsCheck.ProtectDrawingObjects || wsCheck.ProtectScenarios)
                    {
                        hojasBloqueadas.Add("• " + wsCheck.Name);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCheck);
                }

                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(this,
                        "No se puede remover el formato de colores debido a que el libro tiene hojas protegidas:\n\n" +
                        listaHojas + "\n\n" +
                        "Desprotege estas hojas desde la pestaña Revisión y reintenta la operación.",
                        "Operación Interrumpida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // =========================================================================
                // FASE 2: PURGA Y ELIMINACIÓN DE FORMATOS MEDIANTE RECORRIDO INVERSO
                // =========================================================================
                int hojasLimpiadas = 0;

                foreach (Excel.Worksheet ws in _libroCenso.Worksheets)
                {
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;
                    bool cambioDetectado = false;

                    try
                    {
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;

                        // ESCANEO INVERSO OBLIGATORIO: Evita desajustes en el puntero de la lista al borrar
                        for (int i = formatos.Count; i >= 1; i--)
                        {
                            Excel.FormatCondition fc = null;
                            try
                            {
                                object objFc = formatos[i];
                                if (objFc is Excel.FormatCondition)
                                {
                                    fc = (Excel.FormatCondition)objFc;

                                    if (fc.Type == (int)Excel.XlFormatConditionType.xlExpression)
                                    {
                                        string formulaCondicion = "";
                                        try { formulaCondicion = fc.Formula1; } catch { }

                                        if (!string.IsNullOrEmpty(formulaCondicion))
                                        {
                                            string formulaLimpia = formulaCondicion.Replace(" ", "").ToUpper();

                                            // Buscamos nuestro patrón común anclado a la celda de control $A$1
                                            if (formulaLimpia.Contains("$A$1<>\"\""))
                                            {
                                                fc.Delete();
                                                cambioDetectado = true;
                                            }
                                        }
                                    }
                                }
                            }
                            catch { }
                            finally
                            {
                                if (fc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fc);
                            }
                        }

                        if (cambioDetectado)
                        {
                            hojasLimpiadas++;
                        }
                    }
                    finally
                    {
                        if (formatos != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(formatos);
                        if (celdasHoja != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdasHoja);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    }
                }

                // =========================================================================
                // FASE 3: CONCLUSIÓN UX
                // =========================================================================
                if (hojasLimpiadas > 0)
                {
                    MessageBox.Show(this,
                        $"Se eliminaron por completo las reglas de colorimetria vinculadas al archivo.\n\nSe limpiaron exitosamente {hojasLimpiadas} hoja(s).",
                        "SAVCNG - Limpieza Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this,
                        "No se detectaron formatos condicionales vinculados al archivo actual.\n\nEl libro ya se encuentra limpio.",
                        "SAVCNG - Sin Cambios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Ocurrió un error inesperado al limpiar las macros de formato: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}
