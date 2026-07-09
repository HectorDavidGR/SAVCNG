namespace SAVCNG_ExcelDNA
{
    partial class FrmValidaciones
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblCenso = new System.Windows.Forms.Label();
            this.chkDecimales = new System.Windows.Forms.CheckBox();
            this.btnCapturarRango = new System.Windows.Forms.Button();
            this.btnAplicar = new System.Windows.Forms.Button();
            this.chkCatalogos = new System.Windows.Forms.CheckBox();
            this.lblPregunta = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.lblRangoSeleccionado = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.chkFechas = new System.Windows.Forms.CheckBox();
            this.chkBlancos = new System.Windows.Forms.CheckBox();
            this.chkFormatoTexto = new System.Windows.Forms.CheckBox();
            this.chkBloqueo = new System.Windows.Forms.CheckBox();
            this.chkNS = new System.Windows.Forms.CheckBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.chkEspClave = new System.Windows.Forms.CheckBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblCenso
            // 
            this.lblCenso.AutoSize = true;
            this.lblCenso.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCenso.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblCenso.Location = new System.Drawing.Point(8, 118);
            this.lblCenso.Name = "lblCenso";
            this.lblCenso.Size = new System.Drawing.Size(229, 20);
            this.lblCenso.TabIndex = 0;
            this.lblCenso.Text = "Censo cargado: (Ninguno)";
            // 
            // chkDecimales
            // 
            this.chkDecimales.AutoSize = true;
            this.chkDecimales.Location = new System.Drawing.Point(19, 26);
            this.chkDecimales.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chkDecimales.Name = "chkDecimales";
            this.chkDecimales.Size = new System.Drawing.Size(227, 22);
            this.chkDecimales.TabIndex = 1;
            this.chkDecimales.Text = "Formato Decimales (Enteros)";
            this.chkDecimales.UseVisualStyleBackColor = true;
            // 
            // btnCapturarRango
            // 
            this.btnCapturarRango.Location = new System.Drawing.Point(40, 46);
            this.btnCapturarRango.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnCapturarRango.Name = "btnCapturarRango";
            this.btnCapturarRango.Size = new System.Drawing.Size(168, 53);
            this.btnCapturarRango.TabIndex = 2;
            this.btnCapturarRango.Text = "Capturar Rango";
            this.btnCapturarRango.UseVisualStyleBackColor = true;
            this.btnCapturarRango.Click += new System.EventHandler(this.btnCapturarRango_Click);
            // 
            // btnAplicar
            // 
            this.btnAplicar.Location = new System.Drawing.Point(464, 549);
            this.btnAplicar.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnAplicar.Name = "btnAplicar";
            this.btnAplicar.Size = new System.Drawing.Size(156, 58);
            this.btnAplicar.TabIndex = 3;
            this.btnAplicar.Text = "Aplicar Validación";
            this.btnAplicar.UseVisualStyleBackColor = true;
            this.btnAplicar.Click += new System.EventHandler(this.btnAplicar_Click);
            // 
            // chkCatalogos
            // 
            this.chkCatalogos.AutoSize = true;
            this.chkCatalogos.Location = new System.Drawing.Point(404, 27);
            this.chkCatalogos.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chkCatalogos.Name = "chkCatalogos";
            this.chkCatalogos.Size = new System.Drawing.Size(98, 22);
            this.chkCatalogos.TabIndex = 4;
            this.chkCatalogos.Text = "Catálogos";
            this.chkCatalogos.UseVisualStyleBackColor = true;
            this.chkCatalogos.CheckedChanged += new System.EventHandler(this.chkCatalogos_CheckedChanged);
            // 
            // lblPregunta
            // 
            this.lblPregunta.AutoSize = true;
            this.lblPregunta.Location = new System.Drawing.Point(257, 62);
            this.lblPregunta.Name = "lblPregunta";
            this.lblPregunta.Size = new System.Drawing.Size(137, 17);
            this.lblPregunta.TabIndex = 5;
            this.lblPregunta.Text = "Pregunta detectada:";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(8, 7);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1108, 98);
            this.panel1.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Narrow", 16.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(195, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(760, 31);
            this.label1.TabIndex = 2;
            this.label1.Text = "Sistema Automatizado de Validaciones para Censos Nacionales de Gobierno";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::SAVCNG_ExcelDNA.Properties.Resources.INEGI_Logotipo_1;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(99, 98);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.lblRangoSeleccionado);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.lblPregunta);
            this.panel2.Controls.Add(this.tabControl1);
            this.panel2.Controls.Add(this.btnCapturarRango);
            this.panel2.Controls.Add(this.btnAplicar);
            this.panel2.Location = new System.Drawing.Point(8, 156);
            this.panel2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1108, 618);
            this.panel2.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(16, 565);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(413, 31);
            this.label4.TabIndex = 6;
            this.label4.Text = "Paso 3: Aplicar validaciones";
            // 
            // lblRangoSeleccionado
            // 
            this.lblRangoSeleccionado.AutoSize = true;
            this.lblRangoSeleccionado.Location = new System.Drawing.Point(699, 59);
            this.lblRangoSeleccionado.Name = "lblRangoSeleccionado";
            this.lblRangoSeleccionado.Size = new System.Drawing.Size(143, 17);
            this.lblRangoSeleccionado.TabIndex = 7;
            this.lblRangoSeleccionado.Text = "Rango Seleccionado:";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(11, 113);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(608, 31);
            this.label3.TabIndex = 5;
            this.label3.Text = "Paso 2: Seleccionar validaciones a aplicar";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(11, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(608, 31);
            this.label2.TabIndex = 4;
            this.label2.Text = "Paso 1: Seleccionar el rango a validar";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(16, 153);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1080, 383);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.chkEspClave);
            this.tabPage1.Controls.Add(this.chkFechas);
            this.tabPage1.Controls.Add(this.chkBlancos);
            this.tabPage1.Controls.Add(this.chkFormatoTexto);
            this.tabPage1.Controls.Add(this.chkBloqueo);
            this.tabPage1.Controls.Add(this.chkNS);
            this.tabPage1.Controls.Add(this.chkCatalogos);
            this.tabPage1.Controls.Add(this.chkDecimales);
            this.tabPage1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage1.Location = new System.Drawing.Point(4, 27);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage1.Size = new System.Drawing.Size(1072, 352);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Validaciones ";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // chkFechas
            // 
            this.chkFechas.AutoSize = true;
            this.chkFechas.Location = new System.Drawing.Point(404, 75);
            this.chkFechas.Margin = new System.Windows.Forms.Padding(4);
            this.chkFechas.Name = "chkFechas";
            this.chkFechas.Size = new System.Drawing.Size(79, 22);
            this.chkFechas.TabIndex = 14;
            this.chkFechas.Text = "Fechas";
            this.chkFechas.UseVisualStyleBackColor = true;
            this.chkFechas.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // chkBlancos
            // 
            this.chkBlancos.AutoSize = true;
            this.chkBlancos.Location = new System.Drawing.Point(24, 74);
            this.chkBlancos.Margin = new System.Windows.Forms.Padding(4);
            this.chkBlancos.Name = "chkBlancos";
            this.chkBlancos.Size = new System.Drawing.Size(84, 22);
            this.chkBlancos.TabIndex = 13;
            this.chkBlancos.Text = "Blancos";
            this.chkBlancos.UseVisualStyleBackColor = true;
            // 
            // chkFormatoTexto
            // 
            this.chkFormatoTexto.AutoSize = true;
            this.chkFormatoTexto.Location = new System.Drawing.Point(681, 27);
            this.chkFormatoTexto.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chkFormatoTexto.Name = "chkFormatoTexto";
            this.chkFormatoTexto.Size = new System.Drawing.Size(128, 22);
            this.chkFormatoTexto.TabIndex = 12;
            this.chkFormatoTexto.Text = "Formato Texto";
            this.chkFormatoTexto.UseVisualStyleBackColor = true;
            // 
            // chkBloqueo
            // 
            this.chkBloqueo.AutoSize = true;
            this.chkBloqueo.Location = new System.Drawing.Point(537, 27);
            this.chkBloqueo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chkBloqueo.Name = "chkBloqueo";
            this.chkBloqueo.Size = new System.Drawing.Size(93, 22);
            this.chkBloqueo.TabIndex = 11;
            this.chkBloqueo.Text = "Bloqueos";
            this.chkBloqueo.UseVisualStyleBackColor = true;
            // 
            // chkNS
            // 
            this.chkNS.AutoSize = true;
            this.chkNS.Location = new System.Drawing.Point(304, 28);
            this.chkNS.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.chkNS.Name = "chkNS";
            this.chkNS.Size = new System.Drawing.Size(51, 22);
            this.chkNS.TabIndex = 6;
            this.chkNS.Text = "NS";
            this.chkNS.UseVisualStyleBackColor = true;
            this.chkNS.CheckedChanged += new System.EventHandler(this.chkNS_CheckedChanged);
            // 
            // tabPage2
            // 
            this.tabPage2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage2.Location = new System.Drawing.Point(4, 27);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage2.Size = new System.Drawing.Size(1072, 352);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Revisión / Utilidades";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // chkEspClave
            // 
            this.chkEspClave.AutoSize = true;
            this.chkEspClave.Location = new System.Drawing.Point(130, 71);
            this.chkEspClave.Name = "chkEspClave";
            this.chkEspClave.Size = new System.Drawing.Size(165, 22);
            this.chkEspClave.TabIndex = 15;
            this.chkEspClave.Text = "Especifique (Claves)";
            this.chkEspClave.UseVisualStyleBackColor = false;
            // 
            // FrmValidaciones
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1136, 793);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lblCenso);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "FrmValidaciones";
            this.Text = "SAVCNG";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblCenso;
        private System.Windows.Forms.CheckBox chkDecimales;
        private System.Windows.Forms.Button btnCapturarRango;
        private System.Windows.Forms.Button btnAplicar;
        private System.Windows.Forms.CheckBox chkCatalogos;
        private System.Windows.Forms.Label lblPregunta;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.CheckBox chkNS;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblRangoSeleccionado;
        private System.Windows.Forms.CheckBox chkBloqueo;
        private System.Windows.Forms.CheckBox chkFormatoTexto;
        private System.Windows.Forms.CheckBox chkBlancos;
        private System.Windows.Forms.CheckBox chkFechas;
        private System.Windows.Forms.CheckBox chkEspClave;
    }
}