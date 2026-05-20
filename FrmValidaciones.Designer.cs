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
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label5 = new System.Windows.Forms.Label();
            this.chkNS = new System.Windows.Forms.CheckBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.chkFormatoTexto = new System.Windows.Forms.CheckBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblCenso
            // 
            this.lblCenso.AutoSize = true;
            this.lblCenso.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCenso.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblCenso.Location = new System.Drawing.Point(12, 184);
            this.lblCenso.Name = "lblCenso";
            this.lblCenso.Size = new System.Drawing.Size(356, 31);
            this.lblCenso.TabIndex = 0;
            this.lblCenso.Text = "Censo cargado: (Ninguno)";
            // 
            // chkDecimales
            // 
            this.chkDecimales.AutoSize = true;
            this.chkDecimales.Location = new System.Drawing.Point(27, 130);
            this.chkDecimales.Name = "chkDecimales";
            this.chkDecimales.Size = new System.Drawing.Size(484, 33);
            this.chkDecimales.TabIndex = 1;
            this.chkDecimales.Text = "Validación: Formato Decimales (Enteros)";
            this.chkDecimales.UseVisualStyleBackColor = true;
            // 
            // btnCapturarRango
            // 
            this.btnCapturarRango.Location = new System.Drawing.Point(59, 71);
            this.btnCapturarRango.Name = "btnCapturarRango";
            this.btnCapturarRango.Size = new System.Drawing.Size(252, 83);
            this.btnCapturarRango.TabIndex = 2;
            this.btnCapturarRango.Text = "Capturar Rango";
            this.btnCapturarRango.UseVisualStyleBackColor = true;
            this.btnCapturarRango.Click += new System.EventHandler(this.btnCapturarRango_Click);
            // 
            // btnAplicar
            // 
            this.btnAplicar.Location = new System.Drawing.Point(696, 858);
            this.btnAplicar.Name = "btnAplicar";
            this.btnAplicar.Size = new System.Drawing.Size(234, 91);
            this.btnAplicar.TabIndex = 3;
            this.btnAplicar.Text = "Aplicar Validación";
            this.btnAplicar.UseVisualStyleBackColor = true;
            this.btnAplicar.Click += new System.EventHandler(this.btnAplicar_Click);
            // 
            // chkCatalogos
            // 
            this.chkCatalogos.AutoSize = true;
            this.chkCatalogos.Location = new System.Drawing.Point(1201, 130);
            this.chkCatalogos.Name = "chkCatalogos";
            this.chkCatalogos.Size = new System.Drawing.Size(278, 33);
            this.chkCatalogos.TabIndex = 4;
            this.chkCatalogos.Text = "Validación: Catálogos";
            this.chkCatalogos.UseVisualStyleBackColor = true;
            this.chkCatalogos.CheckedChanged += new System.EventHandler(this.chkCatalogos_CheckedChanged);
            // 
            // lblPregunta
            // 
            this.lblPregunta.AutoSize = true;
            this.lblPregunta.Location = new System.Drawing.Point(22, 29);
            this.lblPregunta.Name = "lblPregunta";
            this.lblPregunta.Size = new System.Drawing.Size(228, 29);
            this.lblPregunta.TabIndex = 5;
            this.lblPregunta.Text = "Pregunta detectada:";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1662, 153);
            this.panel1.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Narrow", 16.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(292, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(1251, 52);
            this.label1.TabIndex = 2;
            this.label1.Text = "Sistema Automatizado de Validaciones para Censos Nacionales de Gobierno";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::SAVCNG_ExcelDNA.Properties.Resources.INEGI_Logotipo_1;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(148, 153);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.tabControl1);
            this.panel2.Controls.Add(this.btnCapturarRango);
            this.panel2.Controls.Add(this.btnAplicar);
            this.panel2.Location = new System.Drawing.Point(12, 245);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1662, 966);
            this.panel2.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(25, 882);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(620, 48);
            this.label4.TabIndex = 6;
            this.label4.Text = "Paso 3: Aplicar validaciones";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(17, 177);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(913, 48);
            this.label3.TabIndex = 5;
            this.label3.Text = "Paso 2: Seleccionar validaciones a aplicar";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(17, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(913, 48);
            this.label2.TabIndex = 4;
            this.label2.Text = "Paso 1: Seleccionar el rango a validar";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(24, 239);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1620, 599);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.chkNS);
            this.tabPage1.Controls.Add(this.lblPregunta);
            this.tabPage1.Controls.Add(this.chkCatalogos);
            this.tabPage1.Controls.Add(this.chkDecimales);
            this.tabPage1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage1.Location = new System.Drawing.Point(8, 43);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1604, 548);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Validaciones Generales";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(810, 41);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(228, 29);
            this.label5.TabIndex = 7;
            this.label5.Text = "Pregunta detectada:";
            // 
            // chkNS
            // 
            this.chkNS.AutoSize = true;
            this.chkNS.Location = new System.Drawing.Point(664, 130);
            this.chkNS.Name = "chkNS";
            this.chkNS.Size = new System.Drawing.Size(203, 33);
            this.chkNS.TabIndex = 6;
            this.chkNS.Text = "Validación: NS";
            this.chkNS.UseVisualStyleBackColor = true;
            this.chkNS.CheckedChanged += new System.EventHandler(this.chkNS_CheckedChanged);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.chkFormatoTexto);
            this.tabPage2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage2.Location = new System.Drawing.Point(8, 43);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1604, 548);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Validaciones Particulares";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // chkFormatoTexto
            // 
            this.chkFormatoTexto.AutoSize = true;
            this.chkFormatoTexto.Location = new System.Drawing.Point(411, 6);
            this.chkFormatoTexto.Name = "chkFormatoTexto";
            this.chkFormatoTexto.Size = new System.Drawing.Size(202, 33);
            this.chkFormatoTexto.TabIndex = 0;
            this.chkFormatoTexto.Text = "Formato Texto";
            this.chkFormatoTexto.UseVisualStyleBackColor = true;
            // 
            // FrmValidaciones
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1686, 1232);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lblCenso);
            this.Name = "FrmValidaciones";
            this.Text = "SAVCNG";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
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
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox chkFormatoTexto;
    }
}