
using System;
using System.Drawing;
using System.Windows.Forms;

namespace SAVCNG_ExcelDNA
{
    // Tipos de validación disponibles
    public enum TipoValidacionContacto
    {
        Horarios,
        CorreoElectronico,
        Telefono
    }

    public sealed class FrmSeleccionValidacionContacto : Form
    {
        private RadioButton rbHorarios;
        private RadioButton rbCorreo;
        private RadioButton rbTelefono;
        private Button btnAceptar;
        private Button btnCancelar;

        // Resultado que recibirá el formulario principal
        public TipoValidacionContacto TipoSeleccionado
        {
            get;
            private set;
        }

        public FrmSeleccionValidacionContacto()
        {
            InicializarVentana();
        }

        private void InicializarVentana()
        {
            // Configuración general
            this.Text = "SAVCNG - Seleccionar validación";
            this.ClientSize = new Size(510, 310);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.AutoScaleMode = AutoScaleMode.Font;
            this.Font = new Font("Segoe UI", 9F);
            this.BackColor = Color.White;

            // Encabezado
            Label lblTitulo = new Label();
            lblTitulo.Text = "Selecciona el tipo de validación";
            lblTitulo.Font = new Font(
                "Segoe UI", 11F, FontStyle.Bold);
            lblTitulo.Location = new Point(20, 15);
            lblTitulo.Size = new Size(470, 28);

            Label lblDescripcion = new Label();
            lblDescripcion.Text =
                "Selecciona una de las tres validaciones " +
                Environment.NewLine +
                "para aplicarla al rango capturado.";
            lblDescripcion.Location = new Point(20, 47);
            lblDescripcion.Size = new Size(470, 40);

            // Grupo de opciones
            GroupBox grupo = new GroupBox();
            grupo.Text = "Validaciones disponibles";
            grupo.Location = new Point(20, 88);
            grupo.Size = new Size(470, 180);

            // Opción 1: Horarios
            rbHorarios = new RadioButton();
            rbHorarios.Text = "Horarios";
            rbHorarios.Location = new Point(18, 27);
            rbHorarios.AutoSize = true;

            Label lblHorarios = new Label();
            lblHorarios.Text =
                "Verifica formatos de horarios de atención.";
            lblHorarios.Location = new Point(38, 49);
            lblHorarios.Size = new Size(410, 20);

            // Opción 2: Correo electrónico
            rbCorreo = new RadioButton();
            rbCorreo.Text = "Correo electrónico de contacto";
            rbCorreo.Location = new Point(18, 76);
            rbCorreo.AutoSize = true;

            Label lblCorreo = new Label();
            lblCorreo.Text =
                "Verifica la estructura de direcciones de correo electrónico.";
            lblCorreo.Location = new Point(38, 98);
            lblCorreo.Size = new Size(410, 20);

            // Opción 3: Teléfono
            rbTelefono = new RadioButton();
            rbTelefono.Text = "Teléfono de contacto";
            rbTelefono.Location = new Point(18, 125);
            rbTelefono.AutoSize = true;

            Label lblTelefono = new Label();
            lblTelefono.Text =
                "Verifica el formato de números telefónicos de contacto.";
            lblTelefono.Location = new Point(38, 147);
            lblTelefono.Size = new Size(410, 20);

            grupo.Controls.Add(rbHorarios);
            grupo.Controls.Add(lblHorarios);
            grupo.Controls.Add(rbCorreo);
            grupo.Controls.Add(lblCorreo);
            grupo.Controls.Add(rbTelefono);
            grupo.Controls.Add(lblTelefono);

            // Botón Seleccionar
            btnAceptar = new Button();
            btnAceptar.Text = "Seleccionar";
            btnAceptar.Location = new Point(353, 280);
            btnAceptar.Size = new Size(137, 32);
            btnAceptar.Enabled = false;
            btnAceptar.BackColor =
                Color.FromArgb(0, 112, 192);
            btnAceptar.ForeColor = Color.White;
            btnAceptar.FlatStyle = FlatStyle.Flat;
            btnAceptar.FlatAppearance.BorderSize = 0;
            btnAceptar.Click += BtnAceptar_Click;

            // Botón Cancelar
            btnCancelar = new Button();
            btnCancelar.Text = "Cancelar";
            btnCancelar.Location = new Point(242, 280);
            btnCancelar.Size = new Size(100, 32);
            btnCancelar.DialogResult = DialogResult.Cancel;

            // Habilitar Seleccionar solamente
            // cuando el usuario elija una opción.
            rbHorarios.CheckedChanged += Opcion_CheckedChanged;
            rbCorreo.CheckedChanged += Opcion_CheckedChanged;
            rbTelefono.CheckedChanged += Opcion_CheckedChanged;

            this.Controls.Add(lblTitulo);
            this.Controls.Add(lblDescripcion);
            this.Controls.Add(grupo);
            this.Controls.Add(btnAceptar);
            this.Controls.Add(btnCancelar);

            this.AcceptButton = btnAceptar;
            this.CancelButton = btnCancelar;
        }

        private void Opcion_CheckedChanged(
            object sender, EventArgs e)
        {
            btnAceptar.Enabled =
                rbHorarios.Checked ||
                rbCorreo.Checked ||
                rbTelefono.Checked;
        }

        private void BtnAceptar_Click(
            object sender, EventArgs e)
        {
            if (rbHorarios.Checked)
            {
                TipoSeleccionado =
                    TipoValidacionContacto.Horarios;
            }
            else if (rbCorreo.Checked)
            {
                TipoSeleccionado =
                    TipoValidacionContacto.CorreoElectronico;
            }
            else if (rbTelefono.Checked)
            {
                TipoSeleccionado =
                    TipoValidacionContacto.Telefono;
            }
            else
            {
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
