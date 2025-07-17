using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Verificador_de_Arquivo
{
    public class LoginForm : Form
    {
        private TextBox txtUsuario;
        private TextBox txtSenha;
        private Button btnRun;

        public string Usuario => txtUsuario.Text.Trim();
        public string Senha => txtSenha.Text;

        public LoginForm()
        {
            InitializeComponents();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= 0x00020000; // CS_DROPSHADOW
                return cp;
            }
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            // Gradiente verde azulado frio
            using (LinearGradientBrush brush = new LinearGradientBrush(this.ClientRectangle,
                Color.FromArgb(20, 40, 50), Color.FromArgb(40, 80, 90), 45f))
            {
                e.Graphics.FillRectangle(brush, this.ClientRectangle);
            }
        }

        private void InitializeComponents()
        {
            // Formulário
            Text = "Credenciais para Acesso ao SEI";
            ClientSize = new Size(340, 210);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            DoubleBuffered = true;
            Font = new Font("Segoe UI", 11F, FontStyle.Regular);

            // Caixa de texto: Usuário
            txtUsuario = new TextBox
            {
                Location = new Point(40, 35),
                Size = new Size(260, 32),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                PlaceholderText = "Usuário",
                BackColor = Color.FromArgb(240, 255, 250), // mint branco
                ForeColor = Color.FromArgb(20, 40, 50),    // verde escuro
                BorderStyle = BorderStyle.None,
                TextAlign = HorizontalAlignment.Center
            };
            txtUsuario.Region = new Region(GetRoundedRect(txtUsuario.ClientRectangle, 12));

            // Caixa de texto: Senha
            txtSenha = new TextBox
            {
                Location = new Point(40, 80),
                Size = new Size(260, 32),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                PlaceholderText = "Senha",
                UseSystemPasswordChar = true,
                BackColor = Color.FromArgb(240, 255, 250),
                ForeColor = Color.FromArgb(20, 40, 50),
                BorderStyle = BorderStyle.None,
                TextAlign = HorizontalAlignment.Center
            };
            txtSenha.Region = new Region(GetRoundedRect(txtSenha.ClientRectangle, 12));

            ApplyFocusGlow(txtUsuario);
            ApplyFocusGlow(txtSenha);

            // Botão de login
            btnRun = new Button
            {
                Location = new Point(90, 135),
                Size = new Size(160, 42),
                Text = "Entrar",
                Font = new Font("Segoe UI Semibold", 13),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 128, 128), // Teal
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnRun.FlatAppearance.BorderSize = 0;
            btnRun.Region = new Region(GetRoundedRect(btnRun.ClientRectangle, 18));

            // Efeitos no botão
            btnRun.MouseEnter += (s, e) => btnRun.BackColor = Color.FromArgb(0, 150, 150);
            btnRun.MouseLeave += (s, e) => btnRun.BackColor = Color.FromArgb(0, 128, 128);
            btnRun.MouseDown += (s, e) => btnRun.BackColor = Color.FromArgb(0, 100, 100);
            btnRun.MouseUp += (s, e) => btnRun.BackColor = Color.FromArgb(0, 150, 150);

            btnRun.Click += BtnRun_Click;

            // Adiciona controles
            Controls.Add(txtUsuario);
            Controls.Add(txtSenha);
            Controls.Add(btnRun);
        }

        private void ApplyFocusGlow(TextBox txt)
        {
            txt.Enter += (s, e) => txt.BackColor = Color.White;
            txt.Leave += (s, e) => txt.BackColor = Color.FromArgb(240, 255, 250);
        }

        private GraphicsPath GetRoundedRect(Rectangle bounds, int radius)
        {
            int diameter = radius * 2;
            GraphicsPath path = new GraphicsPath();

            path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();

            return path;
        }

        private void BtnRun_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Usuario) || string.IsNullOrWhiteSpace(Senha))
            {
                MessageBox.Show(
                    "Informe usuário e senha.",
                    "Atenção",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
