using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace PassGen // <--- CHANGE THIS to match your project name
{
    public class Form1 : Form
    {
        // 1. Declare class-level fields so methods can see them
        private TextBox txtUser, txtPass, txtEmail, txtSecure;
        private RadioButton radioLen8, radioLen12, radioLen16;
        private Button btnGenerate, btnClear, btnCreateDoc, btnDrafts;

        // Theme Colors
        private Color bgDark = Color.FromArgb(10, 18, 42);
        private Color accentBlue = Color.FromArgb(0, 102, 204);

        public Form1()
        {
            // Form Window Setup
            this.Text = "Password Generator";
            this.Size = new Size(500, 600);
            this.BackColor = bgDark;
            this.ForeColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5f);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;

            InitializeCustomComponents();
        }

        private void InitializeCustomComponents()
        {
            // Title
            Label lblTitle = new Label
            {
                Text = "Password Generator",
                Font = new Font("Segoe UI", 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 80
            };

            // Inputs
            Label lblUser = new Label { Text = "User:", Location = new Point(20, 100), AutoSize = true };
            txtUser = CreateStyledInput(120, 98, 340);

            Label lblPass = new Label { Text = "Password:", Location = new Point(20, 140), AutoSize = true };
            txtPass = CreateStyledInput(120, 138, 340);

            // Radio Buttons (Grouped by sharing the same parent: 'this')
            radioLen8 = new RadioButton { Text = "8 Characters", Location = new Point(30, 190), AutoSize = true };
            radioLen12 = new RadioButton { Text = "12 Characters", Location = new Point(30, 220), AutoSize = true, Checked = true };
            radioLen16 = new RadioButton { Text = "16 Characters", Location = new Point(30, 250), AutoSize = true };

            Label lblWip = new Label
            {
                Text = "Other features are WiP",
                ForeColor = Color.DarkGray,
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                Location = new Point(250, 220),
                AutoSize = true
            };

            // Generate Button
            btnGenerate = CreateFlatButton("Generate", 280, 290, 180, 40);
            btnGenerate.Click += BtnGenerate_Click;

            Label lblNote = new Label
            {
                Text = "By default it will generate a 12 characters password",
                Font = new Font("Segoe UI", 8),
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(0, 350),
                Size = new Size(500, 20)
            };

            // Separator
            Panel pnlLine = new Panel { BackColor = accentBlue, Location = new Point(0, 380), Size = new Size(500, 3) };

            // Bottom Section
            Label lblEmail = new Label { Text = "Email to send credentials:", Location = new Point(20, 400), AutoSize = true };
            txtEmail = CreateStyledInput(220, 398, 240);

            Label lblSecure = new Label { Text = "Secure code", Location = new Point(20, 430), AutoSize = true };
            txtSecure = CreateStyledInput(20, 455, 150);
            txtSecure.Text = "Ph6?UW";

            btnCreateDoc = CreateFlatButton("Create document", 280, 430, 180, 35);
            btnDrafts = CreateFlatButton("Create drafts", 280, 475, 180, 35);

            // Footer Clear Button
            btnClear = new Button
            {
                Text = "Clear Form",
                Dock = DockStyle.Bottom,
                Height = 40,
                BackColor = Color.White,
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            btnClear.FlatAppearance.BorderSize = 0;
            btnClear.Click += BtnClear_Click;

            // Add all to Form
            this.Controls.AddRange(new Control[] {
                lblTitle, lblUser, txtUser, lblPass, txtPass, radioLen8, radioLen12, radioLen16,
                lblWip, btnGenerate, lblNote, pnlLine, lblEmail, txtEmail,
                lblSecure, txtSecure, btnCreateDoc, btnDrafts, btnClear
            });
        }

        // Logic for Generation
        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            int length = 12;
            if (radioLen8.Checked) length = 8;
            if (radioLen16.Checked) length = 16;

            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*?";
            var random = new Random();
            txtPass.Text = new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        // Logic for Clear
        private void BtnClear_Click(object sender, EventArgs e)
        {
            txtUser.Clear();
            txtPass.Clear();
            txtEmail.Clear();
            txtSecure.Text = "Ph6?UW";
            radioLen12.Checked = true;
        }

        // Helper Methods for Styling
        private TextBox CreateStyledInput(int x, int y, int width)
        {
            return new TextBox
            {
                Location = new Point(x, y),
                Width = width,
                BackColor = Color.White,
                ForeColor = Color.Black,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private Button CreateFlatButton(string text, int x, int y, int w, int h)
        {
            Button btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                BackColor = Color.White,
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}