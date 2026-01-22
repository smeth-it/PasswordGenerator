using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace PassGen
{
    public partial class Form1 : Form
    {
        private const string TempFolder = @"C:\Temp";
        private const string WordFile = @"C:\Temp\Credentials.docx";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Directory.CreateDirectory(TempFolder);
            textBox2.Text = GeneratePassword(6, DefaultCharset());
            radioLen12.Checked = true;
        }

        // =====================================================
        // PASSWORD GENERATION
        // =====================================================

        private void buttonGenerate_Click(object sender, EventArgs e)
        {
            int length = radioLen8.Checked ? 8 :
                         radioLen16.Checked ? 16 : 12;

            textBox1.Text = GeneratePassword(length, BuildCharset());
        }

        private string GeneratePassword(int length, string charset)
        {
            if (string.IsNullOrWhiteSpace(charset))
                throw new InvalidOperationException("Select at least one character type.");

            var sb = new StringBuilder(length);
            byte[] buffer = new byte[4];

            using var rng = RandomNumberGenerator.Create();
            for (int i = 0; i < length; i++)
            {
                rng.GetBytes(buffer);
                int idx = BitConverter.ToInt32(buffer, 0) & int.MaxValue;
                sb.Append(charset[idx % charset.Length]);
            }
            return sb.ToString();
        }

        private string DefaultCharset() =>
            "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@,.?";

        private string BuildCharset()
        {
            var sb = new StringBuilder();

            if (checkBox4.Checked) sb.Append("0123456789");
            if (checkBox5.Checked) sb.Append("ABCDEFGHIJKLMNOPQRSTUVWXYZ");
            if (checkBox6.Checked) sb.Append("abcdefghijklmnopqrstuvwxyz");
            if (checkBox7.Checked) sb.Append("!@,.?");

            return sb.ToString();
        }

        // =====================================================
        // RESET
        // =====================================================

        private void buttonReset_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox3.Clear();
            richTextBox1.Clear();
            textBox2.Text = GeneratePassword(6, DefaultCharset());
        }

        // =====================================================
        // WORD + OUTLOOK
        // =====================================================

        private void buttonSend_Click(object sender, EventArgs e)
        {
            CreateWordDocument();
            SendMailWithAttachment();
            SendMailWithCode();
            CleanupWordFile();
        }

        private void CreateWordDocument()
        {
            var data = ReadForm();

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false };
                doc = wordApp.Documents.Add();

                doc.Content.Text =
                    $"Your new password for the account {data.User} is {data.Password}";

                doc.Password = data.Code;
                doc.SaveAs2(WordFile);
            }
            finally
            {
                doc?.Close(false);
                wordApp?.Quit(false);
            }
        }

        private void SendMailWithAttachment()
        {
            var data = ReadForm();

            var outlook = new Outlook.Application();
            var mail = (Outlook.MailItem)outlook.CreateItem(Outlook.OlItemType.olMailItem);

            mail.To = data.Recipient;
            mail.Subject = "New Credentials";
            mail.Body =
                "This email contains the encrypted Word document.\r\n" +
                "The password will arrive in a second email.";

            mail.Attachments.Add(WordFile);
            mail.Display(false);
            mail.Save();
        }

        private void SendMailWithCode()
        {
            var data = ReadForm();

            var outlook = new Outlook.Application();
            var mail = (Outlook.MailItem)outlook.CreateItem(Outlook.OlItemType.olMailItem);

            mail.To = data.Recipient;
            mail.Subject = "New Credentials – Access Code";
            mail.Body = $"The encryption code for the Word file is:\r\n\r\n{data.Code}";
            mail.Display(false);
            mail.Save();
        }

        private void CleanupWordFile()
        {
            try
            {
                if (File.Exists(WordFile))
                    File.Delete(WordFile);
            }
            catch
            {
                // intentional silence – file deletion must not block UX
            }
        }

        // =====================================================
        // UTIL
        // =====================================================

        private (string Password, string Code, string User, string Recipient) ReadForm()
        {
            return (
                textBox1.Text.Trim(),
                textBox2.Text.Trim(),
                textBox3.Text.Trim(),
                richTextBox1.Text.Trim()
            );
        }
    }
}
