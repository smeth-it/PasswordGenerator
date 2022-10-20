using System;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace PassGen
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            int len = 6;
            const string validchar = "abcdefghijklmnopqrstuwxyzABCDEFGHIJKLMNOPQRSTUWZ1234567890!@,.?";
            StringBuilder result = new StringBuilder();
            Random rand = new Random();
            while (0 < len--)
            { result.Append(validchar[rand.Next(validchar.Length)]); }
            textBox2.Text = result.ToString();
            string CreatedPass;
            string SecureCode;
            string Usr;
            string besendto;
            CreatedPass = textBox1.Text;
            SecureCode = textBox2.Text;
            Usr = textBox3.Text;
            besendto = richTextBox1.Text;

        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        private void label3_Click(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                PassGen(8);
            }
            else if (checkBox2.Checked)
            {
                PassGen(12);
            }
            else if (checkBox3.Checked)
            {
                PassGen(16);
            }
            else { PassGen(12); }
        }
        public void PassGen(int len)
        {
            const string validchar = "abcdefghijklmnopqrstuwxyzABCDEFGHIJKLMNOPQRSTUWZ1234567890!@,.?";
            StringBuilder result = new StringBuilder();
            Random rand = new Random();
            while (0 < len--)
            { result.Append(validchar[rand.Next(validchar.Length)]); }
            textBox1.Text = result.ToString();
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            int len = 6;
            const string validchar = "abcdefghijklmnopqrstuwxyzABCDEFGHIJKLMNOPQRSTUWZ1234567890!@,.?";
            StringBuilder result = new StringBuilder();
            Random rand = new Random();
            while (0 < len--)
            { result.Append(validchar[rand.Next(validchar.Length)]); }
            textBox2.Text = result.ToString();
            richTextBox1.Text = ("");
            textBox3.Clear();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            CreateDocument();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            CreateMailItem1();
            CreateMailItem2();
        }
        private void CreateDocument()
        {
            try
            {
                string CreatedPass;
                string SecureCode;
                string Usr;
                string besendto;
                CreatedPass = textBox1.Text;
                SecureCode = textBox2.Text;
                Usr = textBox3.Text;
                besendto = richTextBox1.Text;
                //Create an instance for word app  
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

                //Set animation status for word application  
                winword.ShowAnimation = false;

                //Set status for word application is to be visible or not.  
                winword.Visible = false;

                //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;

                //Create a new document  
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                //Add header into the document  
                //foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                // {
                //Get the header range and add the header details.  
                //  Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                //  headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                //headerRange.Font.Size = 10;
                // headerRange.Text = "Header text goes here";
                //  }

                //Add the footers into the document  
                //foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                //{
                //Get the footer range and add the footer details.  
                //Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                //footerRange.Font.Size = 10;
                //footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //footerRange.Text = "Footer text goes here";
                //}

                //adding text to document  
                document.Content.SetRange(0, 0);
                document.Content.Text = "Your new password for the account " + Usr + " is " + CreatedPass;

                //Add paragraph with Heading 1 style  
                // Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                //object styleHeading1 = "Heading 1";
                //para1.Range.set_Style(ref styleHeading1);
                //para1.Range.Text = "Para 1 text";
                //para1.Range.InsertParagraphAfter();

                //Add paragraph with Heading 2 style  
                //Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref missing);
                //object styleHeading2 = "Heading 2";
                //para2.Range.set_Style(ref styleHeading2);
                //para2.Range.Text = "Para 2 text";
                //para2.Range.InsertParagraphAfter();

                //Create a 5X5 table and insert some dummy record  
                //Table firstTable = document.Tables.Add(para1.Range, 5, 5, ref missing, ref missing);

                // firstTable.Borders.Enable = 1;
                //foreach (Row row in firstTable.Rows)
                //{
                //foreach (Cell cell in row.Cells)
                //{
                //Header row  
                //                if (cell.RowIndex == 1)
                //{
                //cell.Range.Text = "Column " + cell.ColumnIndex.ToString();
                //cell.Range.Font.Bold = 1;
                //other format properties goes here  
                //cell.Range.Font.Name = "verdana";
                //cell.Range.Font.Size = 10;
                //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                //cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                //Center alignment for the Header cells  
                //cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                //cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                //
                //}
                //Data row  
                //      else
                //{
                //cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                //}
                //}
                //}

                //Save the document
                object filename = @"c:\Temp\Credentials.docx";
                object Password = SecureCode;
                document.Password = SecureCode;
                document.SaveAs(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                MessageBox.Show("Document created successfully !");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateMailItem1()
    {
        string CreatedPass;
        string SecureCode;
        string Usr;
        string besendto;
        CreatedPass = textBox1.Text;
        SecureCode = textBox2.Text;
        Usr = textBox3.Text;
        besendto = richTextBox1.Text;

        Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
        MailItem item = app.CreateItem((OlItemType.olMailItem));
        item.BodyFormat = OlBodyFormat.olFormatHTML;
        item.To = besendto;
        item.Body = "This email contains the encrypted Word file, please check the second email to find the code.";
        item.Subject = "New Credentials";
        
        item.Display(false);
        item.Save();

    }
        private void CreateMailItem2()
    {
        string CreatedPass;
        string SecureCode;
        string Usr;
        string besendto;
        CreatedPass = textBox1.Text;
        SecureCode = textBox2.Text;
        Usr = textBox3.Text;
        besendto = richTextBox1.Text;

        Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
        MailItem item = app.CreateItem((OlItemType.olMailItem));
        item.BodyFormat = OlBodyFormat.olFormatHTML;
        item.To = besendto;
        item.Body = "The encryption code for the Word file is: " + SecureCode;
        item.Subject = "New Credentials - 2";
        item.Display(false);
        item.Save();

    }
    }
}

