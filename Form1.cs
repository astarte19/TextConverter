using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using SautinSoft.Document;
using Microsoft.Office.Interop.Word;
using SautinSoft;
using Spire.Pdf;
using Aspose.Words;




namespace Kursachwinforms
{
    public partial class TextConverter : Form
    {

        private string path;
        private string fileText;
        private string output;
        private string extension;
              
        public TextConverter()
        {
            InitializeComponent();
            openFileDialog1.Title = "Выбрать файл";
            openFileDialog1.Filter = "Text files(*.txt)|*.txt|Text files(*.pdf)|*.pdf|Text files(*.doc)|*.doc|Text files(*.docx)|*.docx|Text files(*.rtf)|*.rtf|Text files(*.html)|*.html";
            saveFileDialog1.Title = "Сохранить файл";           
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            return;
            path = openFileDialog1.FileName;
            fileText = System.IO.File.ReadAllText(path);
            label5.Text = path;
            extension =path.Substring(path.LastIndexOf('.'));        
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            switch (comboBox1.Text)
            {
                case "RTF":

                    switch (extension)
                    {
                        case ".rtf":
                            saveFileDialog1.Filter = "Файлы формата rtf|*.rtf";
                            saveFileDialog1.DefaultExt = "rtf";
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                            return;
                            output = saveFileDialog1.FileName;
                            System.IO.File.Copy(path, output+".rtf", true);
                            break;
                        case ".doc":
                            saveFileDialog1.Filter = "Text files(*.rtf)|*.rtf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.UseOffice docrtf = new SautinSoft.UseOffice();
                            int dc = docrtf.InitWord();
                            dc = docrtf.ConvertFile(@path, @output+".rtf", SautinSoft.UseOffice.eDirection.DOC_to_RTF);
                            docrtf.CloseWord();
                            break;
                        case ".docx":
                            saveFileDialog1.Filter = "Text files(*.rtf)|*.rtf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore docxrtf = DocumentCore.Load(@path);
                            docxrtf.Save(@output+".rtf");                           
                            break;
                        case ".txt":
                            saveFileDialog1.Filter = "Text files(*.rtf)|*.rtf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.UseOffice txtrtf = new SautinSoft.UseOffice();
                            int tx = txtrtf.InitWord();
                            tx = txtrtf.ConvertFile(@path, @output + ".rtf", SautinSoft.UseOffice.eDirection.TEXT_to_RTF);
                            txtrtf.CloseWord();
                            break;
                        case ".pdf":
                            saveFileDialog1.Filter = "Text files(*.rtf)|*.rtf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore pdfrtf = DocumentCore.Load(@path);
                            pdfrtf.Save(@output+".rtf");                         
                            break;
                        case ".html":
                            saveFileDialog1.Filter = "Text files(*.rtf)|*.rtf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.UseOffice htmlrtf = new SautinSoft.UseOffice();
                            int rtf = htmlrtf.InitWord();
                            rtf = htmlrtf.ConvertFile(@path, @output + ".rtf", SautinSoft.UseOffice.eDirection.HTML_to_RTF);
                            htmlrtf.CloseWord();
                            break;
                    }
                    break;
                case "DOC":
                    switch (extension)
                    {
                        case ".rtf":
                            saveFileDialog1.Filter = "Text files(*.doc)|*.doc";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.UseOffice rtfdoc = new SautinSoft.UseOffice();
                            int rt = rtfdoc.InitWord();
                            rt = rtfdoc.ConvertFile(@path, @output+".doc", SautinSoft.UseOffice.eDirection.RTF_to_DOC);
                            rtfdoc.CloseWord();
                            break;
                        case ".doc":
                            saveFileDialog1.Filter = "Text files(*.doc)|*.doc";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            System.IO.File.Copy(path, output+".doc", true);                           
                            break;
                        case ".docx":
                            saveFileDialog1.Filter = "Text files(*.doc)|*.doc";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            var docxdoc = new Aspose.Words.Document(@path);
                            docxdoc.Save(output+".doc", Aspose.Words.SaveFormat.Doc);
                            break;
                        case ".txt":
                            saveFileDialog1.Filter = "Text files(*.doc)|*.doc";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            System.IO.File.WriteAllText(output+".doc", fileText);                           
                            break;
                        case ".pdf":
                            saveFileDialog1.Filter = "Text files(*.doc)|*.doc";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            PdfDocument pdfdoc = new PdfDocument();
                            pdfdoc.LoadFromFile(path);
                            pdfdoc.SaveToFile(output+".doc", Spire.Pdf.FileFormat.DOC);                           
                            break;
                        case ".html":
                            saveFileDialog1.Filter = "Text files(*.doc)|*.doc";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.UseOffice htmldoc = new SautinSoft.UseOffice();
                            int ht = htmldoc.InitWord();
                            ht = htmldoc.ConvertFile(@path, @output+".doc", SautinSoft.UseOffice.eDirection.HTML_to_DOC);
                            htmldoc.CloseWord();
                            break;
                    }
                    break;
                case "PDF":
                    switch (extension)
                    {
                        case ".rtf":
                            saveFileDialog1.Filter = "Text files(*.pdf)|*.pdf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.UseOffice u = new SautinSoft.UseOffice();
                            int ret = u.InitWord();
                            ret = u.ConvertFile(@path, @output+".pdf", SautinSoft.UseOffice.eDirection.RTF_to_PDF);
                            u.CloseWord();
                            break;
                        case ".doc":
                            saveFileDialog1.Filter = "Text files(*.pdf)|*.pdf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                            Microsoft.Office.Interop.Word.Document file = word.Documents.Open(path);
                            file.ExportAsFixedFormat(output+".pdf", WdExportFormat.wdExportFormatPDF);
                            word.Quit();                           
                            break;
                        case ".docx":
                            saveFileDialog1.Filter = "Text files(*.pdf)|*.pdf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore docxpdf = DocumentCore.Load(@path);
                            docxpdf.Save(@output+".pdf");                            
                            break;
                        case ".txt":
                            saveFileDialog1.Filter = "Text files(*.pdf)|*.pdf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore txtpdf = DocumentCore.Load(@path);
                            txtpdf.Save(@output+".pdf");                           
                            break;
                        case ".pdf":
                            saveFileDialog1.Filter = "Text files(*.pdf)|*.pdf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            System.IO.File.Copy(path, output+".pdf", true);                           
                            break;
                        case ".html":
                            saveFileDialog1.Filter = "Text files(*.pdf)|*.pdf";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore htmlpdf = DocumentCore.Load(@path);
                            htmlpdf.Save(@output+".pdf");                           
                            break;
                    }
                    break;
                case "DOCX":
                    switch (extension)
                    {
                        case ".rtf":
                            saveFileDialog1.Filter = "Text files(*.docx)|*.docx";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore rtfdocx = DocumentCore.Load(@path);
                            rtfdocx.Save(@output+".docx");                          
                            break;
                        case ".doc":
                            saveFileDialog1.Filter = "Text files(*.docx)|*.docx";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.UseOffice u = new SautinSoft.UseOffice();
                            int ret = u.InitWord();
                            ret = u.ConvertFile(@path, @output+".docx", SautinSoft.UseOffice.eDirection.DOC_to_DOCX);
                            u.CloseWord();
                            break;
                        case ".docx":
                            saveFileDialog1.Filter = "Text files(*.docx)|*.docx";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            System.IO.File.Copy(path, output+".docx", true);                           
                            break;
                        case ".txt":
                            saveFileDialog1.Filter = "Text files(*.docx)|*.docx";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore txtdocx = DocumentCore.Load(@path);
                            txtdocx.Save(@output+".docx");                           
                            break;
                        case ".pdf":
                            saveFileDialog1.Filter = "Text files(*.docx)|*.docx";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore pdfdocx = DocumentCore.Load(@path);
                            pdfdocx.Save(@output+".docx");                           
                            break;
                        case ".html":
                            saveFileDialog1.Filter = "Text files(*.docx)|*.docx";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore htmldocx = DocumentCore.Load(@path);
                            htmldocx.Save(@output+".docx");                           
                            break;
                    }
                    break;
                case "TXT":
                    switch (extension)
                    {
                        case ".rtf":
                            saveFileDialog1.Filter = "Text files(*.txt)|*.txt";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore rtftxt = DocumentCore.Load(@path);
                            rtftxt.Save(@output+".txt");                           
                            break;
                        case ".doc":
                            saveFileDialog1.Filter = "Text files(*.txt)|*.txt";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.UseOffice u = new SautinSoft.UseOffice();
                            int ret = u.InitWord();
                            ret = u.ConvertFile(@path, @output+".txt", SautinSoft.UseOffice.eDirection.DOC_to_TEXT);
                            u.CloseWord();
                            break;
                        case ".docx":
                            saveFileDialog1.Filter = "Text files(*.txt)|*.txt";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore docxtxt = DocumentCore.Load(@path);
                            docxtxt.Save(@output+".txt");                          
                            break;
                        case ".txt":
                            saveFileDialog1.Filter = "Text files(*.txt)|*.txt";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            System.IO.File.Copy(path, output+".txt", true);                           
                            break;
                        case ".pdf":
                            saveFileDialog1.Filter = "Text files(*.txt)|*.txt";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore pdftxt = DocumentCore.Load(@path);
                            pdftxt.Save(@output+".txt");                           
                            break;
                        case ".html":
                            saveFileDialog1.Filter = "Text files(*.txt)|*.txt";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore htmltxt = DocumentCore.Load(@path);
                            htmltxt.Save(@output+".txt");                           
                            break;
                    }
                    break;
                case "HTML":
                    switch (extension)
                    {
                        case ".rtf":
                            saveFileDialog1.Filter = "Text files(*.html)|*.html";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.RtfToHtml rtfhtml = new SautinSoft.RtfToHtml();
                            rtfhtml.OpenRtf(path);                           
                            rtfhtml.ToHtml(output+".html");                          
                            break;
                        case ".doc":
                            saveFileDialog1.Filter = "Text files(*.html)|*.html";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.UseOffice u = new SautinSoft.UseOffice();
                            int ret = u.InitWord();
                            ret = u.ConvertFile(path, output+".html", SautinSoft.UseOffice.eDirection.DOC_to_HTML);
                            u.CloseWord();                           
                            break;
                        case ".docx":
                            saveFileDialog1.Filter = "Text files(*.html)|*.html";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore docxhtml = DocumentCore.Load(@path);
                            docxhtml.Save(@output+".html");                           
                            break;
                        case ".txt":
                            saveFileDialog1.Filter = "Text files(*.html)|*.html";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            SautinSoft.RtfToHtml txthtml = new SautinSoft.RtfToHtml();
                            txthtml.OutputFormat = RtfToHtml.eOutputFormat.HTML_5;
                            txthtml.ConvertFile(path, output+".html");                            
                            break;
                        case ".pdf":
                            saveFileDialog1.Filter = "Text files(*.html)|*.html";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            DocumentCore pdfhtml = DocumentCore.Load(@path);
                            pdfhtml.Save(@output+".html");                           
                            break;
                        case ".html":
                            saveFileDialog1.Filter = "Text files(*.html)|*.html";
                            saveFileDialog1.RestoreDirectory = true;
                            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                                return;
                            output = saveFileDialog1.FileName;
                            System.IO.File.Copy(path, output+".html", true);                           
                            break;
                    }
                    break;
            }
            this.timer1.Start();

            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {
           
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            
            this.progressBar1.Increment(1);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();
            newForm.Show();
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void panel14_Paint(object sender, PaintEventArgs e)
        {
            
        }
    }
}
