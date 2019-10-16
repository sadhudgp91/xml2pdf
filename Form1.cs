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
using iText;
using iTextSharp;
using iTextSharp.text.pdf;
using iTextSharp.text;
using PdfSharp.Pdf;

namespace xml2pdf
{
    public partial class Form1 : Form
    {
        DataSet ds = new DataSet();
        public Form1()
        {
            InitializeComponent();
        }

        private void Browse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Browse XML Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xml",
                Filter = "txt files (*.xml)|*.xml",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                toolStripStatusLabel1.Text = openFileDialog1.FileName;
                statusStrip1.BackColor = Color.Green;
            }

            ds.ReadXml(openFileDialog1.FileName);
            dataGridView1.DataSource = ds.Tables[0];
        }

        private void FncPdf_Click(object sender, EventArgs e)
        {
            
            // Converting Dataset to PDF
            Document document = new Document();

            string fn = "zugferd-invoice" + "_" + DateTime.Now.ToShortDateString();
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = fn.Replace("/", "-").Replace(" ", "_");
            sfd.Filter = "(*.pdf)|*.pdf";
            sfd.ShowDialog();
            string path = sfd.FileName;

            if (sfd.FileName != null)
            {
                //Exporting to PDF.
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(path, FileMode.Create));
                toolStripStatusLabel1.Text = "PDF File Generated Successfully in: " + path;
            }

            document.Open();
            for (int table = 0; table < ds.Tables.Count; table++)
            {
                PdfPTable pdfTable = new PdfPTable(ds.Tables[table].Columns.Count);
                //pdfTable.AddCell(new Phrase("heading1"));
                //pdfTable.AddCell(new Phrase("heading2"));
                //pdfTable.AddCell(new Phrase("heading3"));

                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    pdfTable.AddCell(new Phrase(dataGridView1.Columns[i - 1].HeaderText));
                }


                for (int row = 0; row < ds.Tables[table].Rows.Count; row++)
                {

                    for (int column = 0; column < ds.Tables[table].Columns.Count; column++)
                    {
                        pdfTable.AddCell(new Phrase(ds.Tables[table].Rows[row][column].ToString()));
                    }
                }
                document.Add(pdfTable);
            }
            document.Close();
           
            GC.Collect();
            GC.WaitForPendingFinalizers();

            toolStripStatusLabel1.Text = "PDF File Generated Successfully in: " + path;
            statusStrip1.BackColor = Color.Green;

        }

    }
}
