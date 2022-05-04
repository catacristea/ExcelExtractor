using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Windows.Forms;
using _Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelExtractor
{


    public partial class Form1 : Form
    {
        private OpenFileDialog ofd;
        private FolderBrowserDialog fbd;

        string excelPath = string.Empty;
        string locationPath = string.Empty;
        int counter = 1;
        int totalTextCopied = 1;

        public static void WordApp(string text, string locationPath, string excelName, int counter)
        {
            //app
            //Word.Application objWord = new Word.Application();
            //objWord.Visible = true;
            //objWord.WindowState = Word.WdWindowState.wdWindowStateNormal;

            //doc
            Word.Document doc = new Word.Document();

            //paragraf

            Word.Paragraph paragraph = doc.Paragraphs.Add();
            paragraph.Range.Text = text;

            doc.SaveAs2(locationPath + $"\\Text {counter} - {excelName}.docx");
            doc.Close();
            //objWord.Quit();

        }


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ofd = new OpenFileDialog();
            fbd = new FolderBrowserDialog();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                excelPath = ofd.FileName;
                lblRezult.Text = "Documentul " + Path.GetFileName(excelPath) + " a fost încărcat cu succes!";
            }
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                locationPath = fbd.SelectedPath;
                lblFolder.Text = "Directorul a fost selectat cu succes!";
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {

            try
            {
                message.Text = $"Se porecesează!";

                int cellCount = 0;
                int columnCount = 0;
                _Application excel = new _Excel.Application();
                _Excel.Workbook wb = excel.Workbooks.Open(excelPath);
                int sheetNumber = wb.Sheets.Count;

                string text;
                for (int i = 1; i <= sheetNumber; i++)
                {
                    _Excel.Worksheet ws = excel.Worksheets[i];
                    int columnNumber = ws.Columns.Count;
                    int rowNumber = ws.Rows.Count;
                    columnCount = 0;

                    for (int j = 1; j < columnNumber; j++)
                    {
                        cellCount = 0;
                        if (columnCount > 15)
                            break;

                        for (int k = 1; k < rowNumber; k++)
                        {
                            if (ws.Cells[j, k].Value2 != null)
                            {
                                text = (ws.Cells[j, k].Value2).ToString();
                                WordApp(text, locationPath, Path.GetFileNameWithoutExtension(excelPath), counter++);
                                message.Text = $"Se porecesează: \nAu fost create {totalTextCopied} documente Word!\nVă rugăm așteptați!";
                                totalTextCopied++;
                            }
                            else
                            {
                                cellCount++;
                                if (cellCount > 15)
                                {
                                    columnCount++;
                                    break;
                                }
                            }
                        }
                    }

                }
            message.Text = $"Proces finalizat! :)";
            MessageBox.Show($"Au fost copiate {counter - 1} texte");
            wb.Close();
            excel.Quit();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("Vă rugăm alegeți un document excel și o locație în care veți salva fișierele și reîncercați!");
                message.Text = "";
                lblRezult.Text = "";
                lblFolder.Text = "";
                locationPath = "";
            }

            catch (Exception)
            {
                MessageBox.Show("A apărut o eroare! Asigurați-vă că documentul pe care l-ați încărcat este de tip Excel (.xlsx sau .xls) și reîncercați!");
                message.Text = "";
                lblRezult.Text = "";
                lblFolder.Text = "";
                locationPath = "";
            }
        }

        private void anulare_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }
    }
}
