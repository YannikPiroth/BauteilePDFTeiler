using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;
using iTextSharp.text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Globalization;
using Microsoft.Win32;
using System.Windows.Documents;

namespace BauteilePDFTeiler
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string filePath = "";
        string outputPath = "";
        public MainWindow()
        {
            InitializeComponent();
        }
        private void dateiAnzeige_Drop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {

                string[] files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);

                foreach (string file in files)
                {
                    filePath = file;
                }
                dateiAnzeige.Content = filePath;
            }
        }

        private void dateiButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "PDF (.pdf)|*.pdf"; // Filter files by extension

            // Show open file dialog box
            bool? result = dialog.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                filePath = dialog.FileName;
                dateiAnzeige.Content = filePath;
            }
        }

        private void ausführenButton_Click(object sender, RoutedEventArgs e)
        {
            if (filePath != "")
            {
                List<string> pdfInhalt = GetTextFromPDF();
                foreach(Bauteil bauteil in getBauteile(pdfInhalt))
                {
                    SplitAndSaveInterval(filePath, bauteil.Seite, 0, bauteil.Name);
                }
            }
            else
            {
                System.Windows.MessageBox.Show("Es muss eine Datei ausgewählt sein", "Fehler - Keine Datei",
                MessageBoxButton.OKCancel);
            }
        }

        private List<string> GetTextFromPDF()
        {
            StringBuilder text = new StringBuilder();
            var pdfInhalt = new List<string>();
            using (PdfReader reader = new PdfReader(filePath))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                    pdfInhalt.Add(PdfTextExtractor.GetTextFromPage(reader, i));
                }
            }

            return pdfInhalt;
        }

        private void SplitAndSaveInterval(string pdfFilePath, int startPage, int interval, string pdfFileName)
        {

            string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            outputPath = strExeFilePath.Substring(0, strExeFilePath.Length - 21) + DateTime.Now.ToString("yyyy-MM-dd_HH-mm") + "\\";
            System.IO.Directory.CreateDirectory(outputPath);


            using (PdfReader reader = new PdfReader(pdfFilePath))
            {
                Document document = new Document();
                if(!File.Exists(outputPath + "\\" + pdfFileName + ".pdf"))
                {
                    PdfCopy copy = new PdfCopy(document, new FileStream(outputPath + "\\" + pdfFileName + ".pdf", FileMode.Create));
                    document.Open();

                    for (int pagenumber = startPage; pagenumber < (startPage + interval + 1); pagenumber++)
                    {
                        if (reader.NumberOfPages >= pagenumber)
                        {
                            copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                        }
                        else
                        {
                            break;
                        }

                    }

                    document.Close();
                } else
                {
                    int i = 1;
                    do
                    {
                        i++;
                    } while (File.Exists(outputPath + "\\" + pdfFileName + i + ".pdf"));

                    PdfCopy copy = new PdfCopy(document, new FileStream(outputPath + "\\" + pdfFileName + "(" + i + ")" + ".pdf", FileMode.Create));
                    document.Open();

                    for (int pagenumber = startPage; pagenumber < (startPage + interval + 1); pagenumber++)
                    {
                        if (reader.NumberOfPages >= pagenumber)
                        {
                            copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                        }
                        else
                        {
                            break;
                        }

                    }

                    document.Close();
                }

            }
        }

        private List<Bauteil> getBauteile(List<string> pdfInhalt) 
        {
            List<Bauteil> bauteileList = new List<Bauteil>();
            Bauteil bauteil = new Bauteil();
            int seitenZahl = 1;
            foreach(string pdf in pdfInhalt) 
            {
                if(seitenZahl == 50)
                {
                    string test = "";
                }
                bauteil = new Bauteil();
                string[] seite = pdf.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                bauteil.Seite = seitenZahl;
                seitenZahl++;
                bauteil.Name = GetName(pdf);

                bauteileList.Add(bauteil);
            }

            return bauteileList;
        }


        private string GetName(string pString)
        {

            string name = "";

            string input = pString;
            string searchTerm = "Freigabe zur Fertigung";
            string[] content = pString.Split(new string[] { "\r\n", "\r", "\n" },
                StringSplitOptions.None);
            string last = "";
            
            foreach(string s in content)
            {
                if(s.Contains(searchTerm, StringComparison.OrdinalIgnoreCase))
                {
                    name = last;
                    break;
                }
                last = s;
            }



            return name;
        }
    }
}
