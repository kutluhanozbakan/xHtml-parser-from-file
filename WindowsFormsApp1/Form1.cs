using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        List<String> comingDiv = new List<String> { };
        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog()
            {
                Multiselect = false,
                ValidateNames = true,
                Filter = "XHTML|*.xhtml"
            })
            {              
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    SaveFile(ofd);
                    var path = @"kutluhanAgility.txt";
                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.Load(path);
                    foreach (HtmlAgilityPack.HtmlNode node in
                     doc.DocumentNode.SelectNodes("//div[@class='form-group samerow']"))
                    {
                        comingDiv.Insert(0, node.InnerHtml);
                    }
                    SaveOnExcel(comingDiv);
                }

            }
        }
        //EXCEL ÜZERİNE KAYIT ETME
        private void SaveOnExcel(List<String> data)
        {   
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel kurulamadı.");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "Div";

            for (int i = 1; i <= data.Count; i++)
            {
                xlWorkSheet.Cells[i+1, 1] = data[i-1];
            }
            //EXCEL DOSYASININ YAZILACAGI YERİ ŞİMDİLİK ELLE VERİYORUM.
            xlWorkBook.SaveAs("C:\\Users\\kutluhan.ozbakan\\Documents\\ARD\\csharp-Excel.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            MessageBox.Show("Excel dosyası oluşturuldu!");
        }
        private static void SaveFile(OpenFileDialog ofd)
        {
            List<string> codeList = new List<string>();
            string line = "";
            string[] codeArray;
            string[] codeFromFile = System.IO.File.ReadAllLines(ofd.FileName);          
            foreach(String lineCode in codeFromFile)
            {
                line = Regex.Replace(lineCode, @"^\s*$\n|\r", string.Empty, RegexOptions.Multiline).TrimEnd().TrimStart();
                codeList.Add(line);
                codeArray = codeList.ToArray();
                //TXT DOSYASININ YÜKLENECEGİ YER ŞİMDİLİK ELLE VERİYORUM.
                System.IO.File.WriteAllLines(@"C:\\Users\\kutluhan.ozbakan\\Documents\\ARD\\appended.txt", codeArray);
            }
            //TXT DOSYASININ OKUNACAĞI YER ŞİMDİLİK ELLE VERİYORUM.
            string codeFromEditedFile = System.IO.File.ReadAllText(@"C:\\Users\\kutluhan.ozbakan\\Documents\\ARD\\appended.txt");
            var htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(codeFromEditedFile);
            htmlDoc.Save("kutluhanAgility.txt");
        }      
    }
}