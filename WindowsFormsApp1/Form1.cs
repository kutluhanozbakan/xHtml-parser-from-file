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
using System.Xml;
using HtmlAgilityPack;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        List<String> attributeNames = new List<String> { };
        List<String> valueNames = new List<String> { };
        List<String> nodeNames = new List<String> { };
        List<String> comingDiv = new List<String> { };
        List<String> attributeInnerHTMLNames = new List<String> { };
        int attributeLength = 0;
        int x, y, z = 0;
        bool close = false;
        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;
        public Form1()
        {
            InitializeComponent();
            if (xlApp == null)
            {
                MessageBox.Show("Excel kurulamadı.");
                return;
            }
            CreateExcelFile();
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
                    var path = ofd.FileName;
                    
                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.Load(path,Encoding.UTF8);
                    List<HtmlNode> nodes = doc.DocumentNode.ChildNodes.ToList();
                    getNodesInformation(nodes);
                    close = true;

                    SaveOnExcel(attributeNames, valueNames, nodeNames, attributeInnerHTMLNames, x, y, z, close);

                }
            }
        }

        private void getNodesInformation(List<HtmlNode> nodes)
        {
           foreach(var node in nodes)
            {
                if(node.HasChildNodes || node.HasAttributes)
                {
                    Console.WriteLine("Node Name:" + node.Name);
                    nodeNames.Add(node.Name);
                }          
               if(node.HasAttributes)
                {
                    
                    foreach(var attributes in node.Attributes)
                    {
                        if(attributes.Name != "class")
                        {
                            if(attributes.OwnerNode.InnerText != "" && !attributes.OwnerNode.InnerText.Contains("\r"))
                            {
                                attributeInnerHTMLNames.Add(attributes.OwnerNode.InnerText);

                            }
                            else
                            {
                                attributeInnerHTMLNames.Add("");
                            }
                            attributeLength = node.Attributes.Count;
                            Console.WriteLine("Attribute:" + attributes.Name);
                            Console.WriteLine("Value:" + attributes.Value);
                            attributeNames.Add(attributes.Name);
                            valueNames.Add(attributes.Value);

                        }
                    }
                   
                    Console.WriteLine(attributeLength);
                    
                   SaveOnExcel(attributeNames, valueNames, nodeNames, attributeInnerHTMLNames, x, y, z, close);
                  
                }
                if(node.HasChildNodes)
                {
                    List<HtmlNode> childNode = node.ChildNodes.ToList();
                  
                    getNodesInformation(childNode);
                }
             
            }
            attributeNames.Add("");
            valueNames.Add("");
            attributeInnerHTMLNames.Add("");
        }
        //EXCEL ÜZERİNE KAYIT ETME
        private void SaveOnExcel(List<String> attiributes, List<String> values, List<String> nodeNames, List<string> attributeInnerHTMLNames, int x, int y, int z, bool close)
        {
            for (int i = 1; i <= attributeInnerHTMLNames.Count; i++)
            {
                xlWorkSheet.Cells[i + 1, 4] = attributeInnerHTMLNames[i - 1];
            }
            for (int i = 1; i <= attiributes.Count; i++)
            {
                xlWorkSheet.Cells[i+1, 2] = attiributes[i-1];             
            }
            for (int i = 1; i <= values.Count; i++)
            {
                xlWorkSheet.Cells[i + 1, 3] = values[i - 1];                
            }
            if (close)
            {
                closeExcel(xlWorkBook, xlWorkSheet, misValue, xlApp);
            }
            
            
           
        }
        private void CreateExcelFile()
        {
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "Node Name";
            xlWorkSheet.Cells[1, 2] = "Names";
            xlWorkSheet.Cells[1, 3] = "Values";
            xlWorkBook.SaveAs("C:\\Users\\kutluhan.ozbakan\\Documents\\ARD\\csharp-Excel.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        }
        private void closeExcel(Microsoft.Office.Interop.Excel.Workbook xlWorkBook, Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet, object misValue, Microsoft.Office.Interop.Excel.Application xlApp)
        {
            //EXCEL DOSYASININ YAZILACAGI YERİ ŞİMDİLİK ELLE VERİYORUM.
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            MessageBox.Show("Excel dosyası oluşturuldu!");
            System.Windows.Forms.Application.Exit();
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}