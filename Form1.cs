using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using Spire.Pdf;

namespace XmlCreator
{
    public partial class Form1 : Form
    {
        enum NameType
        {
            XML,
            PDF_VSR,
            PD_YIP
        }
        private string Code_Vsr = string.Empty;
        private string Code2_Vsr = string.Empty;
        private string Code_Yip = string.Empty;
        private string BLL_CODES = string.Empty;
        public Form1()
        {
            InitializeComponent();
            
            var dirmPath = Properties.Settings.Default.MainPath;
            if (!string.IsNullOrEmpty(dirmPath))
            {
                txtMain.Text = dirmPath;
            }
            var dirtPath = Properties.Settings.Default.DestPath;
            if (!string.IsNullOrEmpty(dirtPath))
            {
                txtDest.Text = dirtPath;
            }
            var diraPath = Properties.Settings.Default.ArchivePath;
            if (!string.IsNullOrEmpty(diraPath))
            {
                txtArchive.Text = diraPath;
            }
            Code_Vsr = ConfigurationManager.AppSettings.Get("Code_Vsr");
            Code2_Vsr = ConfigurationManager.AppSettings.Get("Code2_Vsr");
            Code_Yip = ConfigurationManager.AppSettings.Get("Code_Yip");
            BLL_CODES = ConfigurationManager.AppSettings.Get("BllCodes");


        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var fbd = new OpenFileDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.FileName))
                {
                    txtMain.Text = fbd.FileName;
                    Properties.Settings.Default.MainPath = fbd.FileName;
                    Properties.Settings.Default.Save();
                }
            }
        }
        private List<Codes> GetCodes()
        {
            var lst = new List<Codes>();
            

            if (string.IsNullOrEmpty(BLL_CODES))
            {
                throw new Exception("קודים של ביטוח לאומי לא הוכנסו כראוי");
            }
            if (!File.Exists(BLL_CODES))
            {
                throw new Exception("קובץ קודים ביטוח לאומי לא נמצא ");
            }
            var xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = null;
            Excel._Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;
            try
            {
                xlWorkbook = xlApp.Workbooks.Open(BLL_CODES);
                xlApp.Visible = false;
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
                var lastRow = xlWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                for (int i = 2; i <= lastRow; i++)
                {
                    var gid = xlWorksheet.Range["A" + i, "A" + i].Value2.ToString();
                    var gtype = xlWorksheet.Range["C" + i, "C" + i].Value2.ToString();
                    lst.Add(new Codes { id = gid, type = gtype });
                }
                
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Log(ex);
                
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                if (xlRange != null)
                    Marshal.ReleaseComObject(xlRange);
                if (xlWorksheet != null)
                    Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                if (xlWorksheet != null)
                {
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);
                }

                if (xlApp != null)
                {
                    //quit and release
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }

            }
            return lst;
        }
        private void btnStart_Click(object sender, EventArgs e)
        {
            var dirName = DateTime.Now.ToString("yyyyMMddHHmmss");
            var archivePath = Path.Combine(txtArchive.Text, dirName);
            var mainDir = Path.GetDirectoryName(txtMain.Text);
            var excelPath = Path.Combine(archivePath, Path.GetFileName(txtMain.Text));
            Directory.Move(mainDir, archivePath);
            Directory.CreateDirectory(mainDir);
            ParseExcel(archivePath, excelPath, dirName);
        }
        private void ParseExcel(string archivePath, string excelPath, string dirName)
        {
            var list = new List<string>();
            var dels = new List<string>();
            var codes = GetCodes();
            if(codes.Count() == 0)
            {
                throw new Exception("קודים של ביטוח לאומי לא נקלטו כראוי");
            }
            var rowCounter = 0;
            var xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = null;
            Excel._Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;
            int i;
            string vsrname = string.Empty;
            string id = string.Empty;
            try
            {

                xlWorkbook = xlApp.Workbooks.Open(excelPath);
                xlApp.Visible = false;
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
                var lastRow = xlWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                lblOf.Text = (lastRow - 1).ToString();
                lblRow.Text = "0";
                System.Windows.Forms.Application.DoEvents();
                for (i = 2; i <= lastRow; i++)
                {
                    try
                    {
                        if (xlWorksheet.Range["A" + i, "A" + i].Value2 == null)
                        {
                            continue;
                        }
                        //var XmlCounter = 1
                        //for (var XmlCounter = 1; XmlCounter <= 2; XmlCounter++)
                        //{
                        var clientFile = xlWorksheet.Range["A" + i, "A" + i].Value2.ToString();
                        var letter = xlWorksheet.Range["B" + i, "B" + i].Value2.ToString();
                        //var OrderCompanyID = clientFile + letter + XmlCounter;
                        var OrderCompanyID = clientFile + letter;
                        var FirstName = xlWorksheet.Range["C" + i, "C" + i].Value2.ToString();
                        var LastName = xlWorksheet.Range["D" + i, "D" + i].Value2.ToString();
                        id = xlWorksheet.Range["E" + i, "E" + i].Value2.ToString();
                        var tickDate = xlWorksheet.Range["F" + i, "F" + i].Value2?.ToString();
                        var code = xlWorksheet.Range["G" + i, "G" + i].Value2?.ToString();
                        var newType = xlWorksheet.Range["H" + i, "H" + i].Value2?.ToString();

                        vsrname = CreateName(OrderCompanyID, id, "pdf", NameType.PDF_VSR);
                        var yipname = CreateName(OrderCompanyID, id, "pdf", NameType.PD_YIP);
                        var xmlname = CreateName(OrderCompanyID, id, "xml", NameType.XML);

                        var typecode = codes.Find(f => f.type == code)?.id;
                        
                        //var type = XmlCounter == 1 ? "201" : "205";
                       // var type = "201";
                        string stickdate = "";
                        if (!string.IsNullOrEmpty(tickDate))
                        {
                            if (tickDate.Contains("1900"))
                            {
                                stickdate = "1900-01-01T00:00:00";
                            }
                            else
                            {
                                var dblTd = double.Parse(tickDate);
                                stickdate = DateTime.FromOADate(dblTd).ToString("yyyy-MM-ddTHH\\:mm\\:ss.mmm");
                            }
                            
                        }
                        


                        var pdvsr = Code_Vsr + "_" + id;
                        var pdvsr2 = Code2_Vsr + "_" + id;
                        var pdyip = id + "_" + Code_Yip;
                        var vsrpath = string.Empty;
                        var yippath = string.Empty;
                        var newVsrPath = string.Empty;
                        var newYipPath = string.Empty;

                        string[] files = Directory.GetFiles(archivePath);
                        foreach (var file in files)
                        {
                            var nm = Path.GetFileNameWithoutExtension(file);
                            if (nm.Contains(pdvsr) || nm.Contains(pdvsr2))
                            {
                                vsrpath = file;
                            }

                            var parts = nm.Split('_');
                            if (Array.Exists(parts, f => f == Code_Yip) && Array.Exists(parts, f => f == id))
                            {
                                yippath = file;
                            }
                        }
                        if (File.Exists(vsrpath) && File.Exists(yippath))
                        {
                            newVsrPath = Path.Combine(archivePath, vsrname);
                            newYipPath = Path.Combine(archivePath, yipname);
                            File.Copy(vsrpath, newVsrPath);
                            File.Copy(yippath, newYipPath);
                            //if (XmlCounter == 2)
                            //{
                            //    File.Delete(vsrpath);
                            //    File.Delete(yippath);
                            //}
                            //File.Delete(vsrpath);
                            //File.Delete(yippath);
                            dels.Add(vsrpath);
                            dels.Add(yippath);
                        }
                        else
                        {
                            lblRow.Text = (i - 1).ToString();
                            System.Windows.Forms.Application.DoEvents();
                            if (!File.Exists(vsrpath))
                            {
                                throw (new Exception("קובץ וסר לא קיים" + " --- " + "שורה" + "  " + i));
                            }
                            if (!File.Exists(yippath))
                            {
                                throw (new Exception("קובץ יפוי כח לא קיים" + " --- " + "שורה" + "  " + i));
                            }

                        }
                        id = id.PadLeft(9, '0');
                        var sts = new XmlWriterSettings()
                        {
                            Indent = true
                        };

                        using (var writer = XmlWriter.Create(Path.Combine(archivePath, xmlname), sts))
                        {
                            writer.WriteStartDocument();
                            writer.WriteStartElement("ActivityData");
                            writer.WriteStartElement("SPDataSetResults");
                            writer.WriteStartElement("SystemID");
                            writer.WriteString("2");
                            writer.WriteEndElement();
                            writer.WriteStartElement("SystemName");
                            writer.WriteString("חברות הביטוח");
                            writer.WriteEndElement();
                            writer.WriteStartElement("CompanyID");
                            writer.WriteString("5");
                            writer.WriteEndElement();
                            writer.WriteStartElement("CompanyName");
                            writer.WriteString("ערן מור");
                            writer.WriteEndElement();
                            writer.WriteStartElement("Interface");
                            writer.WriteString(newType);
                            writer.WriteEndElement();
                            writer.WriteStartElement("ServiceID");
                            writer.WriteString("10");
                            writer.WriteEndElement();
                            writer.WriteStartElement("OrderBtlID");
                            writer.WriteEndElement();
                            writer.WriteStartElement("OrderCompanyID");
                            writer.WriteString(OrderCompanyID);
                            writer.WriteEndElement();
                            writer.WriteStartElement("TimeStamp");
                            writer.WriteString(DateTime.Now.ToString("yyyy-MM-ddTHH\\:mm\\:ss.mmm"));
                            writer.WriteEndElement();
                            writer.WriteStartElement("Zehut");
                            writer.WriteString(id);
                            writer.WriteEndElement();
                            writer.WriteStartElement("FirstName");
                            writer.WriteString(FirstName);
                            writer.WriteEndElement();
                            writer.WriteStartElement("LastName");
                            writer.WriteString(LastName);
                            writer.WriteEndElement();
                            if (!string.IsNullOrEmpty(stickdate))
                            {
                                writer.WriteStartElement("TikDate");
                                writer.WriteString(stickdate);
                                writer.WriteEndElement();
                            }
                            if(newType == "201")
                            {
                                writer.WriteStartElement("KodGimla");
                                writer.WriteEndElement();
                            }
                            else
                            {
                                writer.WriteStartElement("KodGimla");
                                writer.WriteString(typecode);
                                writer.WriteEndElement();
                            }
                            writer.WriteStartElement("vasarDate");
                            writer.WriteString(DateTime.Now.ToString("yyyy-MM-ddT00\\:00\\:00"));
                            writer.WriteEndElement();
                            writer.WriteStartElement("attachmentFile");
                            writer.WriteStartElement("FileName");
                            writer.WriteString(vsrname);
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteStartElement("attachmentFile");
                            writer.WriteStartElement("FileName");
                            writer.WriteString(yipname);
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();

                        }


                        list.Add(newVsrPath);
                        list.Add(newYipPath);
                        list.Add(Path.Combine(archivePath, xmlname));


                        // }
                        rowCounter++;
                        lblRow.Text = (i - 1).ToString();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    catch (Exception ex)
                    {

                        SimpleLogger.SimpleLog.Log(ex);
                    }
                }

                foreach (var item in dels)
                {
                    if (File.Exists(item))
                    {
                        File.Delete(item);
                    }
                }
                //dfdfsdffadfa


                var dest = Path.Combine(txtDest.Text, dirName);
                Directory.CreateDirectory(dest);
                foreach (var file in list)
                {
                    if (file.Contains("YIP"))
                    {
                        using (PdfDocument sourceDoc = new PdfDocument(file))
                        {
                            using (PdfDocument newDoc = new PdfDocument())
                            {
                                newDoc.InsertPage(sourceDoc, sourceDoc.Pages.Count - 1);
                                newDoc.SaveToFile(Path.Combine(dest, Path.GetFileName(file)));
                            }
                        }
                            
                    }
                    else
                    {
                        File.Copy(file, Path.Combine(dest, Path.GetFileName(file)));
                    }
                    
                }
                MessageBox.Show("הפעולה הסתיימה בהצלחה. " + rowCounter + " מתוך " + (lastRow - 1) + " נקלטו ");
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Log(ex);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                if (xlRange != null)
                    Marshal.ReleaseComObject(xlRange);
                if (xlWorksheet != null)
                    Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                if (xlWorksheet != null)
                {
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);
                }

                if (xlApp != null)
                {
                    //quit and release
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }

            }
        }
        private string CreateName(string companyId, string id, string ext, NameType nameType)
        {
            var date = DateTime.Now.ToString("yyyyMMdd");
            return "Sm" + id + "-SM-" + date + "-" + (nameType==NameType.PDF_VSR ? "VSR-" : nameType == NameType.PD_YIP ? "YIP-" : "")  + companyId + "." + ext;
        }

        private void butDest_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    txtDest.Text = fbd.SelectedPath;
                    Properties.Settings.Default.DestPath = fbd.SelectedPath;
                    Properties.Settings.Default.Save();
                }
            }
        }

        private void btnArchive_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    txtArchive.Text = fbd.SelectedPath;
                    Properties.Settings.Default.ArchivePath = fbd.SelectedPath;
                    Properties.Settings.Default.Save();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
    public class Codes
    {
        public string id { get; set; }
        public string type { get; set; }
    }
}
