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

        private string Code_Gimla_Groups = string.Empty;
        private string Code_Malal_Groups = string.Empty;
        public Form1()
        {
            InitializeComponent();

            lblOf.Visible = false;
            lblRow.Visible = false;
            label3.Visible = false;
            label5.Visible = false;

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

            Code_Gimla_Groups = ConfigurationManager.AppSettings.Get("Code_Gimla_Groups");
            Code_Malal_Groups = ConfigurationManager.AppSettings.Get("Code_Malal_Groups");

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
        private List<ExcelVals> ReadExcel(string excelPath)
        {
            var xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = null;
            Excel._Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;
            try
            {
                xlWorkbook = xlApp.Workbooks.Open(excelPath);
                xlApp.Visible = false;
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
                var lastRow = xlWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                var lst = new List<ExcelVals>();
                for (var i = 2; i <= lastRow; i++)
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
                    var id = xlWorksheet.Range["E" + i, "E" + i].Value2.ToString();
                    var tickDate = xlWorksheet.Range["F" + i, "F" + i].Value2?.ToString();
                    var code = xlWorksheet.Range["G" + i, "G" + i].Value2?.ToString();
                    var newType = xlWorksheet.Range["H" + i, "H" + i].Value2?.ToString();


                    lst.Add(new ExcelVals
                    {
                        clientFile = clientFile,
                        letter = letter,
                        OrderCompanyID = OrderCompanyID,
                        FirstName = FirstName,
                        LastName = LastName,
                        id = id,
                        tickDate = tickDate,
                        code = code,
                        newType = newType,
                        line = i
                    });
                }
                return lst;
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Log(ex);
                return new List<ExcelVals>();
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
        private void SetExcelError(int line, string msg, string excelPath)
        {
            var xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = null;
            Excel._Worksheet xlWorksheet = null;
            try
            {
                xlWorkbook = xlApp.Workbooks.Open(excelPath);
                xlApp.Visible = false;
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
                xlWorksheet.Range["I" + line, "I" + line].Value = msg;

                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                if (xlWorksheet != null)
                    Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                if (xlWorksheet != null)
                {
                    
                    Marshal.ReleaseComObject(xlWorkbook);
                }

                if (xlApp != null)
                {
                    //quit and release
                    
                    Marshal.ReleaseComObject(xlApp);
                }
            }
        }
        private void ParseExcel(string archivePath, string excelPath, string dirName)
        {
            var list = new List<string>();
            var dels = new List<string>();
            var Gimlas = Code_Gimla_Groups.Split(',');
            var Malals = Code_Malal_Groups.Split(',');

            lblMsg.Text = "שומר קודים ביטוח לאומי";
            System.Windows.Forms.Application.DoEvents();
            var codes = GetCodes();
            if(codes.Count() == 0)
            {
                throw new Exception("קודים של ביטוח לאומי לא נקלטו כראוי");
            }
            var rowCounter = 0;
            string vsrname = string.Empty;
            string id = string.Empty;
            int line = -1;
            int counter = 0;
            try
            {
                lblMsg.Text = "בודק שורות באקסל. זה יכול לקחת כמה רגעים";
                System.Windows.Forms.Application.DoEvents();

                var lst = ReadExcel(excelPath);
                if(lst.Count() == 0)
                {
                    lblMsg.Text = "בדיקת האקסל נכשלה - בדוק קובץ לוג";
                    System.Windows.Forms.Application.DoEvents();
                    return;
                }
                lblMsg.Text = "בדיקת האקסל עברה בהצלחה";
                System.Windows.Forms.Application.DoEvents();
                var query = lst.GroupBy(x => new { x.clientFile, x.code, x.newType })
                    .Select(g => g.OrderByDescending(c => c.tickDate))
                    .ToList();

                
                var endList = new List<ExcelVals>();
                foreach (var item in query)
                {
                    var ilst = item.ToList();
                    var gmla = ilst[0].code;
                    var mll = ilst[0].newType;
                    if(Gimlas.Contains(gmla) && Malals.Contains(mll))
                    {
                        var itm = ilst[0];
                        ilst = new List<ExcelVals>();
                        ilst.Add(itm);
                    }
                    endList.AddRange(ilst);
                }
                lblMsg.Visible = false;
                lblOf.Visible = true;
                lblRow.Visible = true;
                label3.Visible = true;
                label5.Visible = true;
                System.Windows.Forms.Application.DoEvents();
                var lastRow = endList.Count();
                lblOf.Text = (lastRow - 1).ToString();
                lblRow.Text = "0";
                System.Windows.Forms.Application.DoEvents();
                foreach (var it in endList)
                {
                    try
                    {
                        counter++;
                        var OrderCompanyID = it.OrderCompanyID;
                        id = it.id;
                        line = it.line;
                        var code = it.code;
                        var tickDate = it.tickDate;
                        var newType = it.newType;
                        var FirstName = it.FirstName;
                        var LastName = it.LastName;


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
                            lblRow.Text = (line - 1).ToString();
                            System.Windows.Forms.Application.DoEvents();
                            if (!File.Exists(vsrpath))
                            {
                                throw (new Exception("קובץ וסר לא קיים" + " --- " + "שורה" + "  " + line));
                            }
                            if (!File.Exists(yippath))
                            {
                                throw (new Exception("קובץ יפוי כח לא קיים" + " --- " + "שורה" + "  " + line));
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
                            if (newType == "201")
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
                        lblRow.Text = counter.ToString();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    catch (Exception ex)
                    {
                        SimpleLogger.SimpleLog.Log(ex);
                        SetExcelError(line, ex.Message, excelPath);
                    }


                }

                lblMsg.Visible = true;
                lblOf.Visible = false;
                lblRow.Visible = false;
                label3.Visible = false;
                label5.Visible = false;
                lblMsg.Text = "מוחק קבצים";
                System.Windows.Forms.Application.DoEvents();
                foreach (var item in dels)
                {
                    if (File.Exists(item))
                    {
                        File.Delete(item);
                    }
                }

                lblMsg.Text = "מעתיק קבצים לתקיית יעד";
                System.Windows.Forms.Application.DoEvents();
                var dest = Path.Combine(txtDest.Text, dirName);
                Directory.CreateDirectory(dest);
                string efile = string.Empty; ;
                try
                {
                    foreach (var file in list)
                    {
                        efile = Path.GetFileName(file);
                        if (File.Exists(file) &&  file.Contains("YIP"))
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
                }
                catch (Exception ex)
                {
                    SimpleLogger.SimpleLog.Log(ex);
                    SimpleLogger.SimpleLog.Log("נכשל ביצירת קובץ בשם - " + efile);
                    lblMsgPdfError.Text = "נכשל ביצירת קובץ בשם - " + efile;
                    System.Windows.Forms.Application.DoEvents();

                }
                
                MessageBox.Show("הפעולה הסתיימה בהצלחה. " + rowCounter + " מתוך " + lastRow + " נקלטו ");
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Log(ex);
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
    public class ExcelVals
    {
        public string clientFile { get; set; }
        public string letter { get; set; }
        public string OrderCompanyID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string id { get; set; }
        public string tickDate { get; set; }
        public string code { get; set; }
        public string newType { get; set; }
        public int line { get; set; }
    }
}
