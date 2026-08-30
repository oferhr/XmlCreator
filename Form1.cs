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
using System.Drawing.Drawing2D;
using System.Xml.Linq;


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
        private string[] Gimlas;
        private string[] Malals;
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
            var dirwPath = Properties.Settings.Default.WorkPath;
            if (!string.IsNullOrEmpty(dirwPath))
            {
                txtWork.Text = dirwPath;
            }
            var diraPath = Properties.Settings.Default.ArchivePath;
            if (!string.IsNullOrEmpty(diraPath))
            {
                txtArchive.Text = diraPath;
            }
            var datemll = Properties.Settings.Default.DateMLL;
            if(datemll != DateTime.MinValue)
            {
                dateTimePicker1.Value = datemll;
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
            Directory.CreateDirectory(archivePath);
            Directory.Delete(txtWork.Text, true);
            //Directory.CreateDirectory(txtWork.Text);
            var mainDir = Path.GetDirectoryName(txtMain.Text);
            var excelPath = Path.Combine(txtWork.Text, Path.GetFileName(txtMain.Text));
            CopyDirectory(mainDir, archivePath);
            Directory.Move(mainDir, txtWork.Text);
            Directory.CreateDirectory(mainDir);
            ParseExcel(txtWork.Text, excelPath, dirName);

            //ParseExcel writes the row errors into the working copy of the excel file,
            //and the working folder is deleted right below - keep the annotated file in the archive
            SaveExcelToArchive(excelPath, archivePath);

            Directory.Delete(txtWork.Text, true);
            Directory.CreateDirectory(txtWork.Text);
        }
        private void SaveExcelToArchive(string excelPath, string archivePath)
        {
            try
            {
                if (!File.Exists(excelPath))
                {
                    return;
                }
                Directory.CreateDirectory(archivePath);
                File.Copy(excelPath, Path.Combine(archivePath, Path.GetFileName(excelPath)), true);
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Log(ex);
            }
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
                    var clientType = xlWorksheet.Range["I" + i, "I" + i].Value2?.ToString();
                    var clientCode = xlWorksheet.Range["J" + i, "J" + i].Value2?.ToString();
                    var clientName = xlWorksheet.Range["K" + i, "K" + i].Value2?.ToString();


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
                        clientType = clientType,
                        clientCode = clientCode,
                        clientName = clientName,
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
            if (line <= 0)
            {
                return;
            }
            var xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = null;
            Excel._Worksheet xlWorksheet = null;
            try
            {
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;
                xlWorkbook = xlApp.Workbooks.Open(excelPath);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
                xlWorksheet.Range["L" + line, "L" + line].Value = msg;

                xlWorkbook.Save();
            }
            catch (Exception ex)
            {
                //a failed error report must not stop the remaining rows from being processed
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
                if (xlWorksheet != null)
                    Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false);
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
        
        private void ParseExcel(string workingPath, string excelPath, string dirName)
        {
            var list = new List<string>();
            var dels = new List<string>();
            var rowCounter = 0;
            string vsrname = string.Empty;
            string id = string.Empty;
            int line = -1;
            int counter = 0;
            Gimlas = Code_Gimla_Groups.Split(',');
            Malals = Code_Malal_Groups.Split(',');

            UpdateUI("שומר קודים ביטוח לאומי");
            var codes = GetCodes();
            if(codes.Count() == 0)
            {
                throw new Exception("קודים של ביטוח לאומי לא נקלטו כראוי");
            }
            
            try
            {
                UpdateUI("בודק שורות באקסל. זה יכול לקחת כמה רגעים");
    
                var lst = ReadExcel(excelPath);
                if(lst.Count() == 0)
                {
                    UpdateUI("בדיקת האקסל נכשלה - בדוק קובץ לוג");
                    return;
                }
                UpdateUI("בדיקת האקסל עברה בהצלחה");
                
                var fileList = prepareList(lst);
                if(fileList == null)
                {
                    UpdateUI("בדיקת קבצים נכשלה - בדוק קובץ לוג");
                    return;
                }   
                updateScreen(true);
                var lastRow = fileList.Count();
                lblOf.Text = (lastRow - 1).ToString();
                updateCounters(0);
                
                foreach (var it in fileList)
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
                        var clientType = it.clientType;
                        var clientCode = it.clientCode;
                        var clientName = it.clientName;


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
                            else if (double.TryParse(tickDate, out double dblTd))
                            {
                                // Numeric Excel serial (OADate) coming from a date-typed cell.
                                stickdate = DateTime.FromOADate(dblTd).ToString("yyyy-MM-ddTHH\\:mm\\:ss");
                            }
                            else if (DateTime.TryParse(tickDate, out DateTime dtTd))
                            {
                                // Text date such as "2021-01-01" - keep the date and append the time.
                                stickdate = dtTd.ToString("yyyy-MM-ddTHH\\:mm\\:ss");
                            }
                            else
                            {
                                // Unrecognized format - pass through as-is rather than throwing.
                                stickdate = tickDate;
                            }

                        }

                        var pdvsr = Code_Vsr + "_" + id;
                        var pdvsr2 = Code2_Vsr + "_" + id;
                       // var pdyip = id + "_" + Code_Yip;
                        var vsrpath = string.Empty;
                        var yippath = string.Empty;
                        var newVsrPath = string.Empty;
                        var newYipPath = string.Empty;

                        string[] files = Directory.GetFiles(workingPath);
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
                            newVsrPath = Path.Combine(workingPath, vsrname);
                            newYipPath = Path.Combine(workingPath, yipname);
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
                            updateCounters(line - 1);
                            
                            if (!File.Exists(vsrpath))
                            {
                                throw (new Exception("קובץ וסר לא קיים" + " --- " + "שורה" + "  " + line));
                            }
                            if (!File.Exists(yippath))
                            {
                                throw (new Exception("קובץ יפוי כח לא קיים" + " --- " + "שורה" + "  " + line));
                            }

                        }


                        xmlCreator(id, workingPath, xmlname, newType, OrderCompanyID, FirstName, LastName, stickdate, typecode, vsrname, yipname, clientType, clientCode, clientName);

                        list.Add(newVsrPath);
                        list.Add(newYipPath);
                        list.Add(Path.Combine(workingPath, xmlname));


                        // }
                        rowCounter++;
                        updateCounters(counter);
                    }
                    catch (Exception ex)
                    {
                        SimpleLogger.SimpleLog.Log(ex);
                        SetExcelError(line, ex.Message, excelPath);
                    }


                }

                updateScreen(false);
                UpdateUI("מוחק קבצים");
                
                foreach (var item in dels)
                {
                    if (File.Exists(item))
                    {
                        File.Delete(item);
                    }
                }
                UpdateUI("מעתיק קבצים לתקיית יעד");

                parseList(list, dirName);
                

                MessageBox.Show("הפעולה הסתיימה בהצלחה. " + rowCounter + " מתוך " + lastRow + " נקלטו ");
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Log(ex);
            }
        }
        private void parseList(List<string> list, string dirName)
        {
            var dest = Path.Combine(txtDest.Text, dirName);
            Directory.CreateDirectory(dest);
            string efile = string.Empty;
            try
            {
                foreach (var file in list)
                {
                    efile = Path.GetFileName(file);
                    if (File.Exists(file) && file.Contains("YIP"))
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
        }
        private void xmlCreator(string id, string workingPath, string xmlname, string newType,
            string OrderCompanyID, string FirstName, string LastName, string stickdate, string typecode,
            string vsrname, string yipname, string clientType, string clientCode, string clientName)
        {
            id = id.PadLeft(9, '0');
            var sts = new XmlWriterSettings()
            {
                Indent = true
            };

            using (var writer = XmlWriter.Create(Path.Combine(workingPath, xmlname), sts))
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
                writer.WriteString(DateTime.Now.ToString("yyyy-MM-ddTHH\\:mm\\:ss"));
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
                writer.WriteStartElement("RetrieveDocDate");
                writer.WriteEndElement();
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
                writer.WriteStartElement("DataRequesterTypeID");
                writer.WriteString(clientType);
                writer.WriteEndElement();
                writer.WriteStartElement("DataRequesterKey");
                writer.WriteString(clientCode);
                writer.WriteEndElement();
                writer.WriteStartElement("DataRequesterName");
                writer.WriteString(clientName);
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
        }
        private List<ExcelVals> prepareList(List<ExcelVals> lst)
        {
            try
            {
                var lowList = new List<ExcelVals>();
                var highList = new List<ExcelVals>();
                // Strip the time-of-day so the picked day is a clean midnight boundary;
                // Excel tick dates are whole-day serials and otherwise the boundary day
                // would be misclassified as "before".
                var threshold = dateTimePicker1.Value.Date.ToOADate();
                foreach (var item in lst)
                {
                    // Rows whose tickDate is empty, "1900..." or any text format do not
                    // parse as a numeric OADate. Treat them as pre-threshold so they
                    // remain in the de-dup pool instead of being silently dropped
                    // (which was the regression introduced with the MLL Date filter).
                    if (double.TryParse(item.tickDate, out double value) && value >= threshold)
                    {
                        highList.Add(item);
                    }
                    else
                    {
                        lowList.Add(item);
                    }
                }

                var query = lowList.GroupBy(x => new { x.clientFile, x.code, x.newType })
                            .Select(g => g.OrderByDescending(c =>
                                double.TryParse(c.tickDate, out double v) ? v : double.MinValue))
                            .ToList();


                var fileList = new List<ExcelVals>(highList);
                foreach (var item in query)
                {
                    var ilst = item.ToList();
                    var gmla = ilst[0].code;
                    var mll = ilst[0].newType;
                    if (Gimlas.Contains(gmla) && Malals.Contains(mll))
                    {
                        var itm = ilst[0];
                        ilst = new List<ExcelVals>
                        {
                            itm
                        };
                    }
                    fileList.AddRange(ilst);
                }
                
                return fileList;
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Log(ex);
                return null;
            }
        }
        static void CopyDirectory(string sourceDir, string destDir)
        {
            // Get the subdirectories in the source directory
            string[] subDirectories = Directory.GetDirectories(sourceDir);

            // Copy files in the current directory
            string[] files = Directory.GetFiles(sourceDir);
            foreach (string file in files)
            {
                string fileName = Path.GetFileName(file);
                string destFile = Path.Combine(destDir, fileName);
                File.Copy(file, destFile, true); // Set overwrite to true to overwrite existing files
            }

            // Recursively copy subdirectories
            foreach (string subDir in subDirectories)
            {
                string dirName = Path.GetFileName(subDir);
                string destSubDir = Path.Combine(destDir, dirName);
                CopyDirectory(subDir, destSubDir);
            }
        }
        private void updateCounters(int counter)
        {
            lblRow.Text = counter.ToString();
            System.Windows.Forms.Application.DoEvents();
        }
        private void updateScreen(bool isStart)
        {
            lblMsg.Visible = !isStart;
            lblOf.Visible = isStart;
            lblRow.Visible = isStart;
            label3.Visible = isStart;
            label5.Visible = isStart;
            System.Windows.Forms.Application.DoEvents();
        }
        private void UpdateUI(string message)
        {
            lblMsg.Text = message;
            System.Windows.Forms.Application.DoEvents();
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.DateMLL = dateTimePicker1.Value;
            Properties.Settings.Default.Save();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    txtWork.Text = fbd.SelectedPath;
                    Properties.Settings.Default.WorkPath = fbd.SelectedPath;
                    Properties.Settings.Default.Save();
                }
            }
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
        public string clientType { get; set; }
        public string clientCode { get; set; }
        public string clientName { get; set; }
        public int line { get; set; }
    }
}
