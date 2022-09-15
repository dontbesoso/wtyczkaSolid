using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using NPOI.HSSF.UserModel;

using NPOI.SS.UserModel;
using NPOI.SS.Util;

using System.Data.SqlClient;

using Newtonsoft.Json;

using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;


namespace TestAddIn
{
    public partial class Form1 : Form
    {

        public static int sessionID;
        public static string sessionTimestamp;

        public static string elementName = "";
        public static string projectElementName = "uzupełnij nazwę";
        public static string elementRev = "";
        
        public static string projectExportPath = "";
        public static string registryPath = "";
        public static string exportIdentity = "";

        public static string osobaWydajaca = "";
        public static string pageRange = "";

        public static string xlsData = "";
        public static string xlsName = "";
        public static string xlsOsoba = "";
        
        public static string configFileName = @"C:\\solidComponent\\config.json";

        public static bool debugMode = false;

        public Form1()
        {
            InitializeComponent();
            loadAppConfig();
            processFile(textBox2, textBox1, textBox4, button3, textBoxLog);
        }


        static void mergeToPdfRange_simple(string filePath, string elementName)
        {
            string[] fileEntries = Directory.GetFiles(filePath);

            List<string> tmpOperations = new List<string>();

            Console.WriteLine("test");

            string tmpOperation = "";

            for (int i = 0; i < fileEntries.Count(); i++)
            {
                char[] separators = new char[] { '_', '.' };
                string[] tmpCatch = fileEntries[i].Split(separators);
                if (tmpOperation != tmpCatch[2])
                {
                    tmpOperations.Add(tmpCatch[2]);
                    tmpOperation = tmpCatch[2];
                }
            }

            for (int i = 0; i < tmpOperations.Count; i++)
            {

                PdfDocument outputPDFDocument = new PdfDocument();

                foreach (string fileName in fileEntries)
                {
                    if (fileName.Contains("_" + tmpOperations[i] + ".pdf"))
                    {

                        //Console.WriteLine("Otiweram: " + tmpOperations[i] + ", przepisuję: " + fileName);
                        PdfDocument inputPDFDocument = PdfReader.Open(fileName, PdfDocumentOpenMode.Import);

                        foreach (PdfPage page in inputPDFDocument.Pages)
                        {
                            outputPDFDocument.AddPage(page);
                        }
                        File.Delete(fileName);
                    }

                }
                outputPDFDocument.Save(filePath + "\\" + elementName + "_" + tmpOperations[i] + ".pdf");

                string sqlFullPath = filePath + "\\" + elementName + "_" + tmpOperations[i] + ".pdf";
                string sqlFileName = elementName + "_" + tmpOperations[i] + ".pdf";

                if (!debugMode)
                {
                    genCatalogEntrySQL(tmpOperations[i], (i+1).ToString(), sqlFullPath, sqlFileName);
                }
            }
        }

        static void delOldCatalogEntrySQL()
        {
            SqlConnection cnn;
            string connetionString = "Data Source=metrix-sql;Initial Catalog=Adam_Asprova;User ID=amada;Password=amada";
            cnn = new SqlConnection(connetionString);
            try
            {
                cnn.Open();
                string queryUpdate = "UPDATE itWydaniaDokumentacji SET isActive = 0 WHERE elementName = @elementName";
                SqlCommand command_update = new SqlCommand(queryUpdate, cnn);
                command_update.Parameters.AddWithValue("@elementName", elementName);
                command_update.ExecuteNonQuery();

                cnn.Close();
            } catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        static void genCatalogEntrySQL(string parOperationNumber, string parOperationPage, string parPath, string parFileName)
        {
            string connetionString = null;
            SqlConnection cnn;
            connetionString = "Data Source=metrix-sql;Initial Catalog=Adam_Asprova;User ID=amada;Password=amada";
            cnn = new SqlConnection(connetionString);

            String timeStamp = sessionTimestamp;

            try
            {
                cnn.Open();

                String queryInsert = "INSERT INTO itWydaniaDokumentacji ( sessionId, elementName, elementRev, operationNumber, operationPage, genPerson, genDate, genSource, filePath, fileName, isActive) VALUES (@sessionId, @elementName ,@elementRev ,@operationNumber, @operationPage, @genPerson, @genDate, @genSource, @filePath, @fileName, @isActive)";

                SqlCommand command_insert = new SqlCommand(queryInsert, cnn);

                command_insert.Parameters.AddWithValue("@sessionId", sessionID);
                command_insert.Parameters.AddWithValue("@elementName", elementName);
                command_insert.Parameters.AddWithValue("@elementRev", elementRev);
                command_insert.Parameters.AddWithValue("@operationNumber", parOperationNumber);
                command_insert.Parameters.AddWithValue("@operationPage", parOperationPage);
                command_insert.Parameters.AddWithValue("@genPerson", xlsOsoba);
                command_insert.Parameters.AddWithValue("@genDate", timeStamp);
                command_insert.Parameters.AddWithValue("@genSource", exportIdentity);
                command_insert.Parameters.AddWithValue("@filePath", parPath);
                command_insert.Parameters.AddWithValue("@filename", parFileName);
                command_insert.Parameters.AddWithValue("@isActive", 1);

                command_insert.ExecuteNonQuery();
                command_insert.Dispose();

                cnn.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static bool getDataFromXls(string needle, TextBox txtBoxLog)
        {
            try
            {
                string fileName = registryPath;
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                HSSFWorkbook workbook = new HSSFWorkbook(fs);
                //XSSFWorkbook workbook = new XSSFWorkbook(fs);
                
                ISheet sheet = workbook.GetSheetAt(1);

                int rowCount = sheet.LastRowNum;


                for (int i = rowCount; i > 0; i--)
                {
                    string tmpNeedle = sheet.GetRow(i).GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue;
                    if (needle == tmpNeedle)
                    {
                        xlsData = sheet.GetRow(i).GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
                        xlsName = sheet.GetRow(i).GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
                        xlsOsoba = sheet.GetRow(i).GetCell(11, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
                        return true;
                    }
                }
                return false;
              } catch (Exception e)
                {
                MessageBox.Show(e.Message, "Błąd odczytu z rejestru", MessageBoxButtons.OK);
                txtBoxLog.AppendText("Nie można pobrać danych z arkusza. Czy arkusz jest otwarty w programie Excell?");
                return false;
                }
         }

        public static string usunSpecjalne(string input, TextBox textBoxLog)
        {
            Regex r = new Regex("(?:[^a-z0-9]|(?<=['\"])s)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
            return r.Replace(input, String.Empty);
        }

   

        public static void generujPdfPoStronie(int strona, string detal, string numerOperacji, TextBox textBoxLog)
        {
            SolidEdgeFramework.Application objApplication = null;
            object objVal = null;
 

            SolidEdgeFramework.SolidEdgeDocument objDraftDocument = null;

            objVal = strona.ToString();

            objApplication = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            objApplication.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSheetsRange, objVal);

            objDraftDocument = (SolidEdgeFramework.SolidEdgeDocument)objApplication.ActiveDocument;

            string longStrona = String.Format("{0:0000}", strona);

            string sqlPath = projectExportPath;
            string sqlFileName = longStrona + "_" + detal + "_" + numerOperacji + ".pdf";
            string fileName = projectExportPath + "\\" + longStrona + "_" + detal + "_" + numerOperacji + ".pdf";
            textBoxLog.AppendText("\r\nPlik eksportu: " + fileName);

            if (File.Exists(fileName)) { 
                File.Delete(fileName); 
            }
            objDraftDocument.SaveAs(fileName, null, true, null, null, null, null, null, null);
        }

        public static void generujPdfZakres(int start, int stop, string detal, string numerOperacji, TextBox textBoxLog)
        {
            SolidEdgeFramework.Application objApplication = null;
            object objVal = null;

            SolidEdgeFramework.SolidEdgeDocument objDraftDocument = null;

            objVal = start.ToString() + "-" + stop.ToString();


            objApplication = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            objDraftDocument = (SolidEdgeFramework.SolidEdgeDocument)objApplication.ActiveDocument;


            objApplication.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSheetsRange, objVal);


            string longStrona = String.Format("{0:000}", start) + "-" + String.Format("{0:000}", stop);

            string sqlPath = projectExportPath;
            string sqlFileName = longStrona + "_" + detal + "_" + numerOperacji + ".pdf";
            string fileName = projectExportPath + "\\" + longStrona + "_" + detal + "_" + numerOperacji + ".pdf";
            textBoxLog.AppendText("\r\n" + fileName);

            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            objDraftDocument.SaveAs(fileName, null, true, null, null, null, null, null, null);
            genCatalogEntrySQL(numerOperacji, start.ToString(), sqlPath, sqlFileName);
        }


        static void loadAppConfig()
        {
            try {
                using (StreamReader r = new StreamReader(configFileName))
                {
                    string json = r.ReadToEnd();
                    List<configFile> items = JsonConvert.DeserializeObject<List<configFile>>(json);

                    foreach (var item in items)
                    {
                        registryPath = item.registryPath;
                        exportIdentity = item.exportIdentity;
                        debugMode = item.debugMode;
                    }
                }
            } catch(Exception e) {
                MessageBox.Show(e.Message, "JSON", MessageBoxButtons.OK);
            }

        }

        static void initSession()
        {
            try
            {
                SqlConnection cnn;
                string connetionString = "Data Source=metrix-sql;Initial Catalog=Adam_Asprova;User ID=amada;Password=amada";
                cnn = new SqlConnection(connetionString);

                cnn.Open();

                string queryGetSessionId = "SELECT MAX(id) FROM [Adam_Asprova].[dbo].[itWydaniaDokumentacji]";

                SqlCommand command_getSession = new SqlCommand(queryGetSessionId, cnn);

                SqlDataReader result = command_getSession.ExecuteReader();

                while (result.Read())
                {
                    sessionID = ((int)result.GetValue(0)) + 1;
                }

                result.Close();
                command_getSession.Dispose();
                cnn.Close();
                sessionTimestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "JSON", MessageBoxButtons.OK);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="txtName"></param>
        /// <param name="txtDate"></param>
        /// <param name="txtWho"></param>
        /// <param name="prcButton"></param>
        /// <param name="txtBoxLog"></param>
        static void processFile(TextBox txtName, TextBox txtDate, TextBox txtWho, Button prcButton, TextBox txtBoxLog)
        {
            SolidEdgeFramework.Application objApplication = null;
            SolidEdgeFramework.SolidEdgeDocument objDraftDocument = null;
            try
            {

                objApplication = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                objDraftDocument = (SolidEdgeFramework.SolidEdgeDocument)objApplication.ActiveDocument;

                // sprawdź warunki wejściowe!

                int indeks = objDraftDocument.Name.IndexOf(" ");

                projectExportPath =  objDraftDocument.Path + "\\" + "dokumentacjaPDF";
                string[] filenameParts = objDraftDocument.Name.Split(' ');

                elementName = (string)filenameParts[0]; 
                elementRev = (string)filenameParts[1];

                Directory.CreateDirectory(projectExportPath);

                if (getDataFromXls(elementName, txtBoxLog))
                {
                    txtWho.Text = xlsOsoba;
                    txtDate.Text = xlsData;
                    txtName.Text = xlsName;
                } else
                {
                    if (!debugMode)
                    {
                        prcButton.Enabled = false;
                    }
                    txtBoxLog.AppendText(string.Format("\r\nNie zanleziono wpisu dla elementu: {0}\r\nEksport dokumentacji jest niemożliwy.", elementName));
                    txtBoxLog.AppendText("Upewnij się, że plik eksportu posiada właściwą nazwę: \"<element> <rewizja> .dft\"");
                }

                try
                {
                    string[] fileEntries = Directory.GetFiles(projectExportPath);
                    foreach (string filename in fileEntries)
                    {
                        File.Delete(filename);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "Błąd", MessageBoxButtons.OK);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        
            static void putStampAndExport(int start, int stop, string ktoWydal, string dataWydania, string coWydal, string revWydania, string nrOperacji, TextBox textBoxLog)
        {
            SolidEdgeFramework.Application objApplication = null;
            SolidEdgeDraft.DraftDocument objDraftDocument = null;
            SolidEdgeDraft.Sections objSections = null;
            SolidEdgeDraft.Section objSection = null;

            SolidEdgeDraft.SectionSheets objSheets = null;
            SolidEdgeDraft.Sheet objSheet = null;
            SolidEdgeDraft.SheetSetup objSheetSetup = null;
            SolidEdgeFrameworkSupport.TextBoxes objTextBoxes = null;
            SolidEdgeFrameworkSupport.TextBox textBox = null;
            try
            {
                objApplication = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                objDraftDocument = (SolidEdgeDraft.DraftDocument)objApplication.ActiveDocument;
                objSections = objDraftDocument.Sections;
                //objSection = objSections.Item(1);
                objSection = objSections.WorkingSection;
                objSheets = objSection.Sheets;
                objSheet = objSheets.Item(start);
                objSheetSetup = objSheet.SheetSetup;

                string sheetSize = TestAddIn.appHelper.getSheetSize(objSheetSetup);


                objTextBoxes = (SolidEdgeFrameworkSupport.TextBoxes)objSheet.TextBoxes;

                double targetY, targetX, targetScale;
                
                switch (sheetSize)
                {
                    case "A4": targetX = 0.005; targetY = 0.297; targetScale = 0.7;
                        break;
                    case "A4R": targetX = 0.000; targetY = 0.210; targetScale = 1.1;
                        break;
                    case "A3": targetX = 0.005; targetY = 0.418; targetScale = 1.1;
                        break;
                    case "A3R": targetX = 0.000; targetY = 0.297; targetScale = 1.1;
                        break;
                    default: targetX = 0.100; targetY = 0.200; targetScale = 1.1;
                        break;
                }

                textBox = objTextBoxes.Add(targetX, targetY, 0); 
                    
                    textBox.TextScale = targetScale;
                    textBox.TextControlType = SolidEdgeFrameworkSupport.TextControlTypeConstants.igTextFitToContent;
                    textBox.VerticalAlignment = SolidEdgeFrameworkSupport.TextVerticalAlignmentConstants.igTextHzAlignVCenter;
                    textBox.HorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter;
                    textBox.Text = "    Wydanie aktualne | " + " Data wydania: " + dataWydania + " | Osoba wydająca: " + ktoWydal + "    ";
                    textBox.Edit.Color = 255 * 65536 + 255 * 256 + 255; 
                    textBox.Fill = true;
                    textBox.FillColor = 128 * 65536 + 128 * 256 + 0;
                    textBox.BorderOffset = 2;
                    textBox.Edit.Font = "Arial";
                    
                    generujPdfPoStronie(start, coWydal, nrOperacji, textBoxLog);
                    
                    textBox.Delete();

            }
            catch (Exception error)
            {
                string errorMessageCaption = "Wystąpił błąd podczas nanoszenia stempla";
                MessageBoxButtons przycisk = MessageBoxButtons.OK;

                MessageBox.Show(error.Message, errorMessageCaption, przycisk);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!debugMode)
            {
                initSession();
                delOldCatalogEntrySQL();
            }
            
            if (textBox4.Text == "")
            {
                string errorMessageCaption = "Wystąpił błąd";
                MessageBoxButtons przycisk = MessageBoxButtons.OK;

                MessageBox.Show("Nie wskazano osoby wydającej dokumentację.\nUzupełnij pole.", errorMessageCaption, przycisk);
                return;

            } else
            {
                osobaWydajaca = textBox4.Text;
            }

            textBoxLog.AppendText("\r\n-- Rozpoczęto eksport dokumentacji -- ");

            textBoxLog.AppendText(string.Format("\r\n-- Autor: {0} ", osobaWydajaca));
            textBoxLog.AppendText(string.Format("\r\n-- Element: {0}, rewizja: {1} ", elementName, elementRev));


            SolidEdgeFramework.Application objApplication = null;
            SolidEdgeDraft.DraftDocument objDraftDocument = null;

            SolidEdgeDraft.Sections sections = null;
            SolidEdgeDraft.Section section = null;

            SolidEdgeDraft.SectionSheets sectionSheets = null;

            try
            {
                objApplication = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                objDraftDocument = (SolidEdgeDraft.DraftDocument)objApplication.ActiveDocument;

                sections = objDraftDocument.Sections;
                section = sections.WorkingSection;
                sectionSheets = section.Sheets;

                int tmpPage = 1;
                string tmpOperationNumber = "";
                foreach (SolidEdgeDraft.Sheet sheet in sectionSheets)
                {
                    sheet.Activate();
                    
                    SolidEdgeDraft.SheetSetup objSheetSetup = sheet.SheetSetup;


                    bool tmpFoundPageNumber = false;

                    foreach (SolidEdgeFrameworkSupport.TextBox objTextBox in (SolidEdgeFrameworkSupport.TextBoxes)sheet.TextBoxes)
                    {
                        // zagnieżdżone *.asm są albo w zbiorze arkuszy, albo w zbiorze textboxes, coś w obiektach, a powiązania są zerwane
                        if ((Regex.IsMatch(usunSpecjalne(objTextBox.Text, textBoxLog), @"(^\d{2}$)|(^\d{3}$)")) && ((int.Parse(usunSpecjalne(objTextBox.Text, textBoxLog))) % 5 == 0))
                        {
                            double x, y, z;
                            objTextBox.GetOrigin(out x, out y, out z);         
                            if ((x >= (objSheetSetup.SheetWidth - 0.030)) && (y >= objSheetSetup.SheetHeight - 0.030))
                            {
                                tmpOperationNumber = usunSpecjalne(objTextBox.Text, textBoxLog);
                                tmpFoundPageNumber = true;
                            }
                                
                        }
                    }

                    if (tmpFoundPageNumber == true)
                        try
                        {
                            putStampAndExport(tmpPage, tmpPage, xlsOsoba, xlsData, elementName, elementRev, tmpOperationNumber, textBoxLog);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "błąd", MessageBoxButtons.OK);
                        }
                    else
                    {
                        string sheetName = sheet.Name.ToString();

                        if (sheetName.IndexOf("KOOP", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            putStampAndExport(tmpPage, tmpPage, xlsOsoba, xlsData, elementName, elementRev, "kooperacja", textBoxLog);
                        }
                        else
                        {
                            textBoxLog.AppendText(string.Format("\r\n================\r\nNa stronie {0} nie znaleziono indeksu strony.\r\nEksport tej ztrony został pominięty.\r\nSprawdź poprawność opisu tej strony", sheet.Name));
                        }
                    }
                    tmpPage++;
                }
                
                mergeToPdfRange_simple(projectExportPath, elementName);

                textBoxLog.AppendText("\r\n-- Zakończono eksport dokumentacji -- ");
                return;

            } catch(Exception error)
            {   
                string errorMessageCaption = "Wystąpił błąd";
                MessageBoxButtons przycisk = MessageBoxButtons.OK;

                MessageBox.Show(error.InnerException.ToString(), errorMessageCaption, przycisk);
                return;
            }


        }
    }
}
