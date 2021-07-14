using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DocumentFormat.OpenXml.Packaging;
using OSCALHelperClasses;
using System.Data;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Text;
using System.Configuration;
using System.Data.SqlClient;

namespace OSCAL_POAM_Mapper
{
  
    public partial class _Default : Page
    {

       
        public const string XMLNamespace = @"http://csrc.nist.gov/ns/oscal/1.0";
        public const string POAMschema = "oscal_poam_schema.xsd";
        public string CurrentTableName;
        public string WordTempFilePath;
        public string TemplateFile = "FedRAMP-POAM-Template.xlsx";
        protected private bool OverwriteXMLMapping;
        string message = "";
        int percent = 0;
        delegate string ProcessTask(string id);
        static string ProcessingPage=String.Empty;
        string Filename;
        static string poamTitle = String.Empty;
        static string poamDate = String.Empty;

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void UploadButton_Click(object sender, EventArgs e)
        {

            StatusLabel1.Text = "";
            OpenMyFile.Visible = false;

            if (FileUpload1.HasFile)
            {
                try
                {
                    string filename = Path.GetFileName(FileUpload1.FileName);
                    //  PrintProgressBar("", 0);
                    Filename = filename;
                    message = string.Format("Starting the conversion of the OSCAL POAM XML {0}", filename);
                    percent = 2;
                    PrintProgressBar(message, percent, true);
                    CollapseDiv("mainForm");
                    string fileExtension = System.IO.Path.GetExtension(FileUpload1.FileName).ToLower();
                    if (fileExtension != ".xml")
                    {
                        StatusLabel1.ForeColor = System.Drawing.Color.Red;
                        StatusLabel1.Text = "Invalid File Extension - Not an XML File!";
                    }
                    else
                    {

                        FileUpload1.SaveAs(Server.MapPath("~/Uploads/") + filename);
                        StatusLabel1.ForeColor = System.Drawing.Color.Green;
                        StatusLabel1.Text = "Upload status: File sucessfully uploaded...  Processing File...Please stand by.";

                        string DocumentPath;
                        string xmlSchemaPath = HttpContext.Current.Server.MapPath(string.Format(@"~\Template\{0}", POAMschema));

                        //PseudoValidator(filename, xmlSchemaPath);
                        message = string.Format("Conversion of {0}: Successful validation of the xml file  {0} against OSCAL schema with namespace {1}", Filename, XMLNamespace);
                        percent = 10;
                        PrintProgressBar(message, percent);


                        ProcessData(filename, TemplateFile, out DocumentPath);

                        StatusLabel1.Text = "Processing Complete.. Click below to open file.";

                        OpenMyFile.Visible = true;
                        Cache["outputFile"] = filename;

                    }
                }
                catch (Exception ex)
                {


                    StatusLabel1.ForeColor = System.Drawing.Color.Red;
                    StatusLabel1.Text = "Upload status: The file could not be uploaded. The following error occured: " + ex.Message;


                }


            }

        }
        public static string GetXMLElement(string XMLFileName, string ElementName)
        {
            string xmlDataFile = HttpContext.Current.Server.MapPath(string.Format(@"~\Uploads\{0}", XMLFileName));
            string bytes = File.ReadAllText(xmlDataFile);
            System.IO.MemoryStream myStream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(bytes));
            string myElementValue = "";
            System.Xml.XmlReader xr = System.Xml.XmlReader.Create(myStream);
            while (xr.Read())
            {
                if (xr.NodeType == System.Xml.XmlNodeType.Element)
                    if (xr.Name == ElementName.ToString())
                    {
                        myElementValue = xr.ReadElementContentAsString();
                        break;
                    }
            }

            return (myElementValue);
        }

        public static void ProcessData(string poam_file, string poam_template_file, out string DocumentPath)
        {
            string xmlDataFile = HttpContext.Current.Server.MapPath(string.Format(@"~\Uploads\{0}", poam_file));
            string templateDocument = HttpContext.Current.Server.MapPath(string.Format(@"~/Template/{0}", poam_template_file));
            string tempDocument = HttpContext.Current.Server.MapPath(string.Format(@"~/Downloads/{0}", "MyGeneratedDocument.csv"));
            //string outputDocument = HttpContext.Current.Server.MapPath(string.Format(@"~/Downloads/{0}", poam_template_file.Replace("Template", "OSCAL")));
            string outputDocument = HttpContext.Current.Server.MapPath(string.Format(@"~/Downloads/{0}", poam_file.Replace("xml","xlsx")));
            string message = "";
            int percent = 0;

            // Get first occurrence meta data
            string xmlElement = "title";
            poamTitle = GetXMLElement(poam_file, xmlElement);
            xmlElement = "last-modified";
            poamDate = GetXMLElement(poam_file, xmlElement);

            message = string.Format("Conversion of {0} in progress...", poam_file);

          

            percent = 10;
            PrintProgressBar(message, percent);

            // Get data via SQL call into DataTable
            DataTable mydt = new DataTable();
            mydt = BindPOAMData(xmlDataFile);
            DataSet myDS = new DataSet("poams");
            myDS.Tables.Add(mydt);


            //if (File.Exists(tempDocument))
            //{
            //    File.Delete(tempDocument);
            //}
            //File.Copy(templateDocument, tempDocument);
    

            // Export data to .xlsx file
            if (File.Exists(outputDocument))
            {
                File.Delete(outputDocument);
            }
            ExportDataSet(myDS, outputDocument);
;

            //File.Copy(tempDocument, outputDocument);
            //File.Delete(tempDocument);
            DocumentPath = outputDocument;

        }

        public static  DataTable BindPOAMData(string fname)
        {
            string CS = ConfigurationManager.ConnectionStrings["CONVERTREPOConnectionString"].ConnectionString;
            SqlConnection con = new SqlConnection(CS);
            SqlCommand cmd = new SqlCommand("[dbo].[GetPOAMLIST]", con);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("@SourceFile", fname.ToString());

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                con.Open();
                da.Fill(dt);
                con.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return (dt);
        }

        public static void ExportDataSet(DataSet ds, string destination)
        {


            using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                foreach (System.Data.DataTable table in ds.Tables)
                {

                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                    sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                    DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    uint sheetId = 1;
                    if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                    {
                        sheetId =
                            sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    // Add Poam Document Title
                    DocumentFormat.OpenXml.Spreadsheet.Row titleRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    DocumentFormat.OpenXml.Spreadsheet.Cell tcell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    tcell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    tcell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(poamTitle);
                    titleRow.AppendChild(tcell);
                    sheetData.AppendChild(titleRow);

                    //Add Information Row Title
                    DocumentFormat.OpenXml.Spreadsheet.Row infoRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    DocumentFormat.OpenXml.Spreadsheet.Cell icell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    icell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    icell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("CSP");
                    infoRow.AppendChild(icell);

                    icell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    icell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    icell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("System Name");
                    infoRow.AppendChild(icell);

                    icell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    icell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    icell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Impact Level");
                    infoRow.AppendChild(icell);

                    icell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    icell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    icell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("POAM Date");
                    infoRow.AppendChild(icell);
                    sheetData.AppendChild(infoRow);


                    //  Add information row values
                    DocumentFormat.OpenXml.Spreadsheet.Row blankRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    sheetData.AppendChild(blankRow);
                    //DocumentFormat.OpenXml.Spreadsheet.Row dRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    //DocumentFormat.OpenXml.Spreadsheet.Cell dcell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    //dcell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    //dcell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("My CSP");
                    //infoRow.AppendChild(dcell);
                    //dcell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    //dcell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    //dcell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("TestSystem");
                    //infoRow.AppendChild(dcell);
                    //dcell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    //dcell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    //dcell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("MODERATE");
                    //infoRow.AppendChild(dcell);
                    //dcell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    //dcell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    //dcell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(poamDate);
                    //dRow.AppendChild(dcell);
                    //sheetData.AppendChild(dRow);




                    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    List<String> columns = new List<string>();
                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }


                    sheetData.AppendChild(headerRow);

                    foreach (System.Data.DataRow dsrow in table.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        foreach (String col in columns)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }

                }
            }
        }

        static void PrintProgressBar(string Message, int PercentComplete, bool first = false)
        {
            var sb = new StringBuilder();
            sb.Append("<script>");
            var iis = string.Format("\"{0}%\"", PercentComplete);
            sb.AppendLine(string.Format("myFunction(\"{0}\",{1})", Message, iis));
            sb.Append("</script>");

            var file = "";
            var update = sb.ToString();
            if (first)
            {
                HttpContext.Current.Response.Clear();
                HttpContext.Current.Response.ClearContent();
                file = ProcessingPage + update;  //// 
                HttpContext.Current.Response.Write(file);
            }
            else
            {
                HttpContext.Current.Response.Clear();
                HttpContext.Current.Response.ClearContent();
                file = update;
            }
            HttpContext.Current.Response.Write(file);

            HttpContext.Current.Response.Flush();
        }

        protected void OpenMyFile_Click(object sender, EventArgs e)
        {
            TemplateFile = Cache["outputFile"].ToString(); ;
            Response.Redirect(string.Format(@"~/Downloads/{0}", TemplateFile.Replace("xml", "xlsx")));
        }
        private void CollapseDiv(string entity)
        {
            var sb = new StringBuilder();
            sb.Append("<script type=\"text/javascript\">");
            sb.Append(string.Format("document.getElementById(\"{0}\").style.display='none';", entity));
            sb.Append("</script>");

            var update = sb.ToString();
            HttpContext.Current.Response.Write(update);

            HttpContext.Current.Response.Flush();

        }
       
        public void PseudoValidator(string XmlDocument, string XsdSchemaPath)
        {
            try
            {
                string XmlDocumentPath = HttpContext.Current.Server.MapPath(string.Format(@"~\Uploads\{0}", XmlDocument));

                XmlReaderSettings OscalSettings = new XmlReaderSettings();
                OscalSettings.Schemas.Add(XMLNamespace, XsdSchemaPath);
                OscalSettings.ValidationType = ValidationType.Schema;
                OscalSettings.ValidationEventHandler += new ValidationEventHandler(OSCALSettingsValidationEventHandler);
                XmlReader OscalDoc = XmlReader.Create(XmlDocumentPath, OscalSettings);

                while (OscalDoc.Read()) { }

                OscalDoc.Close();
                // SuccessfulValidation = true;
            }
            catch (Exception ex)
            {
                // SuccessfulValidation = false;
                throw ex;
            }
        }
        private void OSCALSettingsValidationEventHandler(object sender, ValidationEventArgs e)
        {
            if (e.Severity == XmlSeverityType.Warning)
            {

                // StatusLabel1.ForeColor = System.Drawing.Color.Yellow;
                StatusLabel1.Text = e.Message;
            }
            else if (e.Severity == XmlSeverityType.Error)
            {

                // StatusLabel1.ForeColor = System.Drawing.Color.Red;
                StatusLabel1.Text = e.Message;
            }
        }

    }
}