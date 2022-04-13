using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net;
using Newtonsoft.Json;
using System.Web;
using System.Web.Configuration;
using System.Diagnostics;
using System.Web.UI.WebControls;
using System.Collections.Specialized;
using System.Web.Http.Cors;
using ExcelDataReader;
using System.Configuration;

namespace EWApp_Service
{
    public class ExcelHandlerController : ApiController
    {


        //creating object of datatable present on clientside
        static private System.Data.DataTable dt = new System.Data.DataTable();
        static XElement xelement;
        static string LogFileName = "Logs.txt";
        static  ILogger log;
        static ExcelHandlerController()
        {
            try
            {
            /*if (ConfigurationManager.AppSettings["LogType"]=="EventLogger")
             {
                  log = new EventLogger();
             }
            else
            {
                  log = new FileLogger();
            }*/

            string str = ConfigurationManager.AppSettings["LogType"];
            Type instanceType = Type.GetType("EWApp_Service."+str);
            //Type instanceType  = typeof(FileLogger);
            log = (ILogger)Activator.CreateInstance(instanceType);
            
            }
            catch (Exception ex)
            {
                string message = "Failed To Write Log into LogFile Error: " + ex.Message;
                ILogger Ev = new EventLogger();
                Ev.WriteLog(String.Format("{0} @ {1}", DateTime.Now, message, EventLogEntryType.Error),EventLogEntryType.Error);
            }

        }

 /*       public static bool WriteLog(string strFileName, string strMessage)
        {
            try
            {
                FileStream objFilestream = new FileStream(string.Format("{0}\\{1}", ConfigurationManager.AppSettings["DBLocation"], strFileName), FileMode.Append, FileAccess.Write);
                StreamWriter objStreamWriter = new StreamWriter((Stream)objFilestream);
                objStreamWriter.WriteLine(strMessage);
                objStreamWriter.Close();
                objFilestream.Close();
               

               
                return true;

            }
            catch (Exception ex)
            {
                string message = "Failed To Write Log into LogFile Error: "+ex.Message;
                *//*if (!EventLog.SourceExists("EWApp"))
                {
                    EventLog.CreateEventSource("EWApp", "EWLog");
                    EventLog log = new EventLog("EWLog");
                    log.Source = "EWApp";
                    log.WriteEntry(String.Format("{0} @ {1}", DateTime.Now, "LogFile Created "), EventLogEntryType.Information);
                }
                else
                {
                    EventLog log = new EventLog("EWLog");
                    log.Source = "EWApp";
                    log.WriteEntry(String.Format("{0} @ {1}", DateTime.Now, "Failed To Write Log into LogFile,  Error: " + ex.Message), EventLogEntryType.Error);
                }*//*

                
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry(String.Format("{0} @ {1}", DateTime.Now, message  ), EventLogEntryType.Error);
                }

                return false;
            }
        }*/

        public String GetString(Int32 id)
        {
            return "This String";
        }

        [EnableCors(origins: "http://127.0.0.1:5500", headers: "*", methods: "*")]
        public string GetAllFiles()
        {
            try { 
            List<string> files = new List<string>();  
            Dictionary<string, string[]> rawJsonData = new Dictionary<string, string[]>();
            string[] allfiles = Directory.GetDirectories(ConfigurationManager.AppSettings["DBLocation"], "*", SearchOption.AllDirectories);

            foreach (string folderName in allfiles)
            {
                string[] presentExcelfiles = Directory.GetFiles(folderName, "*.xlsx", SearchOption.AllDirectories);
                string[] presentXmlfiles = Directory.GetFiles(folderName, "*.xml", SearchOption.AllDirectories);
                if (presentExcelfiles.Length == 1 && presentXmlfiles.Length == 1)
                {
                   files.Add(Path.GetFileNameWithoutExtension(folderName));
                }
            }

            rawJsonData.Add("FileName", files.ToArray() );
            var json = JsonConvert.SerializeObject(rawJsonData);

                //json = "{ \"FileName\" : "+ json + "}";

            log.WriteLog(String.Format("{0} @ {1}", DateTime.Now, "Clicked On Load File Button" ), EventLogEntryType.Information);
            return json;


            
            }catch (Exception ex)
            {
                log.WriteLog(String.Format("{0} @ {1}", DateTime.Now," Clicked On Load File Button Error: "+ex.Message ), EventLogEntryType.Error);
                return ex.Message;
            }
        }

        public String GetExcelData(string fileName)
        {
            try
            {

            
            string filePath = $"{ConfigurationManager.AppSettings["DBLocation"]}{fileName}\\{fileName}.xlsx";

            //opning selected excel file in background 
            Excel.Application app = new Excel.Application();
            Excel.Workbook xlWorkbook = app.Workbooks.Open(filePath,ReadOnly:true);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            //app.Visible = true;

            //dt.Rows.Clear();
            //dt.Columns.Clear();
            
            //initial values for rows and columns
            int col = 1;
            int row = 1;
            //temporary list to grab values and insert into row
            List<string> list = new List<string>();

           // List<string> Clist = new List<string>();

            Dictionary<string, string[]> ExcelData = new Dictionary<string, string[]>();
            Dictionary<string, Dictionary<string, string[]> > ExcelDataAsjson = new Dictionary<string, Dictionary<string,string[]>>();
            
            //while loops to grab value from imported workbook
            //1st loop will loop till the last row
            //2nd loop will loop till the last columns

            while (xlRange.Cells[row, col] != null && xlRange.Cells[row, col].Value2 != null)
            {
                //creating new row to assgning data to it.
                //DataRow dr = dt.NewRow();
                while (xlRange.Cells[row, col] != null && xlRange.Cells[row, col].Value2 != null)
                {
                    list.Add(xlRange.Cells[row, col].Value.ToString());
                    col++;
                }

                ExcelData.Add($"row{row}",list.ToArray());
                //clearing list so next row could have its own unique data
                list.Clear();
                col = 1;
                row++;
            }

            ExcelDataAsjson.Add("ExcelData",ExcelData);
            var json = JsonConvert.SerializeObject(ExcelDataAsjson);

            string XMLdirectory = $"{ConfigurationManager.AppSettings["DBLocation"]}+{fileName}+\\+{fileName}+_setting.xml";

            if (!File.Exists(XMLdirectory) )
            {
                //creating XML file for imported file
                XDocument doc = new XDocument(new XElement("SettingGroup"));

                foreach (string colName in ExcelData["row1"])
                {
                    doc.Element("SettingGroup").Add(new XElement(colName.Replace(" ","_"), new XAttribute("colName", colName.ToString())
                                                                          , new XAttribute("ReadOnly", false)
                                                                          , new XAttribute("Hidden", false)
                                                    ));
                }
                doc.Save(XMLdirectory);
            }
            
            //Dictionary<string, Dictionary<string, Boolean> > XMLData = ReadXML(XMLdirectory);

            //closing the imported file..after fetching data from it.
            xlWorkbook.Close(true);

                log.WriteLog( String.Format("{0} @ {1}",  DateTime.Now, $"Getting Excel Data FileName:{fileName}, Method: MSXlInterop"), EventLogEntryType.Information);
                return json;
            
            }catch (Exception ex)
            {
                log.WriteLog( String.Format("{0} @ {1}",  DateTime.Now, $"Getting Excel Data FileName:{fileName}, Method: MSXlInterop Error: "+ ex.Message), EventLogEntryType.Error);
                return ex.Message;
            }
        }


       

        [HttpGet]
        public String NewGetExcelData(string fileName)
        {
            try
            {

            
           // Dictionary<string, System.Data.DataTable > XlData = new Dictionary<string, System.Data.DataTable>();
          //  Dictionary<string, string[][]> XlData = new Dictionary<string, string[][]>();
            string filePath = $"{ ConfigurationManager.AppSettings["DBLocation"]}{fileName}\\{fileName}.xlsx";
            

            DataSet result = null;
            //List<string[][]> ts = new List<string[][]>();

            //https://reposhub.com/dotnet/office/ExcelDataReader-ExcelDataReader.html
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    result = reader.AsDataSet();
                    
                }
            }
            System.Data.DataTable dt = result.Tables[0];
            string[] columnNames = (from string x in dt.Rows[0].ItemArray select x.ToString()).ToArray();
            string json = JsonConvert.SerializeObject(result.Tables[0], Newtonsoft.Json.Formatting.Indented);
   
              /*foreach (System.Data.DataTable table in result.Tables)
              {
                  *//*if (//my conditions)
                  {
                      continue;
                  }*//*
                  var rows = table.AsEnumerable().ToArray();

                  var dataTable = new string[table.Rows.Count][];//[table.Rows[0].ItemArray.Length];
                  //Dictionary<string, string[]> dataTable = new Dictionary<string, string[]>;
                  Parallel.For(0, rows.Length, new ParallelOptions { MaxDegreeOfParallelism = 8 },
                      i =>
                      {
                          var row = rows[i];
                          dataTable[i] = row.ItemArray.Select(x => x.ToString()).ToArray();
                      });


                  ts.Add(dataTable);
                  
               
              }*/

             //var json = JsonConvert.SerializeObject(XlData);

             // XlData.Add("XlData", result.Tables[0]);
             // string json = JsonConvert.SerializeObject(XlData, Newtonsoft.Json.Formatting.Indented);




             //System.Web.Script.Serialization.JavaScriptSerializer serializer =
             //new System.Web.Script.Serialization.JavaScriptSerializer();

             //List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
             //Dictionary<string, object> row;
             //foreach (DataRow dr in result.Tables[0].Rows)
             //{
             //    row = new Dictionary<string, object>();
             //    foreach (DataColumn col in result.Tables[0].Columns)
             //    {
             //        //Check if the Column is Filter then Process it through Serialization
             //        var JsonValue = dr[col];

             //        row.Add(col.ColumnName, JsonValue.ToString());
             //    }
             //    rows.Add(row);
             //}

             ////Convert DataTable to Json Format
             //var res = serializer.Serialize(rows);

             string XMLdirectory = $"{ConfigurationManager.AppSettings["DBLocation"]}{fileName}\\{fileName}_setting.xml";

            if (!File.Exists(XMLdirectory))
            {
                //creating XML file for imported file
                XDocument doc = new XDocument(new XElement("SettingGroup"));

                foreach (string colName in columnNames)
                {
                    doc.Element("SettingGroup").Add(new XElement(colName.Replace(" ", "_"), new XAttribute("colName", colName.ToString())
                                                                          , new XAttribute("ReadOnly", false)
                                                                          , new XAttribute("Hidden", false)
                                                    ));
                }
                doc.Save(XMLdirectory);
            }

                //Dictionary<string, Dictionary<string, Boolean> > XMLData = ReadXML(XMLdirectory);

                //closing the imported file..after fetching data from it.

                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Getting Excel Data FileName:{fileName} ,Method: ExcelDataReader "), EventLogEntryType.Information);
                return json;
            
            }catch (Exception ex)
            {
                log.WriteLog(String.Format("{0} @ {1}", DateTime.Now, $"Getting Excel Data FileName:{fileName}, Method: MSXlInterop Error: "+ex.Message), EventLogEntryType.Error);
                return ex.Message;
            }
        }

        [HttpGet]
        public String MultiNewGetExcelData(string fileName)
        {
            try
            {

            
            string filePath = $"{ ConfigurationManager.AppSettings["DBLocation"]}{fileName}\\{fileName}.xlsx";


            DataSet result = null;
            List<System.Data.DataTable> worksheetsList = new List<System.Data.DataTable>();

            //https://reposhub.com/dotnet/office/ExcelDataReader-ExcelDataReader.html
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    result = reader.AsDataSet();

                }
            }
            
            List<string[]> colNamesList = new List<string[]>();

            foreach (System.Data.DataTable table in result.Tables)
            {
                System.Data.DataTable dt = table;
                string[] columnNames = (from string x in dt.Rows[0].ItemArray select x.ToString()).ToArray();
                colNamesList.Add(columnNames);
                worksheetsList.Add(table);
            }
            var json = JsonConvert.SerializeObject(worksheetsList);

            // XlData.Add("XlData", result.Tables[0]);
            // string json = JsonConvert.SerializeObject(XlData, Newtonsoft.Json.Formatting.Indented);


            string XMLdirectory = $"{ConfigurationManager.AppSettings["DBLocation"]}{fileName}\\{fileName}_setting.xml";

                /*if (!File.Exists(XMLdirectory))
                {
                    //creating XML file for imported file
                    XDocument doc = new XDocument(new XElement("SettingGroup"));

                    foreach (string[] columnNames in colNamesList)
                    {

                        foreach (string colName in columnNames)
                        {
                            doc.Element("SettingGroup").Add(new XElement(colName.Replace(" ", "_"), new XAttribute("colName", colName.ToString())
                                                                                  , new XAttribute("ReadOnly", false)
                                                                                  , new XAttribute("Hidden", false)
                                                            ));
                        }               
                    }
                    doc.Save(XMLdirectory);
                }*/

                //Dictionary<string, Dictionary<string, Boolean> > XMLData = ReadXML(XMLdirectory);

                //closing the imported file..after fetching data from it.
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Getting Multi-Sheets Excel Data FileName:{fileName}, Method: ExceDataReader Error:"), EventLogEntryType.Information);
                return json;
            }catch (Exception ex)
            {
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Getting Multi-Sheets Excel Data FileName:{fileName}, Method: ExcelDataReader Error: "+ ex.Message), EventLogEntryType.Error);
                return ex.Message;
            }
        }


        [HttpGet]
        public string ReadXML(string fileName)
        {
            try
            {

            
            string xmlfilePath =$"{ConfigurationManager.AppSettings["DBLocation"]}{fileName}\\{fileName}_setting.xml";
            xelement = XElement.Load(xmlfilePath);
            IEnumerable<XElement> settingGroup = xelement.Elements();
            // Excel.Worksheet ws = Globals.ThisWorkbook.Worksheets[1];
            //int ind = 1;

            Dictionary<string, Dictionary<string,Boolean> > XMLData = new Dictionary<string, Dictionary<string, Boolean> >();
            //Dictionary<string, Boolean> EachColData = new Dictionary<string, Boolean>();

          

            foreach (var setting in settingGroup)
            {
                /*EachColData["ReadOnly"] = Convert.ToBoolean(setting.Attribute("ReadOnly").Value);
                EachColData["Hidden"]   = Convert.ToBoolean(setting.Attribute("Hidden").Value);*/
                XMLData.Add(setting.Attribute("colName").Value.ToString(), new Dictionary<string, Boolean>(){ 
                                                                                    {"ReadOnly",Convert.ToBoolean(setting.Attribute("ReadOnly").Value) },
                                                                                    {"Hidden", Convert.ToBoolean(setting.Attribute("Hidden").Value) } });
                //XMLData.Add(setting.Attribute("colName").Value.ToString(), EachColData);

                /*dt.Columns[setting.Attribute("colName").Value.ToString()].ReadOnly = Convert.ToBoolean(setting.Attribute("ReadOnly").Value);
                ws.Columns[ind].Hidden = Convert.ToBoolean(setting.Attribute("Hidden").Value);
                ind++;*/
            }

            Dictionary<string, Dictionary<string, Dictionary<string, Boolean>>> jsonXMLData = new Dictionary<string, Dictionary<string, Dictionary<string, Boolean>>>();
            jsonXMLData.Add("XMLData",XMLData);
            var json = JsonConvert.SerializeObject(jsonXMLData);
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Reading Xml Data FileName:{fileName}_setting"), EventLogEntryType.Information);
                return json;
            }
            catch (Exception ex)
            {
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Reading Xml Data FileName:{fileName}_setting, Error: "+ ex.Message), EventLogEntryType.Error);
                return ex.Message;
            }
        }



        [HttpGet]
        public string GetFileName(string key)
        {
            try
            {
            string XMLdirectory = $"{ConfigurationManager.AppSettings["DBLocation"]}Entries.xml";
            if (!File.Exists(XMLdirectory))
            {
                //creating XML file for imported file
                XDocument docm = new XDocument(new XElement("Users"));
                docm.Save(XMLdirectory);
            }

            string fileName=string.Empty;
            
            XmlDocument doc = new XmlDocument();
            doc.Load(XMLdirectory);
            // xelement = XElement.Load(XMLdirectory);
            XmlNode node = doc.DocumentElement.SelectSingleNode($"//{key}");
            if(node != null)
            {
                fileName =  node.Attributes["fileName"].Value;  
                doc.DocumentElement.RemoveChild(node);
            }
            
            doc.Save(XMLdirectory);

                /*var xel = AllUsers.Descendants("param")
                  .Where(xElement => xElement.Attribute("name")?.Value == "Super");*/
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Getting FileName corresponds to Key: {key}"), EventLogEntryType.Information);
                return fileName;
            }
            catch (Exception ex)
            {
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Getting FileName corresponds to Key: {key}, Error: "+ ex.Message), EventLogEntryType.Error);
                return ex.Message;
            }
        }


        [HttpPost]
        [Route("/ExcelHandler/StoreKeyFile")]
        [EnableCors(origins: "http://127.0.0.1:5500", headers: "*", methods: "*")]
        public async Task<HttpResponseMessage> StoreKeyFile( string[] fileName)
        {
            try
            {
                //var name = fileName;
              
                string XMLdirectory = $"{ConfigurationManager.AppSettings["DBLocation"]}Entries.xml";

                if (!File.Exists(XMLdirectory))
                {
                    //creating XML file for imported file
                    XDocument doc = new XDocument(new XElement("Users"));
                    doc.Element("Users").Add(new XElement($"user{fileName[0]}", new XAttribute("fileName", fileName[1])));
                    //XDocument doc = new XDocument(new XElement($"user{fileName[0]}", new XAttribute("fileName", fileName[1])));

                    doc.Save(XMLdirectory);
                }
                else
                {

                    XDocument xDocument = XDocument.Load(XMLdirectory);
                    // XElement root = xDocument.Element("Users");
                    xDocument.Element("Users").Add(new XElement($"user{fileName[0]}", new XAttribute("fileName", fileName[1])));
                    xDocument.Save(XMLdirectory);
                }
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Stroring FileName{fileName[1]} And Key:{fileName[0]} (Web-request)"), EventLogEntryType.Information);
                return Request.CreateResponse<string>("Saved");
            }
            catch (Exception ex)
            {
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Stroring FileName:{fileName[1]} And Key:{fileName[0]}, Error: "+ ex.Message), EventLogEntryType.Error);
                return Request.CreateResponse<string>(ex.ToString());
            }
            
        }


        [HttpPost]
        [Route("/ExcelHandler/SaveXML")]
        public async Task<HttpResponseMessage> SaveXML(List<Object> li)
        {
            try
            {

            
            //Console.WriteLine(li);
            string fileName= li[0].ToString();
            string xmlfilePath = $"{ConfigurationManager.AppSettings["DBLocation"]}{fileName}\\{fileName}_setting.xml";

            xelement = XElement.Load(xmlfilePath);
            IEnumerable<XElement> settingGroup = xelement.Elements();

            object[] colsData = li.ToArray();

            //ind 1 bcus fileName is Stored at index 0
            int ind = 1;
            foreach (XElement setting in settingGroup)
            {
                ColData col = JsonConvert.DeserializeObject<ColData>(colsData[ind].ToString());
                
                setting.Attribute("ReadOnly").Value = col.ReadOnly.ToString();
                setting.Attribute("Hidden").Value =   col.Hidden.ToString();
                //setting.Attribute("Hidden").Value = li[ind][""].toString();
                ind++;
            }

            xelement.Save(xmlfilePath);
            log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Saving XML:{fileName}_setting"), EventLogEntryType.Information);

            return Request.CreateResponse<string>("Saved");
            
            }catch (Exception ex)
            {
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Saving XML:{li[0]}_setting Error: "+ ex.Message), EventLogEntryType.Error);
                return Request.CreateResponse<string>(ex.Message);
            }
        }


        [HttpPost]
        [Route("/ExcelHandler/HandleNewFile")]
        public async Task<HttpResponseMessage> HandleNewFile()
        {
            try
            {
            
            if (!Request.Content.IsMimeMultipartContent())
            {
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }
               
            var provider = new MultipartMemoryStreamProvider();     

            await Request.Content.ReadAsMultipartAsync(provider);
           // var filename=string.Empty;
            var file = provider.Contents[0];
            var fileName = Path.GetFileNameWithoutExtension(file.Headers.ContentDisposition.FileName.Trim('\"'));
            //var buffer = await file.ReadAsByteArrayAsync();
            /*foreach (var file in provider.Contents)
            {
                filename = file.Headers.ContentDisposition.FileName.Trim('\"');
                //var buffer = await file.ReadAsByteArrayAsync();
                //Do whatever you want with filename and its binary data.

            }*/

            Stream input = await file.ReadAsStreamAsync();
            var filename = string.Empty;
            string directoryName = String.Empty;
            string URL = String.Empty;
            string tempDocUrl = WebConfigurationManager.AppSettings["DocsUrl"];

            /*if (formData["ClientDocs"] == "ClientDocs")
            {*/
               // var path = HttpRuntime.AppDomainAppPath;
            directoryName = $"{ConfigurationManager.AppSettings["DBLocation"]}{fileName}";
            // directoryName = System.IO.Path.Combine(path, "ClientDocument");
            // filename = System.IO.Path.Combine(directoryName, fileName);
            filename = directoryName+"\\"+fileName+".xlsx";

                //Detecting exists file  
            if(File.Exists(filename))
                {
                   return Request.CreateResponse<string>("File Already Exists");
                }

                string DocsPath = tempDocUrl + "/" + "ClientDocument" + "/";
                URL = DocsPath + fileName;

            //}

            /*using (Stream excelFile = File.OpenWrite(filename))
            {  
                input.CopyTo(excelFile);
                //close file  
                excelFile.Close();
            }*/

            Directory.CreateDirectory(@directoryName);   
            using (var fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write))
             {
                input.CopyTo(fileStream);
             }

            var response = Request.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("DocsUrl", URL);
                //return response;


            log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Importing New File{filename}"), EventLogEntryType.Information);
            return Request.CreateResponse<string>("Success");
            }catch (Exception ex)
            {
            log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, "Importing New File Error: "+ ex.Message), EventLogEntryType.Error);
            return Request.CreateResponse<string>(ex.Message);
            }
        }


        [HttpPost]
        [Route("/ExcelHandler/HandleSaveFile")]
        public async Task<HttpResponseMessage> HandleSaveFile([FromBody] Dictionary<string, string[]> ExcelData )
        {
            try
            {
                var sd = ExcelData;
                string fileName = ExcelData["FileName"][0];
                string filePath = $"{ConfigurationManager.AppSettings["DBLocation"]}{fileName}\\{fileName}.xlsx";

                Excel.Application app = new Excel.Application();
                Excel.Workbook xlWorkbook = app.Workbooks.Open(filePath);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

                xlWorksheet.Rows.Clear();
                xlWorksheet.Columns.Clear();
                var coll = xlWorksheet.Columns;
                int rowNo = 0;
                //int col = 1;
                foreach (string row in ExcelData.Keys)
                {
                    //col = 1;
                    if (row != "FileName")
                    {
                        /*DataColumn dc = new DataColumn();
                        dc.ColumnName = result[row];
                        dt.Columns.Add(dc);*/
                        for (int col = 1; col <= ExcelData[row].Length; col++)
                        {
                            xlWorksheet.Cells[rowNo, col].Value = ExcelData[row][col - 1];
                        }
                    }
                    rowNo++;
                    // MessageBox.Show(result[row][0]);
                }

                //app.DisplayAlerts = false;
                string tempPath = $"{ConfigurationManager.AppSettings["DBLocation"]}{fileName}\\{fileName}.xlsx";
                //  if (File.Exists(filePath)) File.Delete(filePath);
                // xlWorkbook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

                //xlWorkbook.Close();
               // xlWorkbook.SaveCopyAs(filePath);
                xlWorkbook.Save();
               // app.Save();
                xlWorkbook.Close();
                app.Quit();

                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Saving Excel Data FileName:{fileName}"), EventLogEntryType.Information);
                return Request.CreateResponse<string>("Data Saved Successfully");
            }catch (Exception ex)
            {
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Saving Excel Data FileName:{ExcelData["FileName"][0]} Error: "+ ex.Message), EventLogEntryType.Error);
                return Request.CreateResponse<string>(ex.ToString());
            }
        }


        [HttpPost]
        [Route("/ExcelHandler/HandleDeleteFile")]
        [EnableCors(origins: "http://127.0.0.1:5500", headers: "*", methods: "*")]
        public async Task<HttpResponseMessage> HandleDeleteFile([FromBody] string fileName)
        {
            try
            {
                string filePath = $"{ConfigurationManager.AppSettings["DBLocation"]}\\{fileName}"; 
                if (Directory.Exists(filePath))
                {
                    Directory.Delete(filePath, true);
                    return Request.CreateResponse<string>("Deleted");
                }

                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Deleting File(Web-request) FileName:{fileName}" ), EventLogEntryType.Information);
                return Request.CreateResponse<string>("Success");
            }
            catch (Exception ex)
            {
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Deleting File(Web-request) FileName:{fileName} , Error: "+ ex.Message), EventLogEntryType.Error);
                return Request.CreateResponse<string>(ex.ToString());
            }
        }


        [HttpPost]
        [Route("/ExcelHandler/HandleClientLogs")]
        [EnableCors(origins: "http://127.0.0.1:5500", headers: "*", methods: "*")]
        public async Task<HttpResponseMessage> HandleClientLogs([FromBody] string fileName)
        {
            try
            {
                var errStr = Request.Content.ReadAsStringAsync().Result;
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, errStr.ToString()), EventLogEntryType.Error);
                return Request.CreateResponse<string>("Success");
            }
            catch (Exception ex)
            {
                log.WriteLog( String.Format("{0} @ {1}", DateTime.Now, $"Logging ClientSideError , Error: " + ex.Message), EventLogEntryType.Error);
                return Request.CreateResponse<string>(ex.ToString());
            }
        }


        internal class ColData
        {
            public string colName { get; set; }
            public Boolean ReadOnly { get; set; }
            public Boolean Hidden { get; set; }

            public Dictionary<string, string[]> ExcelData { get; set; }


        }
    }
};