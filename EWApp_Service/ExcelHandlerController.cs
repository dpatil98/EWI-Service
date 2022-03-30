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
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Collections.Specialized;
using System.Web.Http.Cors;
using ExcelDataReader;

namespace EWApp_Service
{
    public class ExcelHandlerController : ApiController
    {
        
        //creating object of datatable present on clientside
        static private System.Data.DataTable dt = new System.Data.DataTable();
        static XElement xelement;

        public String GetString(Int32 id)
        {
            return "This String";
        }

        [EnableCors(origins: "http://127.0.0.1:5500", headers: "*", methods: "*")]
        public string GetAllFiles()
        {
            List<string> files = new List<string>();  
            Dictionary<string, string[]> rawJsonData = new Dictionary<string, string[]>();
            string[] allfiles = Directory.GetDirectories("D:\\EW-WEB\\EWApp\\bin\\AllFiles\\", "*", SearchOption.AllDirectories);

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
            return json;
        }

        public String GetExcelData(string fileName)
        {
            string filePath = $@"D:\EW-WEB\EWApp\bin\AllFiles\\{fileName}\{fileName}.xlsx";
           
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

            string XMLdirectory = "D:\\EW-WEB\\EWApp\\bin\\AllFiles\\"+fileName+"\\"+fileName+"_setting.xml";

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
            return json;
        }


       

        [HttpGet]
        public String NewGetExcelData(string fileName)
        {
           // Dictionary<string, System.Data.DataTable > XlData = new Dictionary<string, System.Data.DataTable>();
          //  Dictionary<string, string[][]> XlData = new Dictionary<string, string[][]>();
            string filePath = $@"D:\EW-WEB\EWApp\bin\AllFiles\\{fileName}\{fileName}.xlsx";
         
            DataSet result = null;
           // List<string[][]> ts = new List<string[][]>();

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
            string json = JsonConvert.SerializeObject(result.Tables[0], Newtonsoft.Json.Formatting.Indented);


            string[] columnNames = (from string x in dt.Rows[0].ItemArray select x.ToString()).ToArray();
                                    

                //string[] columnNames = (string[])dt.Rows[0].ItemArray[0];
             /* foreach (System.Data.DataTable table in result.Tables)
              {
                  *//*if (//my conditions)
                  {
                      continue;
                  }*//*
                  var rows = table.AsEnumerable().ToArray();

                  //var dataTable = new string[table.Rows.Count][];//[table.Rows[0].ItemArray.Length];
                  Dictionary<string, string[]> dataTable = new Dictionary<string, string[]>;
                  Parallel.For(0, rows.Length, new ParallelOptions { MaxDegreeOfParallelism = 8 },
                      i =>
                      {
                          var row = rows[i];
                          dataTable[i] = row.ItemArray.Select(x => x.ToString()).ToArray();
                      });

                  XlData.Add("XlData", dataTable);
                  //only for one table
                  break;
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

             string XMLdirectory = "D:\\EW-WEB\\EWApp\\bin\\AllFiles\\" + fileName + "\\" + fileName + "_setting.xml";

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
            
            return json;
        }


        [HttpGet]
        public string ReadXML(string fileName)
        {
            string xmlfilePath = "D:\\EW-WEB\\EWApp\\bin\\AllFiles\\"+fileName+"\\"+fileName+"_setting.xml";

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
            return json;
        }



        [HttpGet]
        public string GetFileName(string key)
        {
            string fileName=string.Empty;
            string XMLdirectory = "D:\\EW-WEB\\EWApp\\bin\\AllFiles\\Entries.xml";
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
            return fileName;
        }


        [HttpPost]
        [Route("/ExcelHandler/StoreKeyFile")]
        [EnableCors(origins: "http://127.0.0.1:5500", headers: "*", methods: "*")]
        public async Task<HttpResponseMessage> StoreKeyFile( string[] fileName)
        {
            try
            {
                var name = fileName;
              
                string XMLdirectory = "D:\\EW-WEB\\EWApp\\bin\\AllFiles\\Entries.xml";

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

                return Request.CreateResponse<string>("Saved");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse<string>(ex.ToString());
            }
            
        }


        [HttpPost]
        [Route("/ExcelHandler/SaveXML")]
        public async Task<HttpResponseMessage> SaveXML(List<Object> li)
        {
            //Console.WriteLine(li);
            string fileName= li[0].ToString();
            string xmlfilePath = "D:\\EW-WEB\\EWApp\\bin\\AllFiles\\"+fileName+"\\"+fileName+"_setting.xml";
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

            return Request.CreateResponse<string>("Saved");
        }


        [HttpPost]
        [Route("/ExcelHandler/HandleNewFile")]
        public async Task<HttpResponseMessage> HandleNewFile()
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
            directoryName = "D:\\EW-WEB\\EWApp\\bin\\AllFiles\\"+fileName;
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



            return Request.CreateResponse<string>("Success");
        }


        [HttpPost]
        [Route("/ExcelHandler/HandleSaveFile")]
        public async Task<HttpResponseMessage> HandleSaveFile([FromBody] Dictionary<string, string[]> ExcelData )
        {
            try
            {
                var sd = ExcelData;
                string fileName = ExcelData["FileName"][0];
                string filePath = $@"D:\EW-WEB\EWApp\bin\AllFiles\{fileName}\{fileName}.xlsx";

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
                string tempPath = $@"D:\EW-WEB\EWApp\bin\AllFiles\{fileName}\{fileName}.xlsx";
                //  if (File.Exists(filePath)) File.Delete(filePath);
                // xlWorkbook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

                //xlWorkbook.Close();
               // xlWorkbook.SaveCopyAs(filePath);
                xlWorkbook.Save();
               // app.Save();
                xlWorkbook.Close();
                app.Quit();


                return Request.CreateResponse<string>("Success");
            }catch (Exception ex)
            {
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
                string filePath = $@"D:\EW-WEB\EWApp\bin\AllFiles\\{fileName}"; 
                if (Directory.Exists(filePath))
                {
                    Directory.Delete(filePath, true);
                    return Request.CreateResponse<string>("Deleted");
                }

                return Request.CreateResponse<string>("Success");
            }
            catch (Exception ex)
            {
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