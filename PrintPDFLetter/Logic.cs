using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing.Printing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using LSEXT;
using LSSERVICEPROVIDERLib;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;
using Patholab_DAL_V1;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using System.IO;
using System.Windows.Forms;
using Patholab_Common;
using System.Diagnostics;
using ADODB;

//docx
//using Novacode;
using Spire.Doc;

using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using Spire.Doc.Documents;
using Spire.Doc.Fields;


namespace PrintPDFLetter
{
    class Logic
    {
        private LSExtensionParameters Parameters;
        private double sessionId;

        private string _connectionString;

        private OracleConnection _connection;

        private string sdgName;
        private long sdgId;
        private string computerName;
        private string printerName;
        private string PDFDirectory;
        private PHRASE_HEADER SystemParams;





        INautilusServiceProvider sp;
        private DataLayer dal;
        private const string Type = "1";
        private bool debug;
        private dynamic rs;
        private string _printerName;
        private long _workstationId;
        private long _operatorId;
        private SDG sdg = null;
        private string _sdg_status;
        private bool _printFlag=false;
        private string _printFileName="";
        private string _ghostScriptPath;
        private int _copies=1;
        private string _ghostscriptArguments;

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);


        public Logic(LSExtensionParameters _Parameters)
        {
            try
            {

                this.Parameters = _Parameters;

                #region params

                string tableName = Parameters["TABLE_NAME"];
                string role = Parameters["ROLE_NAME"];


                debug = (role.ToUpper() == "DEBUG");
                    if (debug) Debugger.Launch();
                  //    Debugger.Launch();
                sp = Parameters["SERVICE_PROVIDER"];
                rs = Parameters["RECORDS"];
                rs.MoveLast();
                _workstationId = (long) Parameters["WORKSTATION_ID"];
                _operatorId = (long) Utils.GetNautilusUser(sp).GetOperatorId();


                #endregion

                ////////////יוצר קונקשן//////////////////////////
                var ntlCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlCon);
                /////////////////////////////           

                _connection = GetConnection(ntlCon);
                dal = new DataLayer();
                dal.Connect(ntlCon);
                WORKSTATION workstation =
                    dal.FindBy<WORKSTATION>(w => w.WORKSTATION_ID == _workstationId).SingleOrDefault();
                if (workstation != null)
                {
                    _printerName = workstation.WORKSTATION_USER.U_PRINTER_NAME ?? "";
                }
                letterCreated = false;

                debug = (role.ToUpper() == "DEBUG");
              //  if (debug) Debugger.Launch();
                if (tableName == "SDG")
                {
                    sdgName = (string) rs.Fields["NAME"].Value;
                    sdgId = (long) rs.Fields["SDG_ID"].Value;
                    sdg = dal.FindBy<SDG>(d => d.SDG_ID == sdgId).FirstOrDefault();
                    _sdg_status = "'" + sdg.STATUS + "'";
                    //_ghostScriptPath = ConfigurationManager.AppSettings["ghostscriptgswin32cFullPath"];
                    //_ghostscriptArguments = ConfigurationManager.AppSettings["ghostscriptArguments"];
                    string assemblyPath = Assembly.GetExecutingAssembly().Location;
                    ExeConfigurationFileMap map = new ExeConfigurationFileMap();
                    map.ExeConfigFilename = assemblyPath + ".config";
                    Configuration cfg = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                    AppSettingsSection appSettings = cfg.AppSettings;
                    _ghostScriptPath = appSettings.Settings["ghostscriptgswin32cFullPath"].Value;
                    _ghostscriptArguments = appSettings.Settings["ghostscriptArguments"].Value;

                    //create letter only if first print on authorise or if not yet authorised
                    if (sdg.STATUS != "A" || sdg.SDG_USER.U_PDF_PATH == null)
                    {
                        //initLetterName();
                        _printFlag = false;
                        _printFileName = "";
                    //   _copies =_copies?? 1;
                      
                        CreateLetter();
                        if (_printFlag && _printFileName != "")
                        {

                            int i = 100;
                            while (!File.Exists(_printFileName) && i-- > 0)
                            {
                                Thread.Sleep(300);
                            }
                            if (i > 0)
                            {
                                try
                                {
                                    Thread.Sleep(1000);
                                    PrintPdfFile(_printFileName, _copies);
                                }


                                catch (Exception ex)
                                {
                                    if (debug) Logger.WriteLogFile(ex);
                                }
                            }
                            else
                            {
                                MessageBox.Show(
                                    "Cannot Find the pdf file in location after 30 seconds. The process letter must be down");
                               if (debug) if (debug) Logger.WriteLogFile(
                                    new Exception(
                                        "cannot Find the pdf file in location after 30 seconds. the process letter must be down"));
                            }
                        }
                    }
                    else
                    {
                        try
                        {

                            PrintPdfFile(sdg.SDG_USER.U_PDF_PATH, 1);
                        }


                        catch (Exception ex)
                        {
                            if (debug) Logger.WriteLogFile(ex);
                        }
                    }
                }
                else
                {
                    if (debug) MessageBox.Show("זה לא SDG");
                }

            }

            catch (Exception ex)
            {
                if (debug) MessageBox.Show("תקלה ביצירת מכתב");
                if (debug) Logger.WriteLogFile(ex);

            }
            finally
            {
                dal.Close();
                dal = null;
            }
        }

        private void PrintPdfFileOld(string printFileName,int numberOfCopies)
        {
            Process p = new Process();
            p.StartInfo = new ProcessStartInfo()
                {
                    CreateNoWindow = true,
                    Verb = "Print",
                    FileName = printFileName //put the correct path here
                };
            p.Start();
        }
        public bool PrintPdfFile(string pdfFileName, int numberOfCopies)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            //startInfo.Arguments = " -dPrinted -dBATCH -dNOPAUSE -dNOSAFER -q -dNumCopies=" + Convert.ToString(numberOfCopies) + " -sDEVICE=ljet4 -sOutputFile=\"\\\\spool\\" + printerName + "\" \"" + pdfFileName + "\" ";
            PrinterSettings settings = new PrinterSettings();
           string localPrinterName=settings.PrinterName;
            startInfo.Arguments = string.Format(_ghostscriptArguments, Convert.ToString(numberOfCopies),localPrinterName, pdfFileName);
            startInfo.FileName = _ghostScriptPath;
            //startInfo.UseShellExecute = false;

            //startInfo.RedirectStandardError = true;
            //startInfo.RedirectStandardOutput = true;

            Process process = Process.Start(startInfo);

          //  Console.WriteLine(process.StandardError.ReadToEnd() + process.StandardOutput.ReadToEnd());

            process.WaitForExit(30000);
            //if (process.HasExited == false) process.Kill();


            //return process.ExitCode == 0;
            return true;
        }
        private bool CreateLetter()
        {
       
            int workflowNodeId = Parameters["WORKFLOW_NODE_ID"];
            WORKFLOW_NODE extentionNode = dal.FindBy<WORKFLOW_NODE>(wn => wn.WORKFLOW_NODE_ID == workflowNodeId).SingleOrDefault();

            //U_WREPORT[] wreports = dal.FindBy<U_WREPORT>(wr => (";" + wr.U_WREPORT_USER.U_WORKFLOW_EVENT + ";").Contains(";" + extentionNode.PARENT_NODE.NAME + ";")
            //    && wr.U_WREPORT_USER.U_WORKFLOW_NAME == extentionNode.WORKFLOW.NAME).ToArray();

            U_WREPORT[] wreports =
                dal.FindBy<U_WREPORT>(wr => wr.U_WREPORT_USER.U_WORKFLOW_EVENT == extentionNode.PARENT_NODE.NAME).ToArray();
            if (extentionNode.PARENT_NODE.NAME == "A" || extentionNode.PARENT_NODE.NAME == "Authorised" || extentionNode.PARENT_NODE.NAME == "ToAuthorise")
            {
                wreports =
                dal.FindBy<U_WREPORT>(wr => wr.U_WREPORT_USER.U_WORKFLOW_EVENT == "Print PDF Letter").ToArray();
                _sdg_status = "'A'";

            }
            foreach (U_WREPORT wreport in wreports)
            {
                if (!(";" + wreport.U_WREPORT_USER.U_WORKFLOW_NAME + ";").Contains(";" + extentionNode.WORKFLOW.NAME + ";"))
               {
                   continue;
               }
                //  if (debug) MessageBox.Show("Error: the wreport '" + letterName.ToString() + "' was no found");
                // return false;

                if (wreport.U_WRDESTINATION_USER == null)
                {
                    if (debug)
                    {
                        MessageBox.Show("Error:Can not find the print/save Destination for wreport '" +
                                         letterName.ToString() + "' ");
                        Logger.WriteQueries("Error:Can not find the print/save Destination for wreport ");

                    }
                    return false;
                }
                string docLocation = wreport.U_WREPORT_USER.U_WORD_TEMPLATE;
                if (debug) Logger.MyLog("pre run docLocation == " + docLocation);
                docLocation = ExecuteOrGetString(docLocation);
                
                if (docLocation == null)
                {
                    if (debug)
                        MessageBox.Show("Error: the doc file location(U_WORD_TEMPLATE) is missing for wreport '" +
                                        letterName + "'");
                    Logger.WriteQueries("Error: the doc file location(U_WORD_TEMPLATE) is missing for wreport '" +
                                        letterName + "'");

                    return false;
                }
                if (debug) Logger.MyLog("docLocation == " +docLocation );
                Spire.Doc.Document document = new Spire.Doc.Document();
                document.LoadFromFile(docLocation);
                if (document == null )
                {
                    if (debug)
                        MessageBox.Show("Error loading the doc '" + docLocation + "' for the wreport '" + letterName +
                                        "'");
                    return false;
                }

                U_WREPORT_QUERY_USER[] queries = wreport.U_WREPORT_QUERY_USER.ToArray();

                //U_WREPORT_QUERY_USER[] queries = dal.FindBy<U_WREPORT_QUERY_USER>(qu => qu.U_WREPORT_ID == wreport.U_WREPORT_ID).ToArray();
                foreach (U_WREPORT_QUERY_USER query in queries)
                {
                    Logger.WriteQueries("start on " + query.U_QUERY_NAME);
                    //todo: run sql query here and return the data as column name and val
                    //   var queryResults = dal.RunQuery( query.U_QUERY.Replace("#SDG_ID",sdgId.ToString());
                    //string queryString = query.U_QUERY.Replace("#SDG_ID#", sdgId.ToString());
                    string queryString = Regex.Replace(query.U_QUERY, "#SDG_ID#", sdgId.ToString(),
                                                       RegexOptions.IgnoreCase);
                    queryString = Regex.Replace(queryString, "#SDG_STATUS#", _sdg_status, RegexOptions.IgnoreCase);

                    queryString = Regex.Replace(queryString, "#OPERATOR_ID#", _operatorId.ToString(), RegexOptions.IgnoreCase);
                    queryString = Regex.Replace(queryString, "#WORKSTATION_ID#", _workstationId.ToString(), RegexOptions.IgnoreCase);
                    queryString = Regex.Replace(queryString, "#PRINTER_NAME#", _printerName, RegexOptions.IgnoreCase);
                    //run the query and replace the text
                    OracleDataReader reader = RunQuery(queryString);
                    // Run query in U_QUERY, 


                    if (reader == null || !reader.HasRows)
                    {

                        //if no resulst
                        if (debug)
                        {
                            MessageBox.Show("Query  for wreport '" + wreport.NAME + "' query_name: '" +
                                          query.U_QUERY_NAME + "' returned null");
                            Logger.WriteQueries("Query  for wreport '" + wreport.NAME + "' query_name: '" +
                                                 query.U_QUERY_NAME + "' returned null");
                        }

                        try
                        {


                            string matchString = query.U_QUERY_NAME;// +"_" + columnName;

                            //      Debugger.Launch();
                            var fs = document.FindAllString(matchString, false, false);


                            if (fs == null) continue;
                            var list = fs.GetEnumerator();


                            while (list.MoveNext())
                            {
                                TextSelection textSelection = list.Current as TextSelection;


                                TextRange fullText = textSelection.GetPrivatePropertyValue<TextRange>("EndTextRange");
                              

                                document.Replace(fullText.Text, "", false, true);

                            }
                        }
                        catch (Exception ex)
                        {

                            Logger_WriteLogFile(ex);
                        }
                    }
                    else
                    {
                        //if results found, replace in doc
                        string columnName = "";
                        string columnValue;
                        //run replace the text  
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            try
                            {
                                columnName = reader.GetName(i);
                                columnValue = reader.GetValue(i).ToString();
                                string matchString = query.U_QUERY_NAME + "_" + columnName;
                                string newString = columnValue;

                                // document.ReplaceText(query.U_QUERY_NAME + "_" + columnName, columnValue, false, RegexOptions.IgnoreCase,);
                                // document.FindAllString(replaceString, false, true).FirstOrDefault().;


                                //replace query in "~^ Select... from... where... ^~"

                                int queryStart;
                                int queryEnd;
                                string replacWith;
                                while ((queryStart = newString.IndexOf("~^")) >= 0)
                                {
                                    replacWith = "";
                                    if ((queryEnd = newString.IndexOf("~^", queryStart + 2)) > queryStart)
                                    {
                                        //~^~^                       ~^ abc ~^
                                        //0123 substring(0+2,2-0-2)  01 234 56 substring(0+2,5-0-2)
                                        string innerQuery = newString.Substring(queryStart + 2,
                                                                                queryEnd - queryStart - 2);
                                        try
                                        {
                                            OracleDataReader reader2 = RunQuery(innerQuery);

                                            if (reader2 != null && reader2.HasRows)
                                            {
                                                replacWith = reader2.GetValue(0).ToString();
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            if (debug) Logger.WriteLogFile(ex);
                                            Logger.WriteQueries(innerQuery);

                                            replacWith = ":)";
                                        }
                                    }
                                    else
                                    {
                                        //if ^~ not found delete only the ~^
                                        queryEnd = queryStart;
                                    }
                                    newString = newString.Substring(0, queryStart)
                                                + replacWith
                                                + newString.Substring(queryEnd + 2);
                                }
                                if (newString.ToUpper().Contains("RTF"))
                                {
                                  
                                    foreach (
                                        TextSelection foundSelection in document.FindAllString(matchString, false, true)
                                        )
                                    {
                                        Spire.Doc.Fields.TextRange range = foundSelection.GetAsOneRange();
                                        //replace the \par with a paragraph placeholder becaus there is a bug in spire appendRtf
                                        newString = newString.Replace(@"\par ", "<~BR~>");

                                        //deal with hebrew RightToLeft RTL RTF text in english LTR Location 
                                        //if (newString.Contains(@"\rtlpar") && !newString.Contains(@"\ltrpar")
                                        //    ||( (newString.Contains(@"\rtlpar") && newString.Contains(@"\ltrpar")
                                        //        && newString.IndexOf(@"\rtlpar") < newString.IndexOf(@"\ltrpar"))
                                        //       ))
                                        //{
                                        //    range.OwnerParagraph.Format.IsBidi = true;
                                        //}
                                        //else if (newString.Contains(@"\ltrpar") && !newString.Contains(@"\rtlpar")
                                        //    || ((newString.Contains(@"\rtlpar") && newString.Contains(@"\ltrpar")
                                        //        && newString.IndexOf(@"\rtlpar") > newString.IndexOf(@"\ltrpar"))
                                        //       ))
                                        //{
                                        //    range.OwnerParagraph.Format.IsBidi = false;
                                        //}
                                        range.OwnerParagraph.AppendRTF( newString + "");

                                        //    }
                                        //document.SaveToFile(@"C:\Q1.docx", FileFormat.Docx);
                                    }
                                    document.Replace

                                        (matchString , "", false, false);

                                    // document.Replace(matchString,rtfDoc.Clone(),false,true);

                                }
                                else
                                {
                                    //newString is not rtf 
                                    document.Replace(matchString, newString, false, true);

                                }
                            }
                            catch (Exception ex)
                            {
                                if (debug) Logger.WriteLogFile(ex);
                                Logger.WriteQueries("Query  for wreport '" + wreport.NAME + "' query_name: '" +
                                                         query.U_QUERY_NAME + "'\n columnName:'" + columnName.ToString() +
                                                         "' returned  error");
                                if (debug)
                                    MessageBox.Show("Query  for wreport '" + wreport.NAME + "' query_name: '" +
                                                    query.U_QUERY_NAME + "'\n columnName:'" + columnName.ToString() +
                                                    "' returned  error");

                                //continue loop 
                            }
                        }

                        // Call Close when done reading.

                        reader.Close();

                    }


                }
                //replace the paragraph placeholder
                document.Replace("<~BR~>", "\n", false, false);
                //save the report 

                Save(document, wreport.U_WRDESTINATION_USER.ToArray());
            }
            return true;
        }

        private OracleDataReader RunQuery(string queryString)
        {
            OracleCommand cmd = new OracleCommand(queryString, _connection);
            OracleDataReader reader;
            try
            {
                reader = cmd.ExecuteReader();
                reader.Read();
            }
            catch (Exception ex)
            {
                if (debug) Logger.WriteLogFile(ex);
                Logger.WriteQueries("reader is null " + queryString);
                reader = null;
                //continue loop 
            }
            return reader;
        }
        public bool Save(Spire.Doc.Document document, U_WRDESTINATION_USER[] destinations)
        {
            //using workstationId, sdg_id


            string copiesCSV = "";
            string saveAsCSV = "";
            string deviceName;
            string type;
            string typeCSV = "";
            string wrDestinationIdCSV = "";
            string wreportId = "";

            try
            {
                if (!GetDefaultDirectoiesFromPhrase())
                {
                    if (debug) MessageBox.Show(@"Error: Could not find entry ""PDF Directory"" in Phrase ""System Parameters""");
                    return false;
                }
                foreach (U_WRDESTINATION_USER destination in destinations)
                {
                    type = ExecuteOrGetString(destination.U_TYPE);
                    if (type.ToUpper() == "PROMPTED FAX")
                    {
                        Fax_Prompt faxForm = new Fax_Prompt();
                         faxForm.ShowDialog();
                        deviceName = faxForm.PhoneNumber;
                        type = "FAX";
                        if (deviceName == "-1")
                        {
                            deviceName = "";
                            type = "";
                        }
                    }
                    else deviceName = ExecuteOrGetString(destination.U_DEVICE_NAME);
                    if (type == "" || deviceName == "") continue;
                    if (type.ToUpper() == "PRINT")
                    {
                        _copies = ExecuteOrGetInt(destination.U_COPIES);
                        copiesCSV += _copies.ToString() + ";";

                        _printFlag = true;
                    }
                    else
                    {
                        copiesCSV += ExecuteOrGetInt(destination.U_COPIES).ToString() + ";";
                        
                    }
                    if (debug)
                    {
                        MessageBox.Show("Adding '" + type + "' destination :'" + destination.U_WRDESTINATION.NAME + "'  for device '" + deviceName + "'");
                    }
                    typeCSV += type.ToUpper() + ";";
                    wrDestinationIdCSV += destination.U_WRDESTINATION_ID.ToString() + ";";
                    wreportId = destination.U_WREPORT_ID.ToString();
                    

                    //saveas should contain the full path of the final file/ printer name / fax number 
                    saveAsCSV += deviceName.Replace(";", "|") + ";";


                    //DateTime now = DateTime.Now;
                    //string yearMonth = now.Year.ToString() + @"\" + now.Month.ToString("MM");
                }

                if (typeCSV != "")
                {
                    //make sure directory ends with a \
                    if (!PDFDirectory.EndsWith(@"\")) PDFDirectory = @"\";
                    Directory.CreateDirectory(PDFDirectory);

                    document.Variables["sdgId"] = sdgId.ToString();
                    document.Variables["wreportId"] = wreportId;
                    document.Variables["typeCSV"] = typeCSV;
                    document.Variables["saveAsCSV"] = saveAsCSV;
                    document.Variables["copiesCSV"] = copiesCSV;
                    document.Variables["wrDestinationIdCSV"] = wrDestinationIdCSV;

                    SDG_USER sdgUser = dal.FindBy<SDG_USER>(du => du.SDG_ID == sdgId).SingleOrDefault();
                    if (sdgUser.U_PATHOLAB_NUMBER != null)
                    {
                        _printFileName = PDFDirectory + @"\Backup\" + DateTime.Now.ToString(@"yyyy\\MM\\") +
                                        MakeValidFileName(sdgUser.U_PATHOLAB_NUMBER.Replace("/", "-")) + "-" +
                           wreportId + ".PDF";

                        try
                        {
                            if (File.Exists(_printFileName))
                            {
                                File.Delete(_printFileName);
                            }
                        }
                        catch (Exception)
                        {
                        }
                        document.SaveToFile(
                            PDFDirectory + MakeValidFileName(sdgUser.U_PATHOLAB_NUMBER.Replace("/", "-")) + "-" +
                            wreportId + ".docx", FileFormat.Docx);
                       
                    }
                    else
                    {
                        document.SaveToFile(PDFDirectory + sdgUser.SDG_ID.ToString() + "-" + wreportId + ".docx",
                                            FileFormat.Docx);
                        //_printFileName = PDFDirectory + @"\Backup\" + DateTime.Now.ToString(@"yyyy\\MM\\") +
                        //                 sdgUser.SDG_ID.ToString() + "-" +
                        //    wreportId + ".PDF";

                    }
                    dal.InsertToSdgLog(sdgId, "DOC.CREATE", (long)sessionId, wreportId);

                    return true;
                }
                else return false;



            }
            catch (Exception ex)
            {
                if (debug) Logger.WriteLogFile(ex);
                return false;
            }


        }
        private string MakeValidFileName(string name)
        {
            string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
            string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);

            return System.Text.RegularExpressions.Regex.Replace(name, invalidRegStr, "_");
        }
        private int ExecuteOrGetInt(string queryOrString)
        {
            //run a query and 
            int result;
            result = 1;

            if (queryOrString == null)
            {
                //default result to 1 
                result = 1;
            }
            //here in the tryparse the result are drawn from the string if it is numeric
            else if (!int.TryParse(queryOrString.ToString(), out result))
            {
                // if parse failed, see if it is a select query)

                if (queryOrString.IndexOf("select", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string query = Regex.Replace(queryOrString, "#SDG_ID#", sdgId.ToString(), RegexOptions.IgnoreCase);
                    query = Regex.Replace(query, "#SDG_STATUS#", _sdg_status, RegexOptions.IgnoreCase);
                    query = Regex.Replace(query, "#OPERATOR_ID#", _operatorId.ToString(), RegexOptions.IgnoreCase);
                    query = Regex.Replace(query, "#WORKSTATION_ID#", _workstationId.ToString(), RegexOptions.IgnoreCase);
                    query = Regex.Replace(query, "#PRINTER_NAME#", _printerName, RegexOptions.IgnoreCase);

                    OracleDataReader reader = RunQuery(query);
                    // Run query in U_copies, 
                    if (reader == null || !reader.HasRows)
                    {
                        //if no resulst
                        if (debug) MessageBox.Show("Query for copies wrdestination \n query: '" + query + "' returned null");
                        //return;
                        result = 1;
                    }
                    else
                    {
                        if (!int.TryParse(reader.GetValue(0).ToString(), out result)) result = 1;
                    }
                }
            }
            return result;
        }

        private string ExecuteOrGetString(string queryOrString)
        {
            string result;
            result = "";

            if (queryOrString == null)
            {
                //default to don`t send/ 1 copy
                result = "";
            }
            //if ther is no select, return the string
            string query = Regex.Replace(queryOrString ?? "", "#SDG_ID#", sdgId.ToString(), RegexOptions.IgnoreCase);
            query = Regex.Replace(query, "#SDG_STATUS#", _sdg_status, RegexOptions.IgnoreCase);

            query = Regex.Replace(query, "#OPERATOR_ID#", _operatorId.ToString(), RegexOptions.IgnoreCase);
            query = Regex.Replace(query, "#WORKSTATION_ID#", _workstationId.ToString(), RegexOptions.IgnoreCase);
            query = Regex.Replace(query, "#PRINTER_NAME#", _printerName, RegexOptions.IgnoreCase);

            if (query.IndexOf("select", StringComparison.OrdinalIgnoreCase) < 0)
            {
                result = query;
            }
            else
            {

                OracleDataReader reader = RunQuery(query);
                // Run query in queryString, 
                if (reader == null || !reader.HasRows)
                {
                    //if no resulst
                    if (debug) MessageBox.Show("Query for String in  wrdestination \n query: '" + query + "' returned null");
                    //return;
                    result = "";
                }
                else
                {
                    result = reader.GetValue(0).ToString();
                }
            }

            return result;
        }

        private bool GetDefaultDirectoiesFromPhrase()
        {
            bool result;
            try
            {
                SystemParams = dal.GetPhraseByName("System Parameters");
                result =
                    //SystemParams.PhraseEntriesDictonary.TryGetValue("Print Directory", out printDirectory)
                    //&&
                   SystemParams.PhraseEntriesDictonary.TryGetValue("PDF Directory", out PDFDirectory);
                // return true;
            }
            catch (Exception ex)
            {
                if (debug) Logger.WriteLogFile(ex);
                if (debug) MessageBox.Show(@"Error: Could not find entry  ""PDF Directory"" in Phrase ""System Parameters""");
                if (debug) Logger.WriteLogFile(new Exception(@"Error: Could not find entry  ""PDF Directory"" in Phrase ""System Parameters"""));
                return false;
            }
            //try
            //{
            //    result = result &&
            //    SystemParams.PhraseEntriesDictonary.TryGetValue("Dr.Marmur", out DrMarmursId)
            //    &&
            //    SystemParams.PhraseEntriesDictonary.TryGetValue("Dr.Barzilay", out DrBarzilaysId);
            //    // return true;
            //}
            //catch
            //{
            //    if (debug) Logger.WriteLogFile(ex);
            //    if (debug) MessageBox.Show(@"Error: Could not find entry ""Dr.Marmur"" or ""Dr.Barzilay"" in Phrase ""System Parameters""");
            //    return false;
            //}

            return result;
        }
        public void Logger_WriteLogFile(Exception ex)
        {
            if (debug) Logger.WriteLogFile(ex);
        }

        OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {

            OracleConnection connection = null;

            if (ntlsCon != null)
            {


                // Initialize variables
                String roleCommand;
                // Try/Catch block
                try
                {
                    _connectionString = ntlsCon.GetADOConnectionString();

                    var splited = _connectionString.Split(';');

                    var cs = "";

                    for (int i = 1; i < splited.Count(); i++)
                    {
                        cs += splited[i] + ';';
                    }


                    //Create the connection
                    connection = new OracleConnection(cs);

                    // Open the connection
                    connection.Open();

                    // Get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    // Set role lims user
                    if (limsUserPassword == "")
                    {
                        // LIMS_USER is not password protected
                        roleCommand = "set role lims_user";
                    }
                    else
                    {
                        // LIMS_USER is password protected.
                        roleCommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    // set the Oracle user for this connecition
                    OracleCommand command = new OracleCommand(roleCommand, connection);

                    // Try/Catch block
                    try
                    {
                        // Execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        // Throw the exception
                        if (debug) Logger.WriteLogFile(ex);
                        throw new Exception("Inconsistent role Security : " + ex.Message);
                    }

                    // Get the session id
                    sessionId = ntlsCon.GetSessionId();

                    // Connect to the same session
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    // Build the command
                    command = new OracleCommand(sSql, connection);

                    // Execute the command
                    command.ExecuteNonQuery();

                }
                catch (Exception e)
                {
                    if (debug) Logger.WriteLogFile(e);

                    // Throw the exception
                    throw e;
                }
                // Return the connection
            }
            return connection;
        }

        public string letterName { get; set; }

        public object letterType { get; set; }

        public bool letterCreated { get; set; }
    }
    public static class PropertyHelper
    {
        /// <summary>
        /// Returns a _private_ Property Value from a given Object. Uses Reflection.
        /// Throws a ArgumentOutOfRangeException if the Property is not found.
        /// </summary>
        /// <typeparam name="T">Type of the Property</typeparam>
        /// <param name="obj">Object from where the Property Value is returned</param>
        /// <param name="propName">Propertyname as string.</param>
        /// <returns>PropertyValue</returns>
        public static T GetPrivatePropertyValue<T>(this object obj, string propName)
        {
            if (obj == null) throw new ArgumentNullException("obj");
            PropertyInfo pi = obj.GetType().GetProperty(propName,
                                                        BindingFlags.Public | BindingFlags.NonPublic |
                                                        BindingFlags.Instance);
            if (pi == null)
                throw new ArgumentOutOfRangeException("propName",
                                                      string.Format("Property {0} was not found in Type {1}", propName,
                                                                    obj.GetType().FullName));
            return (T)pi.GetValue(obj, null);
        }

        /// <summary>
        /// Returns a private Field Value from a given Object. Uses Reflection.
        /// Throws a ArgumentOutOfRangeException if the Property is not found.
        /// </summary>
        /// <typeparam name="T">Type of the Field</typeparam>
        /// <param name="obj">Object from where the Field Value is returned</param>
        /// <param name="propName">Field Name as string.</param>
        /// <returns>FieldValue</returns>
        public static T GetPrivateFieldValue<T>(this object obj, string propName)
        {
            if (obj == null) throw new ArgumentNullException("obj");
            Type t = obj.GetType();
            FieldInfo fi = null;
            while (fi == null && t != null)
            {
                fi = t.GetField(propName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                t = t.BaseType;
            }
            if (fi == null)
                throw new ArgumentOutOfRangeException("propName",
                                                      string.Format("Field {0} was not found in Type {1}", propName,
                                                                    obj.GetType().FullName));
            return (T)fi.GetValue(obj);
        }

        /// <summary>
        /// Sets a _private_ Property Value from a given Object. Uses Reflection.
        /// Throws a ArgumentOutOfRangeException if the Property is not found.
        /// </summary>
        /// <typeparam name="T">Type of the Property</typeparam>
        /// <param name="obj">Object from where the Property Value is set</param>
        /// <param name="propName">Propertyname as string.</param>
        /// <param name="val">Value to set.</param>
        /// <returns>PropertyValue</returns>
        public static void SetPrivatePropertyValue<T>(this object obj, string propName, T val)
        {
            Type t = obj.GetType();
            if (t.GetProperty(propName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance) == null)
                throw new ArgumentOutOfRangeException("propName",
                                                      string.Format("Property {0} was not found in Type {1}", propName,
                                                                    obj.GetType().FullName));
            t.InvokeMember(propName,
                           BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.SetProperty |
                           BindingFlags.Instance, null, obj, new object[] { val });
        }


        /// <summary>
        /// Set a private Field Value on a given Object. Uses Reflection.
        /// </summary>
        /// <typeparam name="T">Type of the Field</typeparam>
        /// <param name="obj">Object from where the Property Value is returned</param>
        /// <param name="propName">Field name as string.</param>
        /// <param name="val">the value to set</param>
        /// <exception cref="ArgumentOutOfRangeException">if the Property is not found</exception>
        public static void SetPrivateFieldValue<T>(this object obj, string propName, T val)
        {
            if (obj == null) throw new ArgumentNullException("obj");
            Type t = obj.GetType();
            FieldInfo fi = null;
            while (fi == null && t != null)
            {
                fi = t.GetField(propName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                t = t.BaseType;
            }
            if (fi == null)
                throw new ArgumentOutOfRangeException("propName",
                                                      string.Format("Field {0} was not found in Type {1}", propName,
                                                                    obj.GetType().FullName));
            fi.SetValue(obj, val);
        }
    }

}
