using ADODB;
using LSEXT;
using LSSERVICEPROVIDERLib;
using Microsoft.Win32.SafeHandles;

using Patholab_Common;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;//for debugger :)
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;


namespace PrintPDFLetter
{

    [ComVisible(true)]
    [ProgId("PrintPDFLetter.PrintPDFLetterCls")]

    public class PrintPDFLetterCls : IWorkflowExtension
    {
        private Logic logic;

        public void Execute(ref LSExtensionParameters Parameters)
        {
            bool debug = false;
            Logger.WriteLogFile("before   logic = new Logic(Parameters);");
            logic = new Logic(Parameters);
            Logger.WriteLogFile("after   logic = new Logic(Parameters);");

            try
            {

                bool letterCreated = logic.letterCreated;

                Logger.WriteLogFile("logic.letterCreated is " + logic.letterCreated);

                string tableName = Parameters["TABLE_NAME"];
                //string role = Parameters["ROLE_NAME"];



                //    debug = (role.ToUpper() == "DEBUG");
                //MessageBox.Show(name);
                //      var rs = Parameters["RECORDS"];

                #region SDG

                if (tableName == "SDG")
                {

                    //   rs.MoveLast();
                    //  string SdgIdSring = rs.Fields["SDG_ID"].Value.ToString();
                    //string workstationId = Parameters["WORKSTATION_ID"].ToString();
                    //long sdgId = long.Parse(SdgIdSring);
                    //string sdgStatus = rs.Fields["STATUS"].Value.ToString();


                    //run the letter create 
                    if (!letterCreated)
                    {
                        if (debug) MessageBox.Show("לא נוצר מכתב");
                    }
                    else
                    {

                        if (debug) MessageBox.Show("המכתב נוצר בהצלחה");

                    }
                }
                else
                {

                    if (debug) MessageBox.Show("התוכנית עובדת על SDG בלבד");



                }

                #endregion
            }
            catch (Exception ex)
            {
                if (debug) MessageBox.Show("תקלה ביצירת מכתב");
                logic.Logger_WriteLogFile(ex);
            }
            finally
            {
                logic.CloseDal();
            }
        }







    }
}
