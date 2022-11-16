using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Activities;
using Microsoft.Office.Interop.Excel;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using static Google.Apis.Sheets.v4.SheetsService;
using System.IO;

namespace SetFormula
{
    //Simple Formula is activity name
    public class SimpleFormula : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> InputStr { get; set; }

        [Category("Output")]
        [RequiredArgument]
        public OutArgument<string> OutputStr { get; set; }

        protected override void Execute(CodeActivityContext context)
        {











            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            object oMissiong = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Open("C:\\Users\\long.ming\\Desktop\\Test.xlsx", oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);

            Worksheet Sheet = workbook.Worksheets["Test"];
            Sheet.Range["C1"].Formula = "=sum(A1:B1)";
            Sheet.Range["C2"].Formula = "=sum(A2:B2)";
            Sheet.Range["C3"].Formula = "=sum(A3:B3)";

            workbook.Close(false, oMissiong, oMissiong);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            workbook = null;
            app.Workbooks.Close();
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            app = null;
            OutputStr.Set(context, "Successfully.");

        }


    }



    public class TestCode
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Dot Tutorials";
        static readonly string sheet = "On-Hold Deals Tracker - Master/TOps";
        static readonly string SpreadsheetId = "1mzzPbzucnhlHK9eC_GejYb-FL4HA_BXrRwJxaZoZ5Kw";


        public void TesCodeA()
        {

            //SpreadsheetApp spp = new SpreadsheetApp();



            /**
            SpreadsheetsResource.ValuesResource.GetRequest request =
           service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            // Getting all records from Column A to E...
            IList<IList<object>> values = response.Values;
            **/

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            object oMissiong = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Open("C:\\Users\\long.ming\\Desktop\\Test.xlsx", oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);

            Worksheet Sheet = workbook.Worksheets["Test"];
            Sheet.Range["C1"].Formula = "=sum(A1:B1)";
            Sheet.Range["C2"].Formula = "=sum(A2:B2)";
            Sheet.Range["C3"].Formula = "=sum(A3:B3)";

            workbook.Close(false, oMissiong, oMissiong);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            workbook = null;
            app.Workbooks.Close();
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            app = null;



        }

        public void HideGoogleSheet()
        {
            try
            {

                string[] Scopes = { SheetsService.Scope.Spreadsheets };
                string tagSheetName = "Test";
                string SpreadsheetId = "1z6bo0GWapjIeBelTS4U-RL24WvStErZJXrGJdcOzas0";
                string ApplicationName = "HideSheet";
                string secretPath = "C:\\lming\\GSuiteSecretkey.json";
                int sheetId = 0;

                GoogleCredential credential;
                //Reading Credentials File...
                using (var stream = new FileStream(secretPath, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(Scopes);
                }


                SheetsService service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                var sheets = service.Spreadsheets.Get(SpreadsheetId);
                var sheetsResponse = sheets.Execute();


                foreach (Google.Apis.Sheets.v4.Data.Sheet item in sheetsResponse.Sheets)
                {
                    if (item.Properties.Title == tagSheetName)
                    {
                        sheetId = (int)item.Properties.SheetId;
                        Console.Write("SheetName: " + item.Properties.Title);
                        break;
                    }
                }
                if (sheetId > 0)
                {
                    BatchUpdateSpreadsheetRequest busRequest = new BatchUpdateSpreadsheetRequest();
                    var request = new Request()
                    {
                        UpdateSheetProperties = new UpdateSheetPropertiesRequest
                        {
                            Properties = new SheetProperties()
                            {
                                Hidden = true,
                                SheetId = sheetId
                            },
                            Fields = "Hidden"
                        }
                    };
                    busRequest.Requests = new List<Request>();
                    busRequest.Requests.Add(request);
                    var bur = service.Spreadsheets.BatchUpdate(busRequest, SpreadsheetId);
                    bur.Execute();
                }

            }
            catch (Exception ex)
            {
                string es = ex.Message;
                throw ex;
            }

        }


        public void HideDemo()
        {
            try
            {

                string[] Scopes = { SheetsService.Scope.Spreadsheets };
                string SpreadsheetId = "1z6bo0GWapjIeBelTS4U-RL24WvStErZJXrGJdcOzas0";
                string ApplicationName = "HideSheet";
                string secretPath = "C:\\lming\\GSuiteSecretkey.json";
                int sheetId = 1212121212;

                GoogleCredential credential;
                //Reading Credentials File...
                using (var stream = new FileStream(secretPath, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(Scopes);
                }


                SheetsService service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                var sheets = service.Spreadsheets.Get(SpreadsheetId);
                var sheetsResponse = sheets.Execute();

                BatchUpdateSpreadsheetRequest busRequest = new BatchUpdateSpreadsheetRequest();
                var request = new Request()
                {
                    UpdateSheetProperties = new UpdateSheetPropertiesRequest
                    {
                        Properties = new SheetProperties()
                        {
                            //Hidden = true,
                            Title = "new sheet name",
                            SheetId = sheetId
                        },
                        Fields = "Title"
                    }
                };
                busRequest.Requests = new List<Request>();
                busRequest.Requests.Add(request);
                var bur = service.Spreadsheets.BatchUpdate(busRequest, SpreadsheetId);
                bur.Execute();


            }
            catch (Exception ex)
            {
                string es = ex.Message;
                throw ex;
            }

        }
    }


}
