using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace WriteGoogleSheet
{
    class Program
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Update Google Sheet Data with Google Sheets API v4";
        static String spreadsheetId = "1xFHgjLLInDenc3g1UpydfGET6nViTcsgu0v7Rr8fl7k";
        static string sheetName = "Лист1";

        static void Main(string[] args)
        {
            var service = OpenSheet();
            while (true)
            {
                UpdateRow(service);
                System.Threading.Thread.Sleep(10000);
            }
        }

        static SheetsService OpenSheet()
        {
            UserCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Path.Combine
                    (System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal),
                     ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        static void UpdateRow(SheetsService service)
        {
            ValueRange rVR;
            String sRange;
            int rowNumber = 1;

            sRange = String.Format("{0}!A:A", sheetName);
            SpreadsheetsResource.ValuesResource.GetRequest getRequest
                = service.Spreadsheets.Values.Get(spreadsheetId, sRange);
            rVR = getRequest.Execute();
            IList<IList<Object>> values = rVR.Values;
            
            if (values != null && values.Count > 0) rowNumber = values.Count + 1;
            sRange = String.Format("{0}!A{1}:B{1}", sheetName, rowNumber);

            ValueRange valueRange = new ValueRange();
            valueRange.Range = sRange;
            valueRange.MajorDimension = "ROWS";

            DateTime dt = new DateTime();
            dt = DateTime.Now;
            List<object> oblist = new List<object>() { String.Format("{0}", rowNumber), dt.ToString("HH:mm:ss") };
            valueRange.Values = new List<IList<object>> { oblist };
            Console.WriteLine("{0}, {1}", oblist[0], oblist[1]);
            
            SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest
                = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, sRange);
            updateRequest.ValueInputOption
                = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            UpdateValuesResponse uUVR = updateRequest.Execute();
        }
    }
}