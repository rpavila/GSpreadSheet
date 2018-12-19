using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Newtonsoft.Json;

namespace GSpreadSheet
{

    public class GoogleSheets
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "GSpreadSheet Shared Library";

        public Session OpenSession(string docID)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            
            SheetsService service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            Session session = new Session();
            session.Service = service;
            session.DocID = docID;
            return session;
        }

        public void CloseSession(Session session)
        {
            session.Close();
        }

        public void WriteCellValues(Session doc, List<CellAddressWithValue> Values)
        {
            if (doc.IsClosed())
            {
                throw new Exception("The session is closed");
            }

            List<ValueRange> data = new List<ValueRange>();
            foreach (CellAddressWithValue cell in Values)
            {
                IList<object> updateValues = new List<object>();
                updateValues.Add(cell.Value);

                ValueRange vr = new ValueRange { Range = cell.NotationA1(), Values = new List<IList<object>> { updateValues } };
                data.Add(vr);
            }
            
            BatchUpdateValuesRequest requestBody = new BatchUpdateValuesRequest();
            requestBody.Data = data;
            requestBody.ValueInputOption = "RAW";

            SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = doc.Service.Spreadsheets.Values.BatchUpdate(requestBody, doc.DocID);
            
            BatchUpdateValuesResponse response = request.Execute();
            Console.WriteLine(JsonConvert.SerializeObject(response));
        }

        public IList<CellAddressWithValue> ReadCellValues(Session doc, List<CellAddress> Values)
        {
            if (doc.IsClosed())
            {
                throw new Exception("The session is closed");
            }

            List<string> ranges = new List<string>();
            foreach (CellAddress cell in Values)
            {
                ranges.Add(cell.NotationA1());
            }

            SpreadsheetsResource.ValuesResource.BatchGetRequest.ValueRenderOptionEnum valueRenderOption = (SpreadsheetsResource.ValuesResource.BatchGetRequest.ValueRenderOptionEnum)0;  // TODO: Update placeholder value.
            SpreadsheetsResource.ValuesResource.BatchGetRequest.DateTimeRenderOptionEnum dateTimeRenderOption = (SpreadsheetsResource.ValuesResource.BatchGetRequest.DateTimeRenderOptionEnum)0;  // TODO: Update placeholder value
            SpreadsheetsResource.ValuesResource.BatchGetRequest request = doc.Service.Spreadsheets.Values.BatchGet(doc.DocID);
            request.Ranges = ranges;
            request.ValueRenderOption = valueRenderOption;
            request.DateTimeRenderOption = dateTimeRenderOption;

            BatchGetValuesResponse response = request.Execute();
            Console.WriteLine(JsonConvert.SerializeObject(response));
            
            IList<CellAddressWithValue> result = new List<CellAddressWithValue>();
            IList<ValueRange> valueRanges = response.ValueRanges;
            if (valueRanges != null && valueRanges.Count > 0)
            {
                foreach (var range in valueRanges)
                {
                    string sheetName = null, address = null;
                    string[] rangeSplit = range.Range.Split('!');
                    switch (rangeSplit.Length)
                    {
                        case 1:
                            address = rangeSplit[0];
                            break;
                        case 2:
                            sheetName = rangeSplit[0];
                            address = rangeSplit[1];
                            break;
                    }
                    var values = range.Values;
                    foreach (var row in values)
                    {
                        foreach (var col in row)
                        {
                            CellAddressWithValue cav = new CellAddressWithValue
                            {
                                Address = address,
                                SheetName = sheetName,
                                Value = col
                            };
                            result.Add(cav);
                        }
                    }
                }
            }
            return result;
        }
    }
}
