using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace GSpreadSheet
{
    public class GoogleSheets
    {
        private string credentialsPath;
        private string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private string ApplicationName = "GSpreadSheet Shared Library";

        public GoogleSheets(string credentialsPath)
        {
            this.credentialsPath = credentialsPath;
        }

        public Session OpenSession(string docID)
        {
            UserCredential credential;

            //using (var stream =
            //    new FileStream(this.credentialsPath, FileMode.Open, FileAccess.Read))
            //{
            //    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(new GoogleAuthorizationCodeFlow.Initializer
            //    {
            //        ClientSecrets = GoogleClientSecrets.Load(stream).Secrets
            //    },
            //    Scopes,
            //    "user",
            //    CancellationToken.None,
            //    new FileDataStore("GoogleSheets.Credentials")).Result;
            //}
            using (var stream =
                new FileStream(this.credentialsPath, FileMode.Open, FileAccess.Read))
            {
                var credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }
            
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            var session = new Session();
            session.Service = service;
            session.DocID = docID;
            return session;
        }

        public void CloseSession(Session session)
        {
            session.Close();
        }

        public ExecutionResult WriteCellValues(Session doc, List<CellAddressWithValue> Values)
        {
            var result = new ExecutionResult();
            if (doc.IsClosed())
            {
                result.Result = ResultTypes.Error;
                result.Messages = new[] { "The session is closed" };
                return result;
            }

            var data = new List<ValueRange>();
            var errors = "";
            foreach (var cell in Values)
            {
                IList<object> updateValues = new List<object>();
                updateValues.Add(cell.Value);

                if (cell.Address.Contains(":"))
                {
                    if (errors.Length > 0)
                    {
                        errors += ";";
                    }
                    errors += "The cell address [" + cell.Address + "] has a invalid format";
                }
                else
                {
                    var vr = new ValueRange { Range = cell.NotationA1(), Values = new List<IList<object>> { updateValues } };
                    data.Add(vr);
                }
            }
            if (errors.Length > 0)
            {
                result.Result = ResultTypes.Error;
                result.Messages = errors.Split(';');
                return result;
            }

            var requestBody = new BatchUpdateValuesRequest();
            requestBody.Data = data;
            requestBody.ValueInputOption = "RAW";

            var request = doc.Service.Spreadsheets.Values.BatchUpdate(requestBody, doc.DocID);

            var response = request.Execute();

            result.Result = ResultTypes.Success;
            result.Messages = new[] { "Operation successfully!!!" };
            return result;
        }

        public ExecutionResultWithData<IList<CellAddressWithValue>> ReadCellValues(Session doc, List<CellAddress> Values)
        {
            var result = new ExecutionResultWithData<IList<CellAddressWithValue>>();
            if (doc.IsClosed())
            {
                result.Result = ResultTypes.Error;
                result.Messages = new[] { "The session is closed" };
                return result;
            }

            var ranges = new List<string>();
            var errors = "";
            foreach (var cell in Values)
            {
                if (cell.Address.Contains(":"))
                {
                    if (errors.Length > 0)
                    {
                        errors += ";";
                    }
                    errors += "The cell address [" + cell.Address + "] has a invalid format";
                }
                else
                {
                    ranges.Add(cell.NotationA1());
                }
            }
            if (errors.Length > 0)
            {
                result.Result = ResultTypes.Error;
                result.Messages = errors.Split(';');
                return result;
            }

            var valueRenderOption = (SpreadsheetsResource.ValuesResource.BatchGetRequest.ValueRenderOptionEnum)0;  // TODO: Update placeholder value.
            var dateTimeRenderOption = (SpreadsheetsResource.ValuesResource.BatchGetRequest.DateTimeRenderOptionEnum)0;  // TODO: Update placeholder value
            var request = doc.Service.Spreadsheets.Values.BatchGet(doc.DocID);
            request.Ranges = ranges;
            request.ValueRenderOption = valueRenderOption;
            request.DateTimeRenderOption = dateTimeRenderOption;

            var response = request.Execute();

            var data = new List<CellAddressWithValue>();
            var valueRanges = response.ValueRanges;
            if (valueRanges != null && valueRanges.Count > 0)
            {
                foreach (var range in valueRanges)
                {
                    string sheetName = null, address = null;
                    var rangeSplit = range.Range.Split('!');
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
                            var cav = new CellAddressWithValue
                            {
                                Address = address,
                                SheetName = sheetName,
                                Value = col
                            };
                            data.Add(cav);
                        }
                    }
                }
            }
            result.Result = ResultTypes.Success;
            result.Data = data;
            return result;
        }
    }
}
