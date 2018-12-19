using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Text;

namespace GSpreadSheet
{
    public class Session : Object
    {
        private SheetsService service;
        private String docID;

        public string DocID { get => docID; set => docID = value; }
        public SheetsService Service { get => service; set => service = value; }

        public void Close()
        {
            this.service = null;
            this.docID = null;
        }

        public bool IsClosed()
        {
            return this.service == null || this.docID == null;
        }
    }
}
