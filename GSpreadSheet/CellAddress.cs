using System;
using System.Collections.Generic;
using System.Text;

namespace GSpreadSheet
{
    public class CellAddress
    {
        public CellAddress()
        { }

        public CellAddress(string sheetName, string address)
        {
            SheetName = sheetName;
            Address = address;
        }
        public string SheetName;
        public string Address;

        public string NotationA1()
        {
            return this.SheetName != null ? (this.SheetName + "!" + this.Address) : this.Address;
        }
    }

    public class CellAddressWithValue : CellAddress
    {
        public CellAddressWithValue() : base()
        { }

        public CellAddressWithValue(string sheetName, string address, object value) : base(sheetName, address)
        {
            Value = value;
        }
        public object Value;
    }
}
