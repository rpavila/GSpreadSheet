using System;
using System.Collections.Generic;
using System.Text;

namespace GSpreadSheet
{
    public class CellAddress : Object
    {
        public String SheetName;
        public String Address;

        public string NotationA1()
        {
            return this.SheetName != null ? (this.SheetName + "!" + this.Address) : this.Address;
        }
    }

    public class CellAddressWithValue : CellAddress
    {
        public Object Value;
    }
}
