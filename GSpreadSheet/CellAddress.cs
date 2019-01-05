namespace GSpreadSheet
{
    public class CellAddress
    {
        public CellAddress()
        { }

        public CellAddress(string address)
        {
            Address = address;
        }

        public CellAddress(string sheetName, string address)
        {
            SheetName = sheetName;
            Address = address;
        }
        public string SheetName;
        public string Address;

        public string NotationA1()
        {
            return SheetName != null ? (SheetName + "!" + Address) : Address;
        }
    }

    public class CellAddressWithValue : CellAddress
    {
        public CellAddressWithValue() : base()
        { }

        public CellAddressWithValue(string address, object value) : base(address)
        {
            Value = value;
        }
        public CellAddressWithValue(string sheetName, string address, object value) : base(sheetName, address)
        {
            Value = value;
        }
        public object Value;
    }
}
