table 50000 "ABInBevBuffer_ABI"
{
    // version DEV11056

    Caption = 'AB InBev Buffer';
    fields
    {
        field(1;"Entry No.";Integer)
        {
        }
        field(10;ExtractionDate;Date)
        {
        }
        field(11;ShipTo_ID;Text[50])
        {
        }
        field(12;ShipTo_Name;Text[100])
        {
        }
        field(13;ShipTo_Address;Text[100])
        {
        }
        field(14;ShipTo_ZIP;Code[20])
        {
        }
        field(15;ShipTo_City;Text[30])
        {
        }
        field(21;BillTo_ID;Code[20])
        {
        }
        field(22;BillTo_Name;Text[100])
        {
        }
        field(23;BillTo_Address;Text[100])
        {
        }
        field(24;BillTo_ZIP;Code[20])
        {
        }
        field(25;BillTo_City;Text[30])
        {
        }
        field(26;BillTo_VAT;Text[50])
        {
        }
        field(30;ItemNo;Code[20])
        {
        }
        field(31;ItemDescription;Text[100])
        {
        }
        field(32;Liter;Decimal)
        {
            DecimalPlaces = 0:5;
        }
        field(33;Quantity;Decimal)
        {
            DecimalPlaces = 0:5;
        }
        field(34;UoM;code[10])
        {
        }
    }

    keys
    {
        key(Key1;"Entry No.")
        {
        }
        key(Key2;ShipTo_ID,ItemNo)
        {
        }
    }

    fieldgroups
    {
    }
}