tableextension 50050 "Customer_ABI" extends Customer  //18
{
    fields
    {
        field(50000; ABInBevReporting_ABI; Boolean)
        {
            Caption = 'AB InBev Reporting';
            DataClassification = CustomerContent;
            Description = 'DEV11056';
        }
    }
}