tableextension 50051 "Item_ABI" extends Item    //27
{
    fields
    {
        field(50000; ABInBevReportingUoM_ABI; Code[10])
        {
            Caption = 'AB InBev Reporting UoM';
            DataClassification = CustomerContent;
            Description = 'DEV11056';
            TableRelation = "Item Unit of Measure".Code WHERE ("Item No."=FIELD("No."));
        }
    }

}