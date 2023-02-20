pageextension 50051 "ItemCard_ABI" extends "Item Card" //30
{
    layout
    {
        addlast(Reporting_DMD)
        {
            field(ABInBevReportingUoM_ABI; ABInBevReportingUoM_ABI)
            {
                ApplicationArea = All;
                Importance = Promoted;
                Description = 'DEV11056';
            }
        }
    }
}