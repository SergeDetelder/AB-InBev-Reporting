query 50050 "ABInBevReporting_ABI"
{
    // version DEV11056

    Caption = 'AB InBev Reporting';

    elements
    {
        dataitem(ABInBevBuffer_DMD; ABInBevBuffer_ABI)
        {
            column(ShipTo_ID; ShipTo_ID)
            {
            }
            column(ItemNo; ItemNo)
            {
            }
            column(Sum_Liter; Liter)
            {
                Method = Sum;
            }
            column(Sum_Quantity; Quantity)
            {
                Method = Sum;
            }
        }
    }
}

