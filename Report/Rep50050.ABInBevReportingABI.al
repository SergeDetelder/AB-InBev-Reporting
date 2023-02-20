report 50050 "ABInBevReporting_ABI"
{
    // version DEV11056

    Caption = 'AB InBev Reporting';
    ProcessingOnly = true;
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = All;

    dataset
    {
        dataitem(Customer; Customer)
        {
            DataItemTableView = SORTING("No.")
                                WHERE(ABInBevReporting_ABI = CONST(true));
            dataitem("Sales Invoice Header"; "Sales Invoice Header")
            {
                DataItemLink = "Sell-to Customer No." = FIELD("No.");
                DataItemTableView = SORTING("Sell-to Customer No.");
                dataitem("Sales Invoice Line"; "Sales Invoice Line")
                {
                    DataItemLink = "Document No." = FIELD("No.");
                    DataItemTableView = SORTING("Document No.", "Line No.")
                                        WHERE(Type = CONST(Item));

                    trigger OnAfterGetRecord();
                    begin
                        IF NOT Item.GET("No.") THEN
                            CurrReport.SKIP;
                        IF Item.ABInBevReportingUoM_ABI = '' THEN
                            CurrReport.SKIP;
                        IF NOT ItemUoM.GET(Item."No.", Item.ABInBevReportingUoM_ABI) THEN
                            CurrReport.SKIP;
                        ItemUoM.TESTFIELD("Qty. per Unit of Measure");
                        ReportingQuantity := "Quantity (Base)" / ItemUoM."Qty. per Unit of Measure";
                        ReportingLiter := ReportingQuantity * ItemUoM.LtrPerUnit_DMD;
                        AddToABInBevBuffer(0);
                    end;
                }

                trigger OnPreDataItem();
                begin
                    SETRANGE("Posting Date", StartDate, ExtractionDate);
                end;
            }
            dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
            {
                DataItemLink = "Sell-to Customer No." = FIELD("No.");
                DataItemTableView = SORTING("Sell-to Customer No.");
                dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
                {
                    DataItemLink = "Document No." = FIELD("No.");
                    DataItemTableView = SORTING("Document No.", "Line No.")
                                        WHERE(Type = CONST(Item));

                    trigger OnAfterGetRecord();
                    begin
                        IF NOT Item.GET("No.") THEN
                            CurrReport.SKIP;
                        IF Item.ABInBevReportingUoM_ABI = '' THEN
                            CurrReport.SKIP;
                        IF NOT ItemUoM.GET(Item."No.", Item.ABInBevReportingUoM_ABI) THEN
                            CurrReport.SKIP;
                        ItemUoM.TESTFIELD("Qty. per Unit of Measure");
                        ReportingQuantity := "Quantity (Base)" / ItemUoM."Qty. per Unit of Measure" * -1;
                        ReportingLiter := ReportingQuantity * ItemUoM.LtrPerUnit_DMD;
                        AddToABInBevBuffer(1);
                    end;
                }

                trigger OnPreDataItem();
                begin
                    SETRANGE("Posting Date", StartDate, ExtractionDate);
                end;
            }

            trigger OnAfterGetRecord();
            begin
                CurrentCustomerNumber += 1;
                IF GUIALLOWED THEN
                    Progress.UPDATE(1, ROUND(CurrentCustomerNumber / NumberOfCustomers * 10000, 1));
            end;

            trigger OnPostDataItem();
            begin
                IF GUIALLOWED THEN
                    Progress.CLOSE;
            end;

            trigger OnPreDataItem();
            begin
                IF GUIALLOWED THEN BEGIN
                    NumberOfCustomers := COUNT;
                    Progress.OPEN('@1@@@@@@@@@@@@@@@@@@');
                END;
            end;
        }
        dataitem(Integer; Integer)
        {
            DataItemTableView = SORTING(Number);

            trigger OnAfterGetRecord();
            begin
                IF NOT ABInBevQuery.READ THEN
                    CurrReport.BREAK;
                ABInBevBuffer.SETCURRENTKEY(ShipTo_ID, ItemNo);
                ABInBevBuffer.SETRANGE(ShipTo_ID, ABInBevQuery.ShipTo_ID);
                ABInBevBuffer.SETRANGE(ItemNo, ABInBevQuery.ItemNo);
                ABInBevBuffer.FINDFIRST;
                FillExcelBuffer;
            end;

            trigger OnPreDataItem();
            begin
                ABInBevQuery.OPEN;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Extraction Date"; ExtractionDate)
                {
                    ApplicationArea = All;
                    Caption = 'Extraction Date';

                    trigger OnValidate();
                    begin
                        ExtractionDate := CALCDATE('<CM>', ExtractionDate);
                        StartDate := CALCDATE('<-CY>', ExtractionDate);
                    end;
                }
                field("Starting Date"; StartDate)
                {
                    ApplicationArea = All;
                    Caption = 'Starting Date';
                    Editable = false;
                }
            }
        }

        actions
        {
        }
    }

    labels
    {
    }

    trigger OnInitReport();
    begin
        ExtractionDate := CALCDATE('<CM>', TODAY);
        StartDate := CALCDATE('<-CY>', ExtractionDate);
    end;

    trigger OnPostReport();
    begin
        IF Row > 0 THEN BEGIN
            ExcelBuffer.WriteSheet('', COMPANYNAME, USERID);
            ExcelBuffer.CloseBook;
            ExcelBuffer.OpenExcel;
        END;
    end;

    trigger OnPreReport();
    var
        AllCust: Record Customer;
        AllItem: Record Item;
    begin
        //Only For TEST Purpose:  delete afterwards
        /*
        AllCust.MODIFYALL(ABInBevReporting_ABI, TRUE);
        AllItem.SETRANGE(ItemType_DMD, Item.ItemType_DMD::Normal);
        IF AllItem.FIND('-') THEN
          REPEAT
            AllItem.ABInBevReportingUoM_ABI := AllItem."Base Unit of Measure";
            AllItem.MODIFY;
          UNTIL AllItem.NEXT = 0;
        */
        ABInBevBuffer.LOCKTABLE;
        ABInBevBuffer.DELETEALL;
    end;

    var
        ExtractionDate: Date;
        StartDate: Date;
        CompanyInfo: Record "Company Information";
        Item: Record Item;
        ItemUoM: Record "Item Unit of Measure";
        ShiptoAddress: Record "Ship-to Address";
        SellToCustomer: Record Customer;
        BillToCustomer: Record Customer;
        ABInBevBuffer: Record ABInBevBuffer_ABI;
        ExcelBuffer: Record "Excel Buffer" temporary;
        ABInBevQuery: Query ABInBevReporting_ABI;
        ReportingLiter: Decimal;
        ReportingQuantity: Decimal;
        Filename: Text;
        Row: Integer;
        Column: Integer;
        NumberOfCustomers: Integer;
        CurrentCustomerNumber: Integer;
        NextEntryNo: Integer;
        Progress: Dialog;

    local procedure AddToABInBevBuffer(FromDoc: Option Invoice,CrMemo);
    begin
        NextEntryNo += 1;
        ABInBevBuffer."Entry No." := NextEntryNo;
        ABInBevBuffer.ExtractionDate := ExtractionDate;
        CASE FromDoc OF
            FromDoc::Invoice:
                BEGIN
                    ABInBevBuffer.ShipTo_ID := "Sales Invoice Header"."Sell-to Customer No.";
                    IF "Sales Invoice Header"."Ship-to Code" <> '' THEN BEGIN
                        ABInBevBuffer.ShipTo_ID += '_' + "Sales Invoice Header"."Ship-to Code";
                        IF ShiptoAddress.GET("Sales Invoice Header"."Sell-to Customer No.", "Sales Invoice Header"."Ship-to Code") THEN BEGIN
                            ABInBevBuffer.ShipTo_Name := ShiptoAddress.Name;
                            ABInBevBuffer.ShipTo_Address := ShiptoAddress.Address;
                            ABInBevBuffer.ShipTo_ZIP := PostCode2ZIPCode(ShiptoAddress."Post Code");
                            ABInBevBuffer.ShipTo_City := ShiptoAddress.City;
                        END;
                    END ELSE BEGIN
                        IF SellToCustomer.GET("Sales Invoice Header"."Sell-to Customer No.") THEN BEGIN
                            ABInBevBuffer.ShipTo_Name := SellToCustomer.Name;
                            ABInBevBuffer.ShipTo_Address := SellToCustomer.Address;
                            ABInBevBuffer.ShipTo_ZIP := PostCode2ZIPCode(SellToCustomer."Post Code");
                            ABInBevBuffer.ShipTo_City := SellToCustomer.City;
                        END;
                    END;
                    ABInBevBuffer.BillTo_ID := "Sales Invoice Header"."Bill-to Customer No.";
                    IF BillToCustomer.GET("Sales Invoice Header"."Bill-to Customer No.") THEN BEGIN
                        ABInBevBuffer.BillTo_Name := BillToCustomer.Name;
                        ABInBevBuffer.BillTo_Address := BillToCustomer.Address;
                        ABInBevBuffer.BillTo_ZIP := PostCode2ZIPCode(BillToCustomer."Post Code");
                        ABInBevBuffer.BillTo_City := BillToCustomer.City;
                    END;
                    ABInBevBuffer.BillTo_VAT := GetCustomerVATNo("Sales Invoice Header"."Bill-to Customer No.");
                END;
            FromDoc::CrMemo:
                BEGIN
                    ABInBevBuffer.ShipTo_ID := "Sales Cr.Memo Header"."Sell-to Customer No.";
                    IF "Sales Cr.Memo Header"."Ship-to Code" <> '' THEN BEGIN
                        ABInBevBuffer.ShipTo_ID += '_' + "Sales Cr.Memo Header"."Ship-to Code";
                        IF ShiptoAddress.GET("Sales Cr.Memo Header"."Sell-to Customer No.", "Sales Cr.Memo Header"."Ship-to Code") THEN BEGIN
                            ABInBevBuffer.ShipTo_Name := ShiptoAddress.Name;
                            ABInBevBuffer.ShipTo_Address := ShiptoAddress.Address;
                            ABInBevBuffer.ShipTo_ZIP := PostCode2ZIPCode(ShiptoAddress."Post Code");
                            ABInBevBuffer.ShipTo_City := ShiptoAddress.City;
                        END;
                    END ELSE BEGIN
                        IF SellToCustomer.GET("Sales Cr.Memo Header"."Sell-to Customer No.") THEN BEGIN
                            ABInBevBuffer.ShipTo_Name := SellToCustomer.Name;
                            ABInBevBuffer.ShipTo_Address := SellToCustomer.Address;
                            ABInBevBuffer.ShipTo_ZIP := PostCode2ZIPCode(SellToCustomer."Post Code");
                            ABInBevBuffer.ShipTo_City := SellToCustomer.City;
                        END;
                    END;
                    ABInBevBuffer.BillTo_ID := "Sales Cr.Memo Header"."Bill-to Customer No.";
                    IF BillToCustomer.GET("Sales Invoice Header"."Bill-to Customer No.") THEN BEGIN
                        ABInBevBuffer.BillTo_Name := BillToCustomer.Name;
                        ABInBevBuffer.BillTo_Address := BillToCustomer.Address;
                        ABInBevBuffer.BillTo_ZIP := PostCode2ZIPCode(BillToCustomer."Post Code");
                        ABInBevBuffer.BillTo_City := BillToCustomer.City;
                    END;
                    ABInBevBuffer.BillTo_VAT := GetCustomerVATNo("Sales Cr.Memo Header"."Bill-to Customer No.");
                END;
        END;
        ABInBevBuffer.ItemNo := Item."No.";
        ABInBevBuffer.ItemDescription := Item.Description;
        ABInBevBuffer.Liter := ReportingLiter;
        ABInBevBuffer.Quantity := ReportingQuantity;
        ABInBevBuffer.UoM := Item.ABInBevReportingUoM_ABI;
        ABInBevBuffer.INSERT;
    end;

    local procedure PostCode2ZIPCode(PostCode: Code[20]): Text;
    begin
        IF COPYSTR(PostCode, 1, 3) = 'BE-' THEN
            EXIT(COPYSTR(PostCode, 4))
        ELSE
            EXIT(PostCode);
    end;

    local procedure FillExcelBuffer();
    begin
        IF Row = 0 THEN
            FillTitleRow;
        Row += 1;
        Column := 1;
        EnterCell(Row, Column, FORMAT(ExtractionDate, 0, '<Day,2>/<Month,2>/<Year4>'), FALSE, FALSE, FALSE, 2);
        EnterCell(Row, Column, ABInBevQuery.ShipTo_ID, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.ShipTo_Name, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.ShipTo_Address, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.ShipTo_City, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.ShipTo_ZIP, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.BillTo_ID, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.BillTo_Name, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.BillTo_Address, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.BillTo_City, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.BillTo_ZIP, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.BillTo_VAT, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevQuery.ItemNo, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, ABInBevBuffer.ItemDescription, FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, FORMAT(ABInBevQuery.Sum_Liter, 0, '<Precision,2:2><Standard Format,1>'), FALSE, FALSE, FALSE, 0);
        EnterCell(Row, Column, 'L', FALSE, FALSE, FALSE, 1);
        EnterCell(Row, Column, FORMAT(ABInBevQuery.Sum_Quantity, 0, '<Precision,2:2><Standard Format,1>'), FALSE, FALSE, FALSE, 0);
        EnterCell(Row, Column, ABInBevBuffer.UoM, FALSE, FALSE, FALSE, 1);
    end;

    local procedure FillTitleRow();
    begin
        CompanyInfo.GET;
        Filename := TEMPORARYPATH +
          'ABInBev_' +
          FORMAT(ExtractionDate, 0, '<Year4>_<Month,2>_<Day,2>') + '_' +
          DELCHR(CompanyInfo.Name, '=', ' /\*?%') + '_' +
          DELCHR(CompanyInfo.City, '=', ' /\*?%') + '_' +
          '.xlsx';
        ExcelBuffer.DELETEALL;
        CLEAR(ExcelBuffer);
        ExcelBuffer.CreateBook(Filename, 'AB InBev');
        Row += 1;
        Column := 1;
        EnterCell(Row, Column, 'ExtractionDate', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'POC_ShipTo_ID', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'POC_ShipTo_Name', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'POC_ShipTo_Address', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'POC_ShipTo_City', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'POC_ShipTo_ZIPCode', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'POC_BillTo_ID', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'POC_BillTo_Name', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'POC_BillTo_Address', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'POC_BillTo_City', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'POC_BillTo_ZIPCode', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'VAT_BillTo', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'SKU_Number', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'SKU_Name', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'Volume', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'Volume_UOM', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'Amount', TRUE, FALSE, FALSE, 1);
        EnterCell(Row, Column, 'Amount_UOM', TRUE, FALSE, FALSE, 1);
    end;

    local procedure EnterCell(RowNo: Integer; var ColumnNo: Integer; CellValue: Text[250]; Bold: Boolean; Italic: Boolean; UnderLine: Boolean; CellType: Option Number,Text,Date,Time);
    begin
        ExcelBuffer.INIT;
        ExcelBuffer.VALIDATE("Row No.", RowNo);
        ExcelBuffer.VALIDATE("Column No.", ColumnNo);
        ExcelBuffer."Cell Value as Text" := CellValue;
        ExcelBuffer.Formula := '';
        ExcelBuffer.Bold := Bold;
        ExcelBuffer.Italic := Italic;
        ExcelBuffer.Underline := UnderLine;
        ExcelBuffer."Cell Type" := CellType;
        ExcelBuffer.INSERT;

        ColumnNo += 1;
    end;

    local procedure GetCustomerVATNo(CustomerNo: Code[20]) VATNo: Text;
    var
        BillToCustomer: Record Customer;
    begin
        IF BillToCustomer.GET(CustomerNo) THEN
            IF BillToCustomer."Enterprise No." <> '' THEN
                VATNo := BillToCustomer."Enterprise No."
            ELSE
                VATNo := BillToCustomer."VAT Registration No.";
    end;
}