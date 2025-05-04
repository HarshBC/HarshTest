pageextension 70101 SalesQuoteextlist extends "Sales Quotes"
{
    actions
    {
        addafter(MakeInvoice)
        {
            action(ExportToExcel)
            {
                Caption = 'Export to Excel';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Export;
                trigger OnAction()
                var
                begin
                    ExportSalesQuotes;
                end;
            }
        }
    }
    local procedure ExportSalesQuotes()
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        CustLedgerEntriesLbl: Label 'Customer Ledger Entries';
        ExcelFileName: Label 'CustomerLedgerEntries_%1_%2';
        SalesHeader: Record "Sales Header";
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Document Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Sell-to Customer No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Sell-to Customer Name"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("External Document No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Order Date"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Document Date"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Due Date"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption(Status), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Salesperson Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Your Reference"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Sell-to Address"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Sell-to Address 2"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Sell-to City"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Sell-to County"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Sell-to E-Mail"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Sell-to Contact No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption(Amount), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Amount Including VAT"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Shipment Date"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption("Shipment Method Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption(SystemCreatedAt), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption(SystemCreatedBy), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption(SystemModifiedAt), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
        TempExcelBuffer.AddColumn(Salesheader.FieldCaption(SystemModifiedBy), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        SalesHeader.Reset();
        Salesheader.SetRange("Document Type", SalesHeader."Document Type"::Quote);
        if Salesheader.FindSet() then
            repeat
                TempExcelBuffer.NewRow();
                TempExcelBuffer.AddColumn(Salesheader."No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Salesheader."Document Type", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
                TempExcelBuffer.AddColumn(Salesheader."Sell-to Customer No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Salesheader."Sell-to Customer Name", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Salesheader."External Document No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Salesheader."Order Date", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::date);
                TempExcelBuffer.AddColumn(Salesheader."Document Date", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::date);
                TempExcelBuffer.AddColumn(Salesheader."Due Date", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::date);
                TempExcelBuffer.AddColumn(Salesheader."Status", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::date);
                TempExcelBuffer.AddColumn(Salesheader."Salesperson Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(Salesheader."Your Reference", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(Salesheader."Sell-to Address", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(Salesheader."Sell-to Address 2", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(Salesheader."Sell-to City", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(Salesheader."Sell-to County", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(Salesheader."Sell-to E-Mail", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(Salesheader."Sell-to Contact No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(Salesheader."Amount", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(Salesheader."Amount Including VAT", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(Salesheader."Shipment Date", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
                TempExcelBuffer.AddColumn(Salesheader."Shipment Method Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Salesheader.SystemCreatedAt, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
                TempExcelBuffer.AddColumn(Salesheader.SystemCreatedBy, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Salesheader.SystemModifiedAt, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
                TempExcelBuffer.AddColumn(Salesheader.SystemModifiedBy, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
            until Salesheader.Next() = 0;
        TempExcelBuffer.CreateNewBook(CustLedgerEntriesLbl);
        TempExcelBuffer.WriteSheet(CustLedgerEntriesLbl, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileName, CurrentDateTime, UserId));
        TempExcelBuffer.OpenExcel();
        //changes 2
    end;
}