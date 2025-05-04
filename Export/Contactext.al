pageextension 70102 Contactlistext extends "Contact List"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addafter("Related Information")
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
                    ExportContact;
                end;
            }
        }
    }

    procedure ExportContact()
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        CustLedgerEntriesLbl: Label 'Conatct';
        ExcelFileName: Label 'Contacts_%1_%2';
        Contact: Record Contact;
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn(Contact.FieldCaption("No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Contact.FieldCaption(Name), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Contact.FieldCaption(Type), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Contact.FieldCaption("Company No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Contact.FieldCaption("Company Name"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Contact.FieldCaption("Contact Business Relation"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
        TempExcelBuffer.AddColumn(Contact.FieldCaption(sky_synctobc), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
        TempExcelBuffer.AddColumn(Contact.FieldCaption(SkySyncing), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Contact.FieldCaption(sky_Sync_to_CRM), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        Contact.Reset();
        if Contact.FindSet() then
            repeat
                TempExcelBuffer.NewRow();
                TempExcelBuffer.AddColumn(Contact."No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Contact.Name, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Contact.Type, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Contact."Company No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Contact."Company Name", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Contact."Contact Business Relation", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(Contact.sky_synctobc, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(Contact.SkySyncing, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(Contact.sky_Sync_to_CRM, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
            until Contact.Next() = 0;
        TempExcelBuffer.CreateNewBook(CustLedgerEntriesLbl);
        TempExcelBuffer.WriteSheet(CustLedgerEntriesLbl, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileName, CurrentDateTime, UserId));
        TempExcelBuffer.OpenExcel();
        //changes 
    end;

    var
        myInt: Integer;
}