namespace cepheo.ExcelImportExport;
using System.IO;
pageextension 88602 DXTEXCConfigPageFields extends "Config. Package Fields"
{
    layout
    {
        // Add changes to page layout here
        addafter("Relation Table Caption")
        {
            field(Treatment; Rec.Treatment)
            {
                ApplicationArea = All;
                Caption = 'Treatment';
                ToolTip = 'Specifies how the field is treated during processing. 0 = No action, >0 = Actions to be taken';
            }
        }
    }

    actions
    {
        // Add changes to page actions here
        addlast(Processing)
        {


            action(SetTemplate)
            {
                Caption = 'Set Template';
                ApplicationArea = All;
                PromotedCategory = Process;
                Visible = ExcelExport;
                Promoted = true;
                Image = Add;
                ToolTip = 'Insert template record for Excel Export.';

                trigger OnAction()
                begin
                    InsertTemplate();
                end;
            }
        }

    }
    var
        ExcelExport: boolean;

    local procedure InsertTemplate()
    var
        MalL: Record "Config. Package Field";
    begin
        if not mall.get(rec."Package Code", rec."Table ID", 99601) then begin
            Mall."Package Code" := rec."Package Code";
            Mall."Table ID" := rec."Table ID";
            Mall."Field ID" := 99601;
            Mall.Insert();
        end;
        case mall."Table ID" of
            18:
                mall."Field Name" := 'Kundemal';
            23:
                mall."Field Name" := 'Leverandørmal';
            27:
                mall."Field Name" := 'Varemal';
            else
                MalL."Field Name" := 'Mal';
        end;

        mall."Include Field" := true;
        mall."Validate Field" := true;
        mall."Processing Order" := 99601;
        mall."Field Caption" := mall."Field Name";
        Mall.Modify();

    end;

    procedure SettExcelExport(Value: boolean)
    begin
        ExcelExport := Value;
    end;





    trigger OnClosePage()
    begin
        ExcelExport := false;
    end;
}
