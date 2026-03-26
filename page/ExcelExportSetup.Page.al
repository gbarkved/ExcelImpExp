namespace cepheo.ExcelImportExport;

using System.IO;
using System.Reflection;

page 88600 "Excel Export Setup"
{
    PageType = Card;
    SourceTable = "Excel export Setup";
    caption = 'Excel Export Setup';
    ApplicationArea = All;
    UsageCategory = Administration;

    layout
    {
        area(content)
        {
            group(General)
            {
                field(Code; rec.Code)
                {
                    ApplicationArea = All;
                    Caption = 'Primary Key';
                    ToolTip = 'Specifies the primary key.';
                    Editable = true;


                }

                field(FileName; rec.FileName)
                {


                    ToolTip = 'Specifies the name of the Excel file to which data is to be exported.';
                    trigger OnAssistEdit()
                    var
                        lFileName: Text[250];
                        lSheetName: Text[50];
                    begin
                        //TempBlob.Blob.CREATEINSTREAM(lInStream);
                        //IF UPLOADINTOSTREAM('Get excel','','Excel Files|*.xlsx|CSV Files|*.csv', FileName, lInStream) THEN
                        ExcelExp.GetExcelDocument('Hent Excel dokument', lFileName, lSheetName);
                        rec.FileName := CopyStr(lFileName, 1, MaxStrLen(rec.FileName));
                        rec.SheetName := CopyStr(lSheetName, 1, MaxStrLen(rec.SheetName));
                        rec.VALIDATE(rec.FileName);
                    end;
                }
                field(SheetName; rec.SheetName)
                {
                    ToolTip = 'Specifies the name of the sheet in the Excel file.';
                }
                field(TableID; rec.TableID)
                {
                    ToolTip = 'Specifies the table to which the data is to be exported.';

                }
                field(ExcelMalCode; rec.ExcelMalCode)
                {
                    ToolTip = 'Specifies a code for the Excel template to use when exporting data.';
                }
                field(ExcelMalName; rec.ExcelMalName)
                {
                    ToolTip = 'Specifies the name of the Excel template to use when exporting data.';
                }
                field(FieldIDs; rec.FieldIDs)
                {
                    ToolTip = 'Specifies the fields to export to Excel.';
                    Visible = false;

                }
                field(ValidateFieldIDs; rec.ValidateFieldIDs)
                {
                    ToolTip = 'Specifies the fields to validate before exporting to Excel.';
                    Visible = false;
                }
                field("Opprett nytt regneark"; rec."Opprett nytt regneark")
                {
                    ToolTip = 'Specifies whether to create a new worksheet in the Excel file for the export.';
                }
                field(TableName; rec.TableName)
                {
                    ToolTip = 'Specifies the name of the table to which the data is to be exported.';
                }
                field(KeyValue; rec.KeyValue)
                {
                    ToolTip = 'Specifies the key value to use to find records in the table.';
                    Visible = false;
                }
                field(ExcelBufRow; rec.ExcelBufRow)
                {
                    ToolTip = 'Specifies the row in the Excel buffer to use for the export.';
                    Visible = false;
                }
                field(ImportLogFil; rec.ImportLogFil)
                {
                    ToolTip = 'Specifies the name of the log file to create during the import.';
                }
                field("Journal Template Name"; Rec."Journal Template Name")
                {
                    ToolTip = 'Specifies the name of the journal template to use when importing data.';
                    //Visible = false;
                }
                Field("Journal Batch Name"; Rec."Journal Batch Name")
                {
                    ToolTip = 'Specifies the name of the journal batch to use when importing data.';
                }
                field(IBDate; Rec.IBDate)
                {
                    ToolTip = 'Specifies the date to use for the opening balance when importing data.';
                }
            }
        }
    }

    actions
    {
        area(creation)
        {
            action("Hent Excel")
            {
                Image = Excel;
                ApplicationArea = All;
                Promoted = true;
                Caption = 'Import Excel document';
                PromotedCategory = Process;
                PromotedIsBig = true;
                ToolTip = 'Imports data from an Excel file.';
                trigger OnAction()
                begin
                    ExcelExp.GetExcelDocument('', Rec.FileName, rec.SheetName);
                    ExcelExp.RUN(Rec);
                end;
            }
            action("Last ned Excel")
            {
                Image = Excel;
                ApplicationArea = All;
                Promoted = true;
                Caption = 'Export Excel document';
                PromotedCategory = Process;
                PromotedIsBig = true;
                ToolTip = 'Exports data to an Excel file.';
                trigger OnAction()
                begin
                    ExcelExp.SaveToExcel(Rec);

                end;
            }
            action(SelectFields)
            {
                Image = Design;
                ApplicationArea = All;
                Promoted = true;
                Caption = 'Select fields';
                PromotedCategory = Process;
                PromotedIsBig = true;
                ToolTip = 'Select fields to be included in the export.';

                trigger OnAction()
                var
                    ConfigRec: record "Config. Package";
                    ConfigFields: record "Config. Package Field";
                    ConfigPackage: record "Config. Package Table";
                    ConfigFieldPage: page "Config. Package Fields";
                //test: page 88603;// DXTEXTest;
                //DXTEXCConfigPageFields: page dxt
                begin
                    if not ConfigRec.get(rec.Code) then begin
                        ConfigRec.init();
                        ConfigRec.Code := rec.Code;
                        ConfigRec."Package Name" := 'Excel Export ' + rec.Code;
                        ConfigRec."Language ID" := 1044;
                        ConfigRec."Exclude Config. Tables" := true;
                        ConfigRec.Insert();
                        ConfigPackage.init();
                        ConfigPackage."Package Code" := ConfigRec.Code;
                        //ConfigPacka
                        ConfigPackage."Table ID" := rec.TableID;
                        ConfigPackage.Insert();
                        SetupFields(rec.TableID, ConfigFields);
                        commit();
                    end;
                    ConfigFields.SETRANGE("Package Code", rec.Code);
                    ConfigFieldPage.SettExcelExport(true);
                    ConfigFieldPage.SetTableView(ConfigFields);
                    ConfigFieldPage.RunModal();
                end;
            }
            action(SetFilter)
            {
                Image = Filter;
                ApplicationArea = All;
                Promoted = true;
                Caption = 'Set filters';
                PromotedCategory = Process;
                PromotedIsBig = true;
                ToolTip = 'Set filters for the data to be exported.';
                trigger OnAction()
                var
                    ConfigRec: record "Config. Package";
                    ConfigFilters: record "Config. Package Filter";
                    ConfigPackage: record "Config. Package Table";

                begin
                    if configrec.get(rec.Code) then begin
                        ConfigFilters.SetRange("Package Code", rec.Code);
                        ConfigFilters.SetRange("Table ID", rec.TableID);
                        ConfigFilters.SetRange("Processing Rule No.", ConfigPackage."Processing Report ID");
                        Page.RunModal(8623, ConfigFilters);
                    end;

                end;
            }
            action(GetSheetNames)
            {
                Image = Refresh;
                ApplicationArea = All;
                Promoted = true;
                Caption = 'Get sheet names';
                PromotedCategory = Process;
                PromotedIsBig = true;
                ToolTip = 'Get the names of the sheets in the selected Excel file.';
                trigger OnAction()
                begin
                    ExcelExp.GetSheetsName(Rec.FileName);

                end;
            }
            action(GetBlobField)
            {
                Image = Excel;
                ApplicationArea = All;
                Promoted = true;
                Caption = 'Get Blob field';
                PromotedCategory = Process;
                PromotedIsBig = true;
                ToolTip = 'Get the name of the Blob field in the selected table.';
                trigger OnAction()
                var

                    ReadXML: codeunit DXTEXLXMLImport;
                    InStr: InStream;



                begin

                    // Nå kan du bruke ContentText som inneholder blob-innholdet som tekst
                    ReadXML.ImportXMLToBuffer(InStr);

                end;
            }
        }
    }

    var

        ExcelExp: Codeunit "Les fra Excel";

    local procedure SetupFields(lTableID: Integer; var lConfigField: Record "Config. Package Field")
    var
        AllFields: record "Field";
        ProcessingOrder: Integer;
    begin
        AllFields.Reset();
        AllFields.SetRange(TableNo, lTableID);
        AllFields.SetRange(Class, AllFields.Class::Normal, AllFields.Class::FlowField);
        AllFields.SetRange("No.", 1, 2000000000 - 1); // Exclude system fields
        ProcessingOrder := 1;
        if AllFields.FindSet() then
            repeat
                lConfigField.Init();
                lConfigField."Package Code" := rec.Code;
                lConfigField."Table ID" := lTableID;
                lConfigField."Field ID" := AllFields."No.";
                lConfigField."Processing Order" := ProcessingOrder;
                ProcessingOrder += 1;

                lConfigField."Field Name" := AllFields.FieldName;
                lConfigField."Field Caption" := AllFields."Field Caption";
                lConfigField."Package Code" := rec.Code;
                lConfigField."Primary Key" := AllFields.IsPartOfPrimaryKey;
                lConfigField."Include Field" := lConfigField."Primary Key";
                lConfigField."Relation Table ID" := AllFields.RelationTableNo;
                lConfigField.Insert();
            until AllFields.Next() = 0;
    end;
}

