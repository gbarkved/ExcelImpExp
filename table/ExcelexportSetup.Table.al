namespace cepheo.ExcelImportExport;
using Microsoft.Finance.GeneralLedger.Journal;
using System.IO;
using System.Reflection;


table 88600 "Excel export Setup"
{

    fields
    {
        field(1; "Code"; Code[10])
        {
            Caption = 'Code';
            DataClassification = CustomerContent;
        }
        field(2; "Utgår"; Option)
        {
            Caption = 'Format';
            DataClassification = CustomerContent;
            OptionCaption = 'BBS,Data Dialog';
            OptionMembers = BBS,"Data Dialog";
        }
        field(10; FileName; Text[250])
        {
            Caption = 'FileName';
            DataClassification = CustomerContent;
        }
        field(11; SheetName; Text[50])
        {
            Caption = 'Arknavn';
            DataClassification = CustomerContent;
            Editable = false;
        }
        field(12; TableID; Integer)
        {
            Caption = 'Tabell ID';
            DataClassification = CustomerContent;
            TableRelation = AllObjWithCaption."Object ID" where("Object Type" = const(TableData));
            // TableRelation = Object.ID WHERE (Type = CONST (Table));
        }
        field(13; ExcelMalCode; Code[10])
        {
            Caption = 'Excelmalkode';
            DataClassification = CustomerContent;
        }
        field(14; ExcelMalName; Text[30])
        {
            Caption = 'Excelmalnavn';
            DataClassification = CustomerContent;
        }
        field(15; FieldIDs; Text[150])
        {
            Caption = 'Felt ID';
            DataClassification = CustomerContent;
        }
        field(16; ValidateFieldIDs; Text[150])
        {
            Caption = 'Valider felt';
            DataClassification = CustomerContent;
        }
        field(17; "Opprett nytt regneark"; Boolean)
        {
            Caption = 'Opprett nytt regneark';
            DataClassification = CustomerContent;
        }
        field(18; TableName; Text[250])
        {
            Caption = 'Table Name';

            FieldClass = FlowField;
            CalcFormula = Lookup(AllObjWithCaption."Object Caption" where("Object Type" = const(TableData), "Object ID" = field(TableID)));
            Editable = false;
        }

        field(19; KeyValue; Text[60])
        {
            DataClassification = CustomerContent;
        }
        field(20; ExcelBufRow; Integer)
        {
            DataClassification = CustomerContent;
        }
        field(21; ImportLogFil; Text[200])
        {
            DataClassification = CustomerContent;
        }
        field(22; "Journal Template Name"; Text[30])
        {
            Caption = 'Journal Template Name';
            DataClassification = CustomerContent;
            TableRelation = "Gen. Journal Template";
        }
        field(23; "Journal Batch Name"; Code[10])
        {
            Caption = 'Journal Batch Name';
            DataClassification = CustomerContent;
            TableRelation = "Gen. Journal Batch".Name WHERE("Journal Template Name" = FIELD("Journal Template Name"));
        }
        field(24; IBLineNo; Integer)
        {
            Caption = 'IB Line No';
            DataClassification = CustomerContent;
        }
        field(25; IBDate; Date)
        {
            Caption = 'IB Date';
            DataClassification = CustomerContent;
        }
    }

    keys
    {
        key(Key1; "Code")
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
    }
    trigger OnDelete()
    var
        ConfigPackage: Record "Config. Package";
        ConfigFields: Record "Config. Package Field";
        ConfigPackageLines: record "Config. Package Table";
    begin
        ConfigFields.SetRange("Package Code", rec.Code);
        if ConfigFields.FindSet() then
            repeat
                ConfigFields.Delete();
            until ConfigFields.Next() = 0;
        ConfigPackageLines.SetRange("Package Code", rec.Code);
        if ConfigPackageLines.FindSet() then
            repeat
                ConfigPackageLines.Delete();
            until ConfigPackageLines.Next() = 0;

        ConfigPackage.SetRange(Code, rec.Code);
        if ConfigPackage.FindSet() then
            repeat
                ConfigPackage.Delete();
            until ConfigPackage.Next() = 0;
    end;

}

