namespace cepheo.ExcelImportExport;

table 88602 "Excel Error"
{

    fields
    {
        field(1; RecId; Integer)
        {
            Caption = 'RecID';
            DataClassification = ToBeClassified;
        }
        field(2; "Table ID"; Integer)
        {
            Caption = 'TabellID';
            DataClassification = ToBeClassified;
        }
        field(3; PackedCode; Code[20])
        {
            Caption = 'Pakkekode';
            DataClassification = ToBeClassified;
        }
        field(4; LineNo; Integer)
        {
            Caption = 'Linjenr';
            DataClassification = ToBeClassified;
        }
        field(5; Errortext; Text[100])
        {
            DataClassification = ToBeClassified;
        }
    }

    keys
    {
        key(Key1; RecId)
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
    }
}

