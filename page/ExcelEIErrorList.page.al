namespace cepheo.ExcelImportExport;

using System.IO;
using System.Reflection;
page 88602 ExcelEIErrorList
{
    PageType = List;
    ApplicationArea = All;
    UsageCategory = Administration;
    SourceTable = "Excel Error";
    Caption = 'Excel Import Error List';

    layout
    {
        area(Content)
        {
            repeater(GroupName)
            {
                field(RecId; rec.RecId)
                {
                    toolTip = 'RecId';

                }
                field("Table ID"; rec."Table ID")
                {
                    toolTip = 'Table ID';

                }
                field(PackedCode; rec.PackedCode)
                {
                    toolTip = 'Packed Code';

                }
                field(LineNo; rec.LineNo)
                {
                    toolTip = 'Line No';

                }
                field(Errortext; rec.Errortext)
                {
                    toolTip = 'Error Text';

                }

            }
        }
    }

    actions
    {
        area(Processing)
        {
        }
    }


}
