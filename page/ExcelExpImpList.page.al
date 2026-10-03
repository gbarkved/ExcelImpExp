namespace cepheo.ExcelImportExport;

using System.IO;
using System.Reflection;
page 88601 ExcelExpImpList
{
    PageType = List;
    ApplicationArea = All;
    Caption = 'Excel Export Setup List';
    UsageCategory = Lists;
    SourceTable = "Excel export Setup";
    CardPageId = 88600;


    layout
    {

        area(Content)
        {
            repeater(GroupName)
            {
                field(Code; rec.Code)
                {
                    ToolTip = 'Code';
                }
                field(FileName; rec.FileName)
                {
                    ToolTip = 'File Name';
                }

            }
        }
        area(Factboxes)
        {

        }
    }

    actions
    {
        area(Processing)
        {

        }

    }
}