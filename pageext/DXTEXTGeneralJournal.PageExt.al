namespace cepheo.ExcelImportExport;
using Microsoft.Finance.GeneralLedger.Journal;
using Microsoft.Finance.Payroll;

pageextension 88603 DXTEXTGeneralJournal extends "General Journal"
{


    actions
    {
        addlast(Processing)
        {
            action(ProplanImport)
            {
                Caption = 'Proplan Import';
                ApplicationArea = all;
                tooltip = 'Import Payroll Transactions from Proplan';
                Promoted = true;
                PromotedCategory = Process;
                trigger OnAction()
                var
                    ImportPayrollTransaction: Codeunit "Import Payroll Transaction";
                begin
                    ImportPayrollTransaction.SelectAndImportPayrollDataToGL(Rec,
                    'PROPLANIMPORT');
                end;
            }
        }

    }


}