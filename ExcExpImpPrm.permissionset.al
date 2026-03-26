namespace cepheo.ExcelImportExport;

using System.IO;

permissionset 88600 ExcExpImpPrm
{
    Assignable = true;
    Permissions = tabledata "Excel Error" = RIMD,
        tabledata "Excel export Setup" = RIMD,
        tabledata "Config. Package Field" = RIMD,
        tabledata "Config. Package Filter" = RIMD,
        tabledata "Config. Package Table" = RIMD,
        tabledata "Config. Package" = RIMD,
        tabledata "Excel Buffer" = RIMD,
        table "Excel Error" = X,
        table "Excel export Setup" = X,
        codeunit CheckAndUpdateaRecord = X,
        codeunit "Les fra Excel" = X,
        page "Excel Export Setup" = X;
}