namespace cepheo.ExcelImportExport;
using Microsoft.Inventory.Item;
using Microsoft.Purchases.Vendor;
tableextension 88601 DXTEXItem extends "Item"
{
    fields
    {
        field(88600; "Vendor Exists"; Boolean)
        {
            caption = 'Vendor Exists';
            FieldClass = FlowField;
            calcformula = exist(Vendor where("No." = field("Vendor No.")));
            editable = false;
            // CalcFormula = Exist("Vendor" where("No." = FIELD("Vendor No.")));
            //Enabled = false;
            // fieldclass = flowfield;
            // calcformula = Exists("Vendor");

        }
    }

}