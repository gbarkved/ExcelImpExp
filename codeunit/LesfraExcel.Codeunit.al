namespace cepheo.ExcelImportExport;
using Microsoft.Foundation.UOM;
using Microsoft.inventory.Item;
using Microsoft.Inventory.Item.Catalog;
using Microsoft.Purchases.Vendor;
using Microsoft.Utilities;

using System.IO;

codeunit 88600 "Les fra Excel"
{
    // ///*SAK-17699-JWW9 Navicom GeirB 10.08.2012: Tekst legges ut med ' til Excel
    // ///*SAK-14635-FTC2 Navicom GeirB 27.03.2012: Utvidet rutinene til å søke opp 10 nøkler.
    // //GeirB, 02.juli.12, Utvidet S fra 60 til 250 tegn

    TableNo = "Excel Export Setup";


    trigger OnRun()
    begin

        setCode(Rec);
    end;

    var
        TempExcelBuff: Record 370 temporary;

        MalL: Record "Config. Package Field";

        MalFilter: Record 8626;
        CheckAndUpdate: Codeunit CheckAndUpdateaRecord;
        AnyTable: RecordRef;

        rad: Integer;
        kol: Integer;
        Felt: array[50] of Integer;

        lExcelStream: InStream;
        gFileName: Text;
        Text0001Msg: label 'Table ID in Excel sheet does not match selected table ID.';


    procedure setCode(arec: Record "Excel Export Setup")
    begin
        if arec.FileName = '' then
            EXIT;
        CLEAR(TempExcelBuff);
        TempExcelBuff.DELETEALL();
        if arec.SheetName = '' then
            arec.SheetName := 'Ark1';
        TempExcelBuff.OpenBookStream(lExcelStream, arec.SheetName);
        //TempExcelBuff.OpenBook(arec.FileName,arec.SheetName);
        TempExcelBuff.ReadSheet();

        Copy2Table(arec);

    end;


    procedure GetSheetsName(lFileName: Text) Result: Text
    var


        Option: integer;
    begin
        Result := TempExcelBuff.SelectSheetsNameStream(lExcelStream);
        Option := StrMenu(Result, 1);
        if Option <> 0 then
            Result := SelectStr(Option, Result);

    end;





    procedure SaveToExcel(ExcelSetup: Record "Excel Export Setup")
    var
        tmpRec: RecordRef;
        tmpfr: FieldRef;
        i: Integer;


        Window: Dialog;
    begin
        Window.OPEN('Overføring fra\' +
        '#1###########################################\' +
        'Rad: #2#########');
        ExcelSetup.TESTFIELD(ExcelMalName);
        //ExcelSetup.CALCFIELDS(TableName);
        Window.UPDATE(1, ExcelSetup.TableName);
        if NOT ExcelSetup."Opprett nytt regneark" then begin
            ExcelSetup.TESTFIELD(FileName);
            ExcelSetup.TESTFIELD(SheetName);
        end;
        TempExcelBuff.DELETEALL();
        tmpRec.OPEN(ExcelSetup.TableID);

        if ExcelSetup.SheetName = '' then
            ExcelSetup.SheetName := CopyStr(ExcelSetup.TableName, 1, MaxStrLen(ExcelSetup.SheetName));
        TempExcelBuff.CreateNewBook(ExcelSetup.SheetName);
        MalL.Reset();
        MalL.SetCurrentKey("Processing Order");
        MalL.SETRANGE("Table ID", ExcelSetup.TableID);
        MalL.SETRANGE("Package Code", ExcelSetup.Code);

        MalL.SETRANGE(MalL."Include Field", TRUE);

        MalFilter.SETRANGE("Table ID", ExcelSetup.TableID);
        MalFilter.SETRANGE("Package Code", ExcelSetup.Code);
        EnterCell(1, 1, ExcelSetup.code, TRUE, FALSE, FALSE);
        Entercell(1, 2, ExcelSetup.TableName, TRUE, FALSE, FALSE);
        EnterCell(1, 3, Format(ExcelSetup.TableID), TRUE, FALSE, FALSE);
        i := 1;

        if MalL.FINDFIRST() then
            repeat
                //MalL.CALCFIELDS(FieldName);
                Felt[i] := MalL."Field ID";
                EnterCell(3, i, MalL."Field Caption", TRUE, FALSE, TRUE);
                //EnterCell(1, i, MalL.FieldName, TRUE, FALSE, FALSE);
                i += 1;
            until MalL.NEXT() = 0;

        if MalFilter.FINDFIRST() then
            repeat
                tmpfr := tmpRec.FIELD(MalFilter."Field ID");
                tmpfr.SETFILTER(MalFilter."Field Filter");
            until MalFilter.NEXT() = 0;


        rad := 4;
        if tmpRec.FINDFIRST() then
            repeat
                Window.UPDATE(2, FORMAT(rad));
                i := 1;
                if MalL.FINDFIRST() then
                    repeat
                        if mall."Field ID" <> 99601 then
                            f2e(rad, i, tmpRec.FIELD(MalL."Field ID"));

                        i += 1;
                    until MalL.NEXT() = 0;
                rad += 1;

            until tmpRec.NEXT() = 0;
        if (ExcelSetup.TableID = 81) and (ExcelSetup."Journal Batch Name" <> '') then
            SetupIBJournal(ExcelSetup);
        Window.CLOSE();
        if ExcelSetup."Opprett nytt regneark" then begin
            TempExcelBuff.WriteSheet(ExcelSetup.SheetName, CompanyName, UserId);
            TempExcelBuff.CloseBook();
            if Strpos(ExcelSetup.FileName, ' (') > 1 then
                ExcelSetup.FileName := CopyStr(CopyStr(ExcelSetup.FileName, 1, strpos(ExcelSetup.FileName, ' (') - 1), 1, MaxStrLen(ExcelSetup.FileName));
            if strpos(ExcelSetup.FileName, '.') > 1 then
                ExcelSetup.FileName := CopyStr(CopyStr(ExcelSetup.FileName, 1, strpos(ExcelSetup.FileName, '.') - 1), 1, MaxStrLen(ExcelSetup.FileName));

            TempExcelBuff.SetFriendlyFilename(ExcelSetup.FileName);
            TempExcelBuff.OpenExcel();

        end;


    end;

    local procedure SetupIBJournal(lExcelSetup: record "Excel export Setup")

    begin
        if Rad > 4 then
            EXIT;
        mall.SetRange("Primary Key", TRUE);
        if Mall.FindFirst() then
            repeat

                case mall."Field ID" of
                    1:
                        EnterCell(rad, 1, lExcelSetup."Journal Template Name", false, FALSE, FALSE);
                    2:
                        EnterCell(rad, 2, Format(10000), false, FALSE, FALSE);


                    51:
                        EnterCell(rad, FindCol(51), lExcelSetup."Journal Batch Name", false, FALSE, FALSE);
                    else
                        EnterCell(rad, 1, '', false, FALSE, FALSE);
                end;
            until Mall.Next() = 0;
        EnterCell(rad, FindCol(5), Format(lExcelSetup.IBDate, 0, 1), false, FALSE, FALSE);
        //EnterCell(rad, 4, lExcelSetup."Journal Batch Name", false, FALSE, FALSE);
    end;

    procedure FindCol(aFieldNo: Integer) Result: Integer
    var
        i: Integer;
    begin
        Result := 0;
        for i := 1 to 50 do
            if Felt[i] = aFieldNo then
                Result := i;
    end;

    procedure Copy2Table(aExcelSetup: Record "Excel Export Setup")
    var


        // felt: array[50] of Integer;
        ValiderFelt: array[50] of Integer;
        KolV: Integer;
        Window: Dialog;

        AntallFeil: Integer;
        LogFileOpen: Boolean;

        tabChar: Text[1];
        KeyFelt: array[10] of Integer;
        KolTreat: array[50] of Integer;
        KolK: Integer;
        lTableID: Integer;
    begin
        Window.OPEN('Overføring til tabell\' +
        '#1#########\' +
        '#2######### ');
        AntallFeil := 0;
        AnyTable.CLOSE();
        LogFileOpen := FALSE;
        AnyTable.OPEN(aExcelSetup.TableID);
        rad := 4;
        Window.UPDATE(1, FORMAT(aExcelSetup.TableID));
        Window.UPDATE(2, FORMAT(rad));
        CLEAR(felt);
        CLEAR(ValiderFelt);
        CLEAR(KeyFelt);
        CLEAR(KolTreat);
        mall.SetCurrentKey("Processing Order");
        MalL.SETRANGE("Table ID", aExcelSetup.TableID);
        MalL.SETRANGE("Package Code", aExcelSetup.Code);
        MalL.SETRANGE("Include Field", TRUE);
        kol := 1;
        KolV := 1;
        KolK := 1;
        if MalL.FINDFIRST() then
            repeat
                felt[kol] := MalL."Field ID";
                KolTreat[Kol] := MalL.Treatment;
                kol += 1;
                if MalL."Validate Field" then begin
                    ValiderFelt[KolV] := MalL."Field ID";
                    KolV += 1;
                end;
                if MalL."Primary Key" then begin
                    KeyFelt[KolK] := MalL."Field ID";
                    KolK += 1;
                end;



            until MalL.NEXT() = 0;

        CheckAndUpdate.SetFelt(felt, ValiderFelt, KeyFelt, KolTreat);
        CheckAndUpdate.SetExcelBuf(TempExcelBuff);
        aExcelSetup.IBLineNo := 10000;
        if TempExcelBuff.FINDFIRST() then
            repeat
                if (TempExcelBuff."Row No." = 1) and (TempExcelBuff."Column No." = 3) then begin
                    if EVALUATE(lTableID, TempExcelBuff."Cell Value as Text") then
                        if lTableID <> aExcelSetup.TableID then begin
                            Window.CLOSE();
                            MESSAGE(Text0001Msg);
                            EXIT;
                        end;


                end
                else
                    if TempExcelBuff."Row No." = rad then begin

                        aExcelSetup.ExcelBufRow := rad;
                        aExcelSetup.KeyValue := '';
                        if TempExcelBuff."Column No." = 1 then
                            aExcelSetup.KeyValue := CopyStr(TempExcelBuff."Cell Value as Text", 1, MaxStrLen(aExcelSetup.KeyValue));
                        rad += 1;
                        COMMIT();
                        if not CheckAndUpdate.RUN(aExcelSetup) then
                            if aExcelSetup.ImportLogFil <> '' then begin
                                tabChar[1] := 9;
                                Message(GetLastErrorText);
                                if not LogFileOpen then;

                                LogFileOpen := TRUE;


                                AntallFeil += 1;

                            end;
                        //aExcelSetup.IBLineNo += 10;
                        COMMIT();
                        Window.UPDATE(2, FORMAT(rad));
                    end;

            until TempExcelBuff.NEXT() = 0;
        Window.CLOSE();
        AnyTable.CLOSE();

        if AntallFeil > 0 then
            MESSAGE('Feil ved %1 poster. Siste melding\' +
                    '%2', AntallFeil, GETLASTERRORTEXT);
    end;

    local procedure EnterCell(RowNo: Integer; ColumnNo: Integer; CellValue: Text[250]; Bold: Boolean; Italic: Boolean; UnderLine: Boolean)
    begin
        TempExcelBuff.INIT();
        TempExcelBuff.VALIDATE("Row No.", RowNo);
        TempExcelBuff.VALIDATE("Column No.", ColumnNo);
        TempExcelBuff."Cell Value as Text" := CellValue;
        TempExcelBuff.Formula := '';
        TempExcelBuff.Bold := Bold;
        TempExcelBuff.Italic := Italic;
        TempExcelBuff.Underline := UnderLine;
        //TempExcelBuff."Cell Type" := TempExcelBuff."Cell Type"::
        TempExcelBuff.INSERT(true);
    end;


    procedure f2e(row: Integer; col: Integer; lField: FieldRef)
    var
        s: Text[250];
    begin
        if FORMAT(lField.CLASS) = 'FlowField' then
            lField.CALCFIELD();
        case FORMAT(lField.TYPE) of
            'Code', 'Text':
                s := FORMAT(lField.VALUE);  //Geirb, 10.aug.12
            else
                s := FORMAT(lField.VALUE, 0, 1);
        end;
        EnterCell(row, col, s, FALSE, FALSE, FALSE);
    end;


    procedure e2f(row: Integer; lCol: Integer; aField: FieldRef; s: Text[150])
    var
        tmpdec: Variant;

    begin
        case FORMAT(aField.TYPE) of
            'Decimal':
                EVALUATE(tmpdec, s);

            else
        end;
    end;


    procedure SetValue(ExcelStr: Text[180]; fRef: FieldRef)
    var

        Value: Variant;
    begin
        CASE FORMAT(fRef.TYPE) of
            'Decimal', 'Date', 'Time', 'Integer':
                EVALUATE(Value, ExcelStr);
            else
                Value := ExcelStr
        end;
        fRef.VALUE(Value);
    end;


    procedure SplittStreng(s: Text[1024]; var f: array[150] of Text[250])
    var
        i: Integer;
        p: Integer;
    begin
        for i := 1 to 150 do
            f[i] := '';
        p := 1;

        for i := 1 to STRLEN(s) do
            if p < 150 then
                if s[i] = ';' then
                    p := p + 1
                else
                    f[p] := f[p] + COPYSTR(s, i, 1);

    end;


    procedure VelgFelt(ExcelSetup: Record "Excel Export Setup")
    var
        ValgtFelt: Record "Config. Package Field";
        allField: Record 2000000041;
    begin

        ExcelSetup.TESTFIELD(Code);
        ExcelSetup.TESTFIELD(TableID);
        ValgtFelt.RESET();
        ValgtFelt.SETRANGE("Table ID", ExcelSetup.TableID);
        ValgtFelt.SETRANGE("Package Code", ExcelSetup.Code);
        if NOT ValgtFelt.FINDFIRST() then begin
            allField.RESET();
            allField.SETRANGE(TableNo, ExcelSetup.TableID);
            allField.SETRANGE(Enabled, TRUE);
            if allField.FindSet() then
                repeat
                    ValgtFelt.INIT();
                    ValgtFelt."Table ID" := ExcelSetup.TableID;
                    ValgtFelt."Package Code" := ExcelSetup.Code;
                    ValgtFelt."Field ID" := allField."No.";
                    ValgtFelt.INSERT(true);
                until allField.NEXT() = 0;
        end;
        COMMIT();
        //FORM.RUNMODAL(84501, ValgtFelt);
    end;


    procedure GetExcelDocument(lDialogText: Text; var lFileName: Text[250]; var lSheetName: Text[50])
    var
        aFileName: Text;
    begin
        //TempBlob.Blob.CREATEINSTREAM(lExcelStream);

        if UPLOADINTOSTREAM(lDialogText, '', 'Excel Files|*.xlsx|CSV Files|*.csv', aFileName, lExcelStream) then begin
            lFileName := CopyStr(aFileName, 1, MaxStrLen(lFileName));
            gFileName := lFileName;
            lSheetName := CopyStr(GetSheetsName(gFileName), 1, MaxStrLen(lSheetName));

        end
        else
            gFileName := ''
    end;



    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Import XML File to Data Exch.", OnParseParentChildLineOnBeforeInsertColumn, '', false, false)]

    local procedure OnParseParentChildLineOnBeforeInsertColumn(var InnerText: Text; InnerXML: Text; OuterXML: Text; DataExchColumnDef: Record "Data Exch. Column Def")
    var
        s: Text[250];
    begin
        s := copystr(StrSubstNo('%1 %2 %3', InnerText, InnerXML, OuterXML), 1, 250);

    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Map DataExch To Intermediate", OnBeforeIntermediateDataImportInsert, '', false, false)]

    local procedure OnBeforeIntermediateDataImportInsert(DataExchField: Record "Data Exch. Field"; DataExchLineDef: Record "Data Exch. Line Def"; var TempNameValueBuffer: Record "Name/Value Buffer" temporary; var IntermediateDataImport: Record "Intermediate Data Import")

    begin

        if (IntermediateDataImport."Table ID" = 39) and (IntermediateDataImport."Field ID" = 5725) then
            CreateItem(IntermediateDataImport);

    end;


    local procedure CreateItem(InterMediateDataImport: Record "Intermediate Data Import")
    var
        Item: Record Item;
        ItemRef: record "Item Reference";
    begin
        if not Item.Get(InterMediateDataImport.Value) then begin
            Item.Init();
            CopyFromTemplate(Item, 'DETALJ');
            Item."No." := CopyStr(InterMediateDataImport.Value, 1, MaxStrLen(Item."No."));
            Item."Vendor Item No." := CopyStr(InterMediateDataImport.Value, 1, MaxStrLen(Item."Vendor Item No."));
            InterMediateDataImport.SetRange("Record No.", InterMediateDataImport."Record No.");
            if InterMediateDataImport.FindSet() then
                repeat
                    if InterMediateDataImport."Field ID" = 11 then
                        Item.Description := CopyStr(InterMediateDataImport.Value, 1, MaxStrLen(Item.Description));
                    if InterMediateDataImport."Field ID" = 5407 then
                        Item."Base Unit of Measure" := CopyStr(InterMediateDataImport.Value, 1, MaxStrLen(Item."Base Unit of Measure"));
                    if InterMediateDataImport."Field ID" = 29 then
                        Item."Unit Cost" := GetDecimal(InterMediateDataImport.Value);
                    if InterMediateDataImport."Field ID" = 86 then
                        item."Vendor No." := CopyStr(InterMediateDataImport.Value, 1, MaxStrLen(Item."Vendor No."));
                // Item."Unit Price" := Str2Decimal(InterMediateDataImport.Value);
                until InterMediateDataImport.NEXT() = 0;
            InterMediateDataImport.Reset();
            InterMediateDataImport.SetRange("Table ID", 23);
            InterMediateDataImport.SetRange("Field ID", 86);
            if InterMediateDataImport.FindFirst() then
                Item."Vendor No." := CopyStr(InterMediateDataImport.Value, 1, MaxStrLen(Item."Vendor No."));
            Item."Item Category Code" := '';
            Item.Validate("Unit Cost");
            Item."Base Unit of Measure" := CopyStr(GetUnitCode(Item."Base Unit of Measure"), 1, MaxStrLen(Item."Base Unit of Measure"));
            UpdateItemUnitOfMeasure(Item);
            item.Validate("Vendor No.", GetVendorNo(Item."Vendor No."));
            Item.Validate("Vendor No.", Item."Vendor No.");
            Item.Validate("Base Unit of Measure");
            Item.Insert();
            ItemRef.Init();
            ItemRef."Item No." := Item."No.";
            ItemRef."Reference Type" := ItemRef."Reference Type"::Vendor;
            ItemRef."Reference Type No." := Item."Vendor No.";

            ItemRef."Unit of Measure" := Item."Base Unit of Measure";
            ItemRef."Reference No." := Item."Vendor Item No.";

            if ItemRef.Insert() then;

        end;
    end;

    local procedure GetDecimal(s: Text) Result: Decimal

    begin
        s := ConvertStr(s, '.', ',');
        Evaluate(Result, s);

    end;

    procedure GetVendorNo(lOrdNo: Text[20]) Result: Text[20]
    var
        Vendor: Record Vendor;
    begin
        Vendor.reset();
        Vendor.SetRange("VAT Registration No.", lOrdNo);
        if Vendor.FindFirst() then
            Result := Vendor."No."
        else
            Result := '';

    end;

    local procedure CopyFromTemplate(var Item: Record Item; TemplateName: Text[30])
    var
        ItemTemplate: Record "Item Templ.";
    begin
        if ItemTemplate.Get(TemplateName) then begin
            Item.Validate("Inventory Posting Group", ItemTemplate."Inventory Posting Group");

            Item.Validate("Gen. Prod. Posting Group", ItemTemplate."Gen. Prod. Posting Group");
            Item.Validate("VAT Prod. Posting Group", ItemTemplate."VAT Prod. Posting Group");

        end;
    end;

    local procedure UpdateItemUnitOfMeasure(Item: Record Item)
    var
        ItemUnitOfMeasure: Record "Item Unit of Measure";
    begin
        if not ItemUnitOfMeasure.Get(Item."No.", Item."Base Unit of Measure") then begin
            ItemUnitOfMeasure.Init();
            ItemUnitOfMeasure."Item No." := Item."No.";
            ItemUnitOfMeasure.Code := Item."Base Unit of Measure";
            ItemUnitOfMeasure.Insert();
        end;


    end;

    procedure GetUnitCode(lUnit: Text[20]) Result: Text[20]
    var
        Unit: Record "Unit of Measure";
    begin
        Unit.reset();
        Unit.SetRange("Code", lUnit);
        if Unit.FindFirst() then
            Result := Unit.Code
        else begin
            Unit.Reset();
            Unit.SetRange("International Standard Code", lUnit);
            if Unit.FindFirst() then
                Result := Unit.Code
            else
                Result := '';
        end;

    end;
}

