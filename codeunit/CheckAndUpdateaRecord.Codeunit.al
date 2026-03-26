namespace cepheo.ExcelImportExport;

using Microsoft.Finance.GeneralLedger.Journal;
using Microsoft.Inventory.Item;
using Microsoft.Inventory.item.Catalog;
using Microsoft.Purchases.Vendor;
using Microsoft.Sales.Customer;
using System.IO;


codeunit 88601 CheckAndUpdateaRecord
{
    // ///*SAK-14635-FTC2 Navicom GeirB 27.03.2012: Utvidet rutinene til å søke opp 10 nøkler.

    TableNo = 88600;

    trigger OnRun()
    var
        LagrePost: Boolean;
        CustMal: Code[20];
        IK: Integer;
    begin
        CLEAR(anyField);
        anyRecord.CLOSE();
        anyRecord.OPEN(rec.TableID);

        //KeyField := anyRecord.KeyIndex(1);
        // for i := 0 to anyrecord.KeyIndex(1).FieldCount - 1 do
        //    KeyFelt[i + 1] := anyRecord.Field(i).Number; //(1).Fieldindex(i).Number;
        //SetFelt();
        // ValidateFelt
        anyRecord.init();
        if (rec.TableID = 81) and (rec."Journal Batch Name" <> '') then
            UpdateFromIBSetup(rec.ExcelBufRow, rec."Journal Template Name", rec."Journal Batch Name", rec.IBLineNo);
        LagrePost := GetRecord(rec.ExcelBufRow, rec.TableID);

        TempexcelBuf.SETRANGE("Row No.", rec.ExcelBufRow);
        i := 1;
        custmal := '';
        if TempexcelBuf.FINDFIRST() then
            repeat

                i := TempexcelBuf."Column No.";


                if (felt[i] = 99601) then
                    custmal := CopyStr(TempexcelBuf."Cell Value as Text", 1, MaxStrLen(CustMal))
                else
                    if (Felt[i] > 0) then begin
                        anyField := anyRecord.FIELD(Felt[i]);
                        case FORMAT(anyField.TYPE) of
                            'Text', 'Code':
                                if UpdateField(anyField.VALUE, TempexcelBuf."Cell Value as Text") then begin
                                    LagrePost := true;
                                    anyField.VALUE := TempexcelBuf."Cell Value as Text";
                                end;

                            'Integer', 'Option':
                                if UpdateField(anyField.VALUE, TempexcelBuf."Cell Value as Text") then begin
                                    anyField.VALUE := GetInteger(TempexcelBuf."Cell Value as Text");
                                    LagrePost := true;
                                end;
                            'Decimal':
                                if UpdateField(anyField.VALUE, TempexcelBuf."Cell Value as Text") then begin
                                    anyField.VALUE := GetDecimal(TempexcelBuf."Cell Value as Text");
                                    LagrePost := true;
                                end;
                            'Date':
                                if UpdateField(anyField.VALUE, TempexcelBuf."Cell Value as Text") then begin
                                    anyField.VALUE := GetDate(TempexcelBuf."Cell Value as Text");
                                    LagrePost := true;
                                end;
                            'Boolean':
                                if UpdateField(anyField.VALUE, TempexcelBuf."Cell Value as Text") then begin
                                    anyField.VALUE := GetBoolean(TempexcelBuf."Cell Value as Text");
                                    LagrePost := true;
                                end;
                            'Time':
                                if UpdateField(anyField.VALUE, TempexcelBuf."Cell Value as Text") then begin
                                    anyField.VALUE := GetTime(TempexcelBuf."Cell Value as Text");
                                    LagrePost := true;
                                end;
                            else
                        end;

                        if Felt[i] = 99601 then
                            CustMal := CopyStr(TempexcelBuf."Cell Value as Text", 1, MaxStrLen(CustMal));
                        if KolTreat[i] > 0 then
                            UpdateExternalRecord(KolTreat[i], Felt[i], anyRecord);
                        for ik := 1 to 50 do
                            if (ValidateFelt[ik] <> 99601) then
                                if anyField.NUMBER = ValidateFelt[ik] then
                                    anyField.VALIDATE();

                    end;


                i += 1;
            until TempexcelBuf.NEXT() = 0;
        UpdateFromTemplate(rec.TableID, CustMal);
        if LagrePost then
            //if not anyRecord.insert(true) then
                if anyRecord.MODIFY(true) then;
        anyRecord.CLOSE();

    end;

    var
        TempexcelBuf: Record "Excel Buffer" temporary;

        anyRecord: RecordRef;
        anyField: FieldRef;
        //KeyField: FieldRef;
        i: Integer;
        //tmp: RecordID;

        Felt: array[50] of Integer;
        KolTreat: array[50] of Integer;
        ValidateFelt: array[50] of Integer;
        KeyFelt: array[10] of Integer;

    local procedure UpdateFromIBSetup(aExcelBufRow: integer; JournalTemplateName: Text[30]; JournalBatchName: Code[10]; var LLineNo: integer)
    var
        lanyField: FieldRef;

        kp: Integer;
        ExcelValue: Text[250];


    begin
        TempexcelBuf.SETRANGE("Row No.", aExcelBufRow);

        anyRecord.INIT();
        for kp := 1 to 10 do
            if KeyFelt[kp] <> 0 then begin
                lanyField := anyRecord.FIELD(KeyFelt[kp]);
                TempexcelBuf.SETRANGE("Column No.", FindCol(KeyFelt[kp]));
                if KeyFelt[kp] in [1, 2, 51] then begin
                    case KeyFelt[kp] of
                        1:

                            ExcelValue := JournalTemplateName;
                        2:
                            begin
                                ExcelValue := Format(lLineNo);
                                LLineNo += 1000;
                            end;


                        51:

                            ExcelValue := JournalBatchName;
                        else
                            ExcelValue := '';
                    end;
                    TempexcelBuf."Row No." := aExcelBufRow;
                    TempexcelBuf."Cell Value as Text" := ExcelValue;
                    TempexcelBuf."Column No." := FindCol(KeyFelt[kp]);
                    if not TempexcelBuf.Get(aExcelBufRow, FindCol(KeyFelt[kp])) then
                        TempexcelBuf.insert()
                    else begin
                        TempexcelBuf."Cell Value as Text" := ExcelValue;
                        TempexcelBuf.MODIFY();
                    end;


                end;
            end;
        TempexcelBuf.setrange("Column No.");

    end;

    local procedure UpdateFromTemplate(lTableID: Integer; lCustTemp: Code[20])
    var

        TemplateRecRef: RecordRef;
        lanyField: FieldRef;
        Ip: Integer;
    begin
        TemplateRecRef.close();
        case lTableID of
            18:
                TemplateRecRef.open(Database::"Customer Templ.");
            23:
                TemplateRecRef.open(Database::"Vendor Templ.");
            27:
                TemplateRecRef.open(Database::"Item Templ.");
            else
                exit;
        end;

        lanyField := TemplateRecRef.FIELD(1);
        lanyField.SetRange(lCustTemp);
        if TemplateRecRef.FindFirst() then
            for ip := 2 to TemplateRecRef.FieldCount() - 1 do
                if UpdateField(TemplateRecRef.Fieldindex(ip).number) then
                    if TemplateRecRef.fieldindex(ip).Number < 99998500 then begin
                        lanyField := TemplateRecRef.Field(TemplateRecRef.FieldIndex(ip).Number);
                        If anyRecord.FieldExist(TemplateRecRef.FieldIndex(ip).number) then
                            anyRecord.Field(TemplateRecRef.FieldIndex(ip).Number).VALUE := lanyField.VALUE;
                    end;



        TemplateRecRef.close();
    end;

    local procedure UpdateField(lFeltNo: Integer): Boolean
    var

        ip: Integer;
    begin
        for ip := 1 to 50 do
            if Felt[ip] = lFeltNo then
                exit(false);
        Exit(true);

    end;

    procedure SetFelt(aFeltRad: array[50] of Integer; aValidateFelt: array[50] of Integer; akeyFelt: array[10] of Integer; aKolTreat: array[50] of Integer)
    var
        ip: Integer;
    begin
        for ip := 1 to 50 do begin
            Felt[ip] := aFeltRad[ip];
            ValidateFelt[ip] := aValidateFelt[ip];
            KolTreat[ip] := aKolTreat[ip];
        end;
        for ip := 1 to 10 do
            KeyFelt[ip] := akeyFelt[ip];
        // Clear(KeyFelt);

    end;


    procedure UpdateField(v: Variant; s: Text[250]) Result: Boolean
    begin
        Result := false;
        if v.ISINTEGER then
            Result := FORMAT(v) <> s
        else
            if v.ISTEXT OR v.ISCODE then
                Result := FORMAT(v) <> s
            else
                if v.ISDECIMAL then
                    Result := FORMAT(v, 0, 1) <> s
                else
                    if v.ISDATE then
                        Result := FORMAT(v) <> s
                    else
                        if v.ISBOOLEAN then
                            Result := FORMAT(v) <> s
                        else
                            if v.ISTIME then
                                Result := FORMAT(v) <> s
                            else
                                if v.IsDateTime then
                                    Result := FORMAT(v) <> s;
    end;


    procedure GetInteger(s: Text[250]) Result: Integer
    begin
        Result := 0;
        if EVALUATE(Result, s) then;
    end;


    procedure GetDecimal(s: Text) Result: Decimal
    begin
        Result := 0;
        if EVALUATE(Result, s) then;
    end;


    procedure GetDate(s: Text) Result: Date
    begin
        Result := 0D;
        if EVALUATE(Result, s) then;
    end;


    procedure GetTime(s: Text) result: Time
    var
        d: Decimal;
        time: Integer;
        "min": Integer;
        sek: Integer;
        sTime: Text[30];
    begin
        result := 0T;
        if EVALUATE(d, s) then begin
            time := ROUND(d * 24, 1, '<');
            d := (d * 24 - time) * 60;
            min := ROUND(d, 1, '<');
            d := (d - min) * 60;
            sek := ROUND((d - min) * 60, 1, '<');
            sTime := FORMAT(time, 2, '<Int>') + ':' + FORMAT(min, 2, '<int>') + ':' + FORMAT(sek, 2, '<int>');
            sTime := CONVERTSTR(sTime, ' ', '0');
            if EVALUATE(result, sTime) then;
            //  sTime := Format(
            //  Result := time;
        end;
    end;


    procedure GetBoolean(s: Text) Result: Boolean
    begin
        Result := FALSE;
        Result := (s = 'Ja')
    end;

    procedure GetOptionID(lanyField: FieldRef; s: Text) Result: Integer
    begin
        Result := 0;
        if lanyField.TYPE = lanyField.TYPE::Option then
            for i := 1 to lanyField.EnumValueCount() do
                if lanyField.GetEnumValueCaption(i) = s then begin
                    Result := i - 1;
                    exit;
                end;

    end;

    procedure GetRecord(aExcelBufRow: Integer; aTableID: Integer) Result: Boolean
    var
        kp: Integer;

    begin
        TempexcelBuf.SETRANGE("Row No.", aExcelBufRow);
        if TempexcelBuf.FINDFIRST() then begin


            Result := FALSE;
            anyRecord.INIT();
            for kp := 1 to 10 do
                if KeyFelt[kp] <> 0 then begin
                    anyField := anyRecord.FIELD(KeyFelt[kp]);
                    TempexcelBuf.SETRANGE("Column No.", FindCol(KeyFelt[kp]));
                    if TempexcelBuf.FINDFIRST() then
                        case FORMAT(anyField.TYPE) of
                            'Code', 'Text':
                                begin
                                    anyField.SETRANGE(TempexcelBuf."Cell Value as Text");
                                    anyField.VALUE(FORMAT(TempexcelBuf."Cell Value as Text"));

                                end;
                            'Integer':
                                begin
                                    anyField.SETRANGE(GetInteger(TempexcelBuf."Cell Value as Text"));
                                    anyField.VALUE(GetInteger(TempexcelBuf."Cell Value as Text"));
                                end;
                            'Option':
                                anyField.SETRANGE(GetOptionID(anyField, TempexcelBuf."Cell Value as Text"));
                            //anyField.VALUE(GetInteger(TempexcelBuf."Cell Value as Text"));

                            'Date':
                                begin
                                    anyField.SETRANGE(GetDate(TempexcelBuf."Cell Value as Text"));
                                    anyField.VALUE(GetDate(TempexcelBuf."Cell Value as Text"));
                                end;
                            'Decimal':
                                begin
                                    anyField.SETRANGE(GetDecimal(TempexcelBuf."Cell Value as Text"));
                                    anyField.VALUE(GetDecimal(TempexcelBuf."Cell Value as Text"));
                                end;
                            else begin
                                anyField.SETRANGE(TempexcelBuf."Cell Value as Text");
                                anyField.VALUE(TempexcelBuf."Cell Value as Text");
                            end;

                        end
                    else
                        anyField.SETRANGE();

                end;
            if NOT anyRecord.FINDFIRST() then begin
                if aTableID = 81 then
                    PrepareGenJournalLine(anyRecord);
                anyRecord.INSERT(true);
                Result := true;
            end;
            TempexcelBuf.SETRANGE("Column No.");
        end;
        //anyField := anyRecord.FIELD(1);
        /*LagrePost := FALSE;
        if FORMAT(anyField.TYPE) = 'Integer' then
          anyField.SETRANGE(GetInteger(KeyValue))
        else
          anyField.SETRANGE(KeyValue);
        if KeyValue = '' then
         anyRecord.INSERT(true)
        else
        if NOT anyRecord.FINDFIRST then
        begin
         anyRecord.INIT;
         if FORMAT(anyField.TYPE) = 'Integer' then
          anyField.VALUE := GetInteger(KeyValue)
         else
          anyField.VALUE := KeyValue;
         anyRecord.INSERT(true);
         LagrePost := true;
        end;*/

    end;


    procedure FindCol(aFieldNo: Integer) Result: Integer
    begin
        Result := 0;
        for i := 1 to 50 do
            if Felt[i] = aFieldNo then
                Result := i;
    end;

    procedure SetExcelBuf(var ExcelBuffTemp: Record "Excel Buffer" temporary)
    begin
        TempexcelBuf.deleteall();
        ExcelBuffTemp.reset();
        if ExcelBuffTemp.FindSet() then
            repeat
                TempexcelBuf.Copy(ExcelBuffTemp);
                if TempexcelBuf.insert() then;
            until ExcelBuffTemp.Next() = 0;

    end;

    local procedure PrepareGenJournalLine(var lAnyRecord: RecordRef)
    var
        GenJournalLineTemplate: Record "Gen. Journal Line";
        GenJournalLine: Record "Gen. Journal Line";

    begin
        GenJournalLineTemplate."Journal Template Name" := lAnyRecord.Field(1).VALUE;
        GenJournalLineTemplate."Journal Batch Name" := lAnyRecord.Field(51).VALUE;
        lAnyRecord.settable(GenJournalLine);
        GenJournalLine.Reset();
        GenJournalLine."Posting Date" := today;
        GenJournalLineTemplate.SetUpNewLine(GenJournalLine, GenJournalLine."Balance (LCY)", true);
        GenJournalLineTemplate."Account Type" := GenJournalLineTemplate."Account Type"::"G/L Account";

        GenJournalLine.SetRange("Journal Template Name", GenJournalLine."Journal Template Name");
        GenJournalLine.SetRange("Journal Batch Name", GenJournalLine."Journal Batch Name");
        if GenJournalLine.FindLast() then begin
            GenJournalLineTemplate."Line No." := GenJournalLine."Line No.";
            GenJournalLineTemplate."Document No." := (GenJournalLine."Document No.");
        end;

        lAnyRecord.field(7).Value := GenJournalLineTemplate."Document No.";

    end;

    procedure UpdateExternalRecord(lTID: integer; lFieldNo: Integer; RRec: RecordRef)
    var
        ItemUnit: record "Item Unit of Measure";
        ItemReference: record "Item Reference";
    begin
        if lTid in [101] then begin
            itemunit.init();
            ItemUnit."Item No." := RRec.Field(1).VALUE;

            itemUnit.Code := RRec.Field(lFieldNo).VALUE;

            if ItemUnit.Insert() then;
        end;
        if lTid = 102 then begin
            ItemReference.init();
            ItemReference."Item No." := RRec.Field(1).VALUE;
            ItemReference."Reference Type" := ItemReference."Reference Type"::Vendor;
            ItemReference."Reference Type No." := RRec.Field(31).VALUE; //Vendor No.
            ItemReference."Unit of Measure" := RRec.Field(8).VALUE; //Base Unit of Measure
            ItemReference."Reference No." := RRec.Field(lFieldNo).VALUE; //Vendor Item No.

            if ItemReference.Insert() then;
        end;

    end;
}