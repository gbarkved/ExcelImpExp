namespace cepheo.ExcelImportExport;
using System.IO;
codeunit 88603 DXTEXLXMLImport
{


    trigger OnRun()
    begin

    end;

    procedure ImportXMLToBuffer(lInStr: InStream)
    var
        XmlDoc: XmlDocument;
        XmlNode: XmlNode;
        XmlElement: XmlElement;
        lNodeList: XmlNodeList;

        InStr: InStream;



    begin

        if not XmlDocument.ReadFrom(InStr, XMLDoc) then
            error('Ikke XML dokument');
        //XmlDocument.ReadFrom(XmlData, XmlDoc);
        XmlDoc.GetRoot(XmlElement);
        lNodeList := XmlElement.GetChildNodes();
        XmlNode := XmlElement.AsXmlNode();

        // Start å fylle XML Buffer her
        //FillXMLBuffer(lNodelist, XmlNode, 0);
        Parse1Level(lNodeList, 1, '', 1, 0);
        Message('XML data imported to buffer successfully.');
        message('Total records imported: ' + Format(TempXMLBuffer.Count()));
    end;


    procedure Parse1Level(lNodeList: XmlNodeList; lNivaa: integer; lHeading: text; NodeID: Integer; ParentNodeId: Integer)
    var

        lNode: XmlNode;
        lElement: XmlElement;

        lNodeList2: XmlNodeList;


        test: text;
        I: integer;


    begin
        lNivaa := lNivaa;


        foreach lNode in lNodeList do begin

            lElement := lNode.AsXmlElement();
            test := lNode.AsXmlElement().name;

            for i := lNivaa to 20 do
                MainName[i] := '';
            if lNode.AsXmlElement().HasElements() then
                MainName[lNivaa] := lNode.AsXmlElement().name;
            lHeading := '';

            for I := 1 to 20 do
                if MainName[i] <> '' then
                    lHeading += '/' + MainName[i];


            lNodeList2 := lElement.GetChildElements();

            if lNode.AsXmlElement().HasElements() then
                Parse1Level(lNodeList2, lNivaa + 1, lHeading, NodeId, ParentNodeId)
            else begin
                TempXMLBuffer.Init();
                TempXMLBuffer."Entry No." := NodeId; // Bruk NodeId som Entry No.
                TempXMLBuffer."Node Number" := NodeId; // Bruk NodeId som Node Number
                TempXMLBuffer."Parent Entry No." := ParentNodeId;

                if lNode.AsXmlElement().HasElements then begin
                    TempXMLBuffer.name := copyStr(lNode.AsXmlElement().Name, 1, 250);
                    TempXMLBuffer.Value := copyStr(lNode.AsXmlElement().InnerText, 1, 250);
                end;

                TempXMLBuffer.Insert();
            end;
            //else

            //lRec.Field(lBCFieldNO).Value := lNode.AsXmlElement().InnerText;







        end;


    end;

    procedure FillXMLBuffer(lNodeList: XmlNodeList; Node: XmlNode; ParentNodeId: Integer)
    var
        ChildNode: XmlNode;
        NodeId: Integer;
    begin
        NodeId := ParentNodeId; // Generer neste unike ID for noden
        if nodeid <> 0 then begin


            TempXMLBuffer.Init();
            TempXMLBuffer."Entry No." := NodeId; // Bruk NodeId som Entry No.
            TempXMLBuffer."Node Number" := NodeId; // Bruk NodeId som Node Number
            TempXMLBuffer."Parent Entry No." := ParentNodeId;

            if Node.AsXmlElement().HasElements then begin
                TempXMLBuffer.name := copyStr(Node.AsXmlElement().Name, 1, 250);
                TempXMLBuffer.Value := copyStr(Node.AsXmlElement().InnerText, 1, 250);
            end;

            TempXMLBuffer.Insert();
        end;
        NodeId += 1; // Inkrementer NodeId for neste node

        foreach ChildNode in lNodeList do
            FillXMLBuffer(lNodeList, ChildNode, NodeId);

    end;

    var
        TempXMLBuffer: Record "XML Buffer" temporary;
        MainName: array[20] of text;
}