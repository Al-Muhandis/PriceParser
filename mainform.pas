unit MainForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, EditBtn, StdCtrls, ExtCtrls, Grids, ComCtrls, SynEdit,
  SynHighlighterXML;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    FileNameEdit1: TFileNameEdit;
    Label1: TLabel;
    StringGrid1: TStringGrid;
    ToolBar1: TToolBar;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private

  public

  end;

var
  Form1: TForm1;

implementation

uses
  Laz2_DOM, laz2_XMLRead, laz2_xmlutils, StrUtils, eventlog
  ;

{$R *.lfm}

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
var
  aDoc: TXMLDocument;
  aContentNode, aNode: TDOMNode;
  aLogger: TEventLog;
  aRow: Integer;

  function ParseFromOneTable(aTableNode: TDOMNode): Boolean;
  var
    aArticle, aName, aPV, aBV,aPPrice, aCPrice: String; 
    aNodeCell, aNodeRow: TDOMNode;
    i: Integer;
  begin
    Result:=False;
    if not aTableNode.HasChildNodes then
      Exit;
    aNodeRow:=aTableNode.FirstChild;
    i:=0;
    while Assigned(aNodeRow) do
    begin
      while not aNodeRow.HasChildNodes do
        aNodeRow:=aNodeRow.NextSibling;
      if i>-1 then
      begin
        aNodeCell:=aNodeRow.FirstChild;
        aArticle:=Trim(aNodeCell.TextContent);
        aNodeCell:=aNodeCell.NextSibling;
        aName:=Trim(DelSpace1(aNodeCell.TextContent));
        aNodeCell:=aNodeCell.NextSibling;
        aPV:=Trim(aNodeCell.TextContent);
        aNodeCell:=aNodeCell.NextSibling;
        aBV:=Trim(aNodeCell.TextContent);
        aNodeCell:=aNodeCell.NextSibling;
        aNodeCell:=aNodeCell.NextSibling;
        aPPrice:=Trim(aNodeCell.TextContent);
        aNodeCell:=aNodeCell.NextSibling;
        aNodeCell:=aNodeCell.NextSibling;
        aCPrice:=Trim(aNodeCell.TextContent);

        Inc(aRow);
        StringGrid1.InsertRowWithValues(aRow, [aArticle, aName, aPV, aBV, aPPrice, aCPrice]);
        //aLogger.Info(aNodeRow.Attributes.GetNamedItem('table:style-name').NodeValue);
      end;
      Inc(i);
      aNodeRow:=aNodeRow.NextSibling;
    end;
    Result:=True;
  end;

begin
  aLogger:=TEventLog.Create(nil);
  aLogger.LogType:=ltFile;
  aLogger.Info('Log started');
  ReadXMLFile(aDoc, FileNameEdit1.FileName);
  try  
    aRow:=0;
    aContentNode:=aDoc.FindNode('office:document').FindNode('office:body').FindNode('office:text');
    aNode:=aContentNode.FindNode('table:table');
    if not Assigned(aNode) or not aNode.HasChildNodes then
      Exit;
    while not SameStr(aNode.NodeName, 'table:table') do
      aNode:=aNode.NextSibling;
    if aRow=0 then
      aNode:=aNode.NextSibling;// first table is the header of a visual table
    ParseFromOneTable(aNode);       
    aLogger.Info(aNode.TextContent);
    aNode:=aNode.NextSibling;
    aLogger.Info(aNode.NamespaceURI);
    aLogger.Info(aNode.TextContent);
    while not SameStr(aNode.NodeName, 'table:table') do
      aNode:=aNode.NextSibling;
    ParseFromOneTable(aNode);
  finally
    aLogger.Info('Log finished');
    aDoc.Free;
    aLogger.Free;
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin

end;

end.

