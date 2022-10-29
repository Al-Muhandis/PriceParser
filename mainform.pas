unit MainForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, EditBtn, StdCtrls, ExtCtrls, Grids, ComCtrls, SynEdit,
  Laz2_DOM
  ;

type

  { TFrmMain }

  TFrmMain = class(TForm)
    Btn: TButton;
    FlNmEdt: TFileNameEdit;
    StrngGrd: TStringGrid;
    TlBr: TToolBar;
    procedure BtnClick({%H-}Sender: TObject);
  private
    function ParseFromOneTable(var aTableNode: TDOMNode; var aRow: Integer): Boolean;
  public

  end;

var
  FrmMain: TFrmMain;

implementation

uses
  laz2_XMLRead, laz2_xmlutils, StrUtils, eventlog
  ;

{$R *.lfm}

{ TFrmMain }

procedure TFrmMain.BtnClick(Sender: TObject);
var
  aDoc: TXMLDocument;
  aContentNode, aNode: TDOMNode;
  aLogger: TEventLog;
  aRow: Integer;
begin
  aLogger:=TEventLog.Create(nil);
  aLogger.LogType:=ltFile;
  aLogger.Info('Log started');
  ReadXMLFile(aDoc, FlNmEdt.FileName);
  try  
    aRow:=0;
    aContentNode:=aDoc.FindNode('office:document').FindNode('office:body').FindNode('office:text');
    aNode:=aContentNode.FindNode('table:table');
    if not Assigned(aNode) or not aNode.HasChildNodes then
      Exit;
    ParseFromOneTable(aNode, aRow);
    ParseFromOneTable(aNode, aRow);  
    ParseFromOneTable(aNode, aRow);
  finally
    aLogger.Info('Log finished');
    aDoc.Free;
    aLogger.Free;
  end;
end;

function TFrmMain.ParseFromOneTable(var aTableNode: TDOMNode; var aRow: Integer): Boolean;
var
  aArticle, aName, aPV, aBV,aPPrice, aCPrice: String;
  aNodeCell, aNodeRow: TDOMNode;
begin
  Result:=False;
  aTableNode:=aTableNode.NextSibling;// first table is the header of a visual table
  while not SameStr(aTableNode.NodeName, 'table:table') do
    aTableNode:=aTableNode.NextSibling;
  if not aTableNode.HasChildNodes then
    Exit;
  aNodeRow:=aTableNode.FirstChild;
  while Assigned(aNodeRow) do
  begin
    while not aNodeRow.HasChildNodes do
      aNodeRow:=aNodeRow.NextSibling;
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
    StrngGrd.InsertRowWithValues(aRow, [aArticle, aName, aPV, aBV, aPPrice, aCPrice]);
    aNodeRow:=aNodeRow.NextSibling;
  end;
  Result:=True;
end;

end.

