unit MainForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, EditBtn, StdCtrls, ExtCtrls, Grids, ComCtrls, SynEdit,
  Laz2_DOM, eventlog
  ;

type

  { TFrmMain }

  TFrmMain = class(TForm)
    Btn: TButton;
    FlNmEdt: TFileNameEdit;
    StrngGrd: TStringGrid;
    TlBr: TToolBar;
    procedure BtnClick({%H-}Sender: TObject);
    procedure FormCreate({%H-}Sender: TObject);
  private
    FEventLog: TEventLog;
    FFormatSettings: TFormatSettings;
    function ParseFromOneTable(var aTableNode: TDOMNode; var aRow: Integer): Boolean;
    function ParseCurrency(const aValue: String): Currency;
    function ParseInteger(const aValue: String): Integer;
    function ParseText(const aValue: String): String;
  public

  end;

var
  FrmMain: TFrmMain;

implementation

uses
  laz2_XMLRead, laz2_xmlutils, StrUtils
  ;

{$R *.lfm}

{ TFrmMain }

procedure TFrmMain.BtnClick(Sender: TObject);
var
  aDoc: TXMLDocument;
  aContentNode, aNode: TDOMNode;
  aRow: Integer;
begin
  FEventLog.Info('Log started');
  ReadXMLFile(aDoc, FlNmEdt.FileName);
  try  
    aRow:=0;
    aContentNode:=aDoc.FindNode('office:document').FindNode('office:body').FindNode('office:text');
    aNode:=aContentNode.FindNode('table:table');
    if not Assigned(aNode) or not aNode.HasChildNodes then
      Exit;
    repeat
    until not ParseFromOneTable(aNode, aRow);
  finally
    FEventLog.Info('Log finished');
    aDoc.Free;
  end;
end;

procedure TFrmMain.FormCreate(Sender: TObject);
begin
  FFormatSettings:=DefaultFormatSettings;
  FFormatSettings.DecimalSeparator:=','; 
  FFormatSettings.ThousandSeparator:='.';
  FEventLog:=TEventLog.Create(Self);
  FEventLog.LogType:=ltFile;
end;

function TFrmMain.ParseFromOneTable(var aTableNode: TDOMNode; var aRow: Integer): Boolean;
var
  aArticle, aName: String;
  aPV: Integer;
  aBV, aPPrice, aCPrice: Currency;
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
    aName:=ParseText(aNodeCell.TextContent);
    aNodeCell:=aNodeCell.NextSibling;
    aPV:=ParseInteger(aNodeCell.TextContent);
    aNodeCell:=aNodeCell.NextSibling;
    aBV:=ParseCurrency(aNodeCell.TextContent);
    aNodeCell:=aNodeCell.NextSibling;
    aNodeCell:=aNodeCell.NextSibling;
    aPPrice:=ParseCurrency(aNodeCell.TextContent);
    aNodeCell:=aNodeCell.NextSibling;
    aNodeCell:=aNodeCell.NextSibling;
    aCPrice:=ParseCurrency(aNodeCell.TextContent);

    Inc(aRow);
    StrngGrd.InsertRowWithValues(aRow, [aArticle, aName, aPV.ToString, CurrToStr(aBV),
      CurrToStr(aPPrice), CurrToStr(aCPrice)]);
    FEventLog.Info('Article: %s; Name: %s; PV: %d; BV: %s; PPrice: %s; CPrice: %s',
      [aArticle, aName, aPV, CurrToStr(aBV), CurrToStr(aPPrice), CurrToStr(aCPrice)]);
    aNodeRow:=aNodeRow.NextSibling;
  end;
  Result:=True;
end;

function TFrmMain.ParseCurrency(const aValue: String): Currency;
var
  S: String;
begin
  S:=ParseText(aValue);
  S:=ReplaceStr(S, '.', EmptyStr);  
  S:=ReplaceStr(S, ' ', EmptyStr);
  Result:=StrToCurr(S, FFormatSettings);
end;

function TFrmMain.ParseInteger(const aValue: String): Integer;
begin
  Result:=StrToInt(Trim(AValue));
end;

function TFrmMain.ParseText(const aValue: String): String;
var
  S: String;
begin
  S:=AdjustLineBreaks(aValue); 
  S:=ReplaceStr(Trim(S), LineEnding, EmptyStr);
  Result:=Trim(DelSpace1(S));
end;

end.

