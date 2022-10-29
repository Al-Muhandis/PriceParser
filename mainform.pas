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
    FRow: Integer;
    function ParseFromOneTable(var aTableNode: TDOMNode): Boolean;
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
begin
  FEventLog.Info('Log started');
  ReadXMLFile(aDoc, FlNmEdt.FileName);
  try  
    FRow:=0;
    aContentNode:=aDoc.FindNode('office:document').FindNode('office:body').FindNode('office:text');
    aNode:=aContentNode.FindNode('table:table');
    if not Assigned(aNode) or not aNode.HasChildNodes then
      Exit;
    while ParseFromOneTable(aNode) do
      aNode:=aNode.NextSibling;
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

function TFrmMain.ParseFromOneTable(var aTableNode: TDOMNode): Boolean;
var
  aArticle, aName: String;
  aPV: Integer;
  aBV, aPPrice, aCPrice: Currency;
  aNodeCell, aNodeRow, aChildTable: TDOMNode;
begin
  Result:=False;
  if not Assigned(aTableNode) then
    Exit;
  if FRow=35 then
    FEventLog.Debug('! '+aTableNode.NodeName+' '+aTableNode.Attributes.GetNamedItem('text:style-name').NodeName+'='+
      aTableNode.Attributes.GetNamedItem('text:style-name').NodeValue);
  while not SameStr(aTableNode.NodeName, 'table:table') do
  begin         // Check if the table can be found at lower level
    if SameStr(aTableNode.NodeName, 'text:section') then
    begin
      aChildTable:=aTableNode.FirstChild;
      while ParseFromOneTable(aChildTable) do
        aChildTable:=aChildTable.NextSibling;
    end;
    aTableNode:=aTableNode.NextSibling;
    if not Assigned(aTableNode) then
      break;
    if (FRow=35) then
      FEventLog.Debug('! '+aTableNode.NodeName);
  end;
  if not Assigned(aTableNode) or not aTableNode.HasChildNodes then
    Exit;
  aNodeRow:=aTableNode.FirstChild;
  while Assigned(aNodeRow) do
  begin
    while not aNodeRow.HasChildNodes do
      aNodeRow:=aNodeRow.NextSibling;
    aNodeCell:=aNodeRow.FirstChild;
    aArticle:=Trim(aNodeCell.TextContent);
    if aArticle.IsEmpty then
      Exit(True);
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

    Inc(FRow);
    StrngGrd.InsertRowWithValues(FRow, [FRow.ToString, aArticle, aName, aPV.ToString, CurrToStr(aBV),
      CurrToStr(aPPrice), CurrToStr(aCPrice)]);
    FEventLog.Info('Row: %d. Article: %s; Name: %s; PV: %d; BV: %s; PPrice: %s; CPrice: %s',
      [FRow, aArticle, aName, aPV, CurrToStr(aBV), CurrToStr(aPPrice), CurrToStr(aCPrice)]);
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
  if SameStr(S, '–') then
    Exit(0);
  Result:=StrToCurr(S, FFormatSettings);
end;

function TFrmMain.ParseInteger(const aValue: String): Integer;
var
  S: String;
begin
  S:=Trim(AValue);
  if SameStr(S, '–') then
    Exit(0);
  Result:=StrToInt(S);
end;

function TFrmMain.ParseText(const aValue: String): String;
var
  S: String;
begin
  S:=AdjustLineBreaks(aValue); 
  S:=ReplaceStr(Trim(S), LineEnding, EmptyStr);
  Result:=Trim(DelSpace1(S));
  if Result.IsEmpty then
    raise Exception.Create(Format('Value is empty! Row %d', [FRow+1]));
end;

end.

