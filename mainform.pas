unit MainForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, EditBtn, StdCtrls, ExtCtrls, Grids, ComCtrls, SynEdit,
  Laz2_DOM, eventlog, fpjson
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
    FArticles: TStringList;
    FEventLog: TEventLog;
    FFormatSettings: TFormatSettings;
    FRow: Integer;
    FOutJSON: TJSONArray;
    procedure AppendPriceItem(aJSONArray: TJSONArray; const aArticle, aName: String; aPV: Integer; aPPrice, aCPrice,
      aBV: Currency);
    procedure ParsePrice(const aPriceFileName: String);
    function ParseFromOneTable(var aTableNode: TDOMNode): Boolean;
    function ParseCurrency(const aValue: String; out aParsed: Boolean): Currency;
    function ParseInteger(const aValue: String; out aParsed: Boolean): Integer;
    function ParseText(const aValue: String): String;
  public

  end;

var
  FrmMain: TFrmMain;

const
  dt_article='article';
  dt_name='name';
  dt_pv='pv';
  dt_bv='bv';
  dt_cprice='cprice';
  dt_pprice='pprice';

implementation

uses
  laz2_XMLRead, laz2_xmlutils, StrUtils
  ;

{$R *.lfm}

{ TFrmMain }

procedure TFrmMain.BtnClick(Sender: TObject);
begin
  ParsePrice(FlNmEdt.FileName);
end;

procedure TFrmMain.FormCreate(Sender: TObject);
begin
  FFormatSettings:=DefaultFormatSettings;
  FFormatSettings.DecimalSeparator:=','; 
  FFormatSettings.ThousandSeparator:='.';
  FEventLog:=TEventLog.Create(Self);
  FEventLog.LogType:=ltFile;
end;

procedure TFrmMain.AppendPriceItem(aJSONArray: TJSONArray; const aArticle, aName: String; aPV: Integer;
  aPPrice, aCPrice, aBV: Currency);
var
  aJSONItem: TJSONObject;
begin
  if FArticles.IndexOf(aArticle)>-1 then
  begin
    FEventLog.Warning(Format('Duplicate! [%s] %s', [aArticle, aName]));
    Exit;
  end;
  Inc(FRow);
  StrngGrd.InsertRowWithValues(FRow, [FRow.ToString, aArticle, aName, aPV.ToString, CurrToStr(aBV),
            CurrToStr(aPPrice), CurrToStr(aCPrice)]);
  aJSONItem:=TJSONObject.Create;
  aJSONItem.Add(dt_article, aArticle); 
  aJSONItem.Add(dt_name, aName);
  aJSONItem.Add(dt_pv, aPV);
  aJSONItem.Add(dt_pprice, aPPrice);
  aJSONItem.Add(dt_cprice, aCPrice); 
  aJSONItem.Add(dt_bv, aBV);
  aJSONArray.Add(aJSONItem);
  FArticles.Add(aArticle);
end;

procedure TFrmMain.ParsePrice(const aPriceFileName: String);
var
  aDoc: TXMLDocument;
  aContentNode, aNode: TDOMNode;
  aFile: TStringList;
begin
  FEventLog.Info('Log started');
  FOutJSON:=TJSONArray.Create;
  ReadXMLFile(aDoc, aPriceFileName);
  FArticles:=TStringList.Create;
  FArticles.CaseSensitive:=False;
  FArticles.Sorted:=True;
  FArticles.Duplicates:=dupError;
  try
    FRow:=0;
    aContentNode:=aDoc.FindNode('office:document').FindNode('office:body').FindNode('office:text');
    aNode:=aContentNode.FindNode('table:table');
    if not Assigned(aNode) or not aNode.HasChildNodes then
      Exit;
    while ParseFromOneTable(aNode) do
      aNode:=aNode.NextSibling;
    aFile:=TStringList.Create;
    aFile.Text:=FOutJSON.FormatJSON();
    aFile.SaveToFile('pricelist.json');
  finally
    aFile.Free;
    FArticles.Free;
    FEventLog.Info('Log finished');
    FOutJSON.Free;
    aDoc.Free;
  end;
end;

function TFrmMain.ParseFromOneTable(var aTableNode: TDOMNode): Boolean;
var
  aArticle, aName: String;
  aPV: Integer;
  aBV, aPPrice, aCPrice: Currency;
  aNodeCell, aNodeRow, aChildTable: TDOMNode;
  aParsed: Boolean;
begin
  Result:=False;
  if not Assigned(aTableNode) then
    Exit;
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
  end;
  if not Assigned(aTableNode) or not aTableNode.HasChildNodes then
    Exit;
  aNodeRow:=aTableNode.FirstChild;
  while Assigned(aNodeRow) do
  begin
    while not aNodeRow.HasChildNodes do
      aNodeRow:=aNodeRow.NextSibling;
    aNodeCell:=aNodeRow.FirstChild;
    aArticle:=ParseText(aNodeCell.TextContent);
    if not aArticle.IsEmpty then
    begin
      aNodeCell:=aNodeCell.NextSibling;
      aName:=ParseText(aNodeCell.TextContent);
      if aName.IsEmpty then
        Exit(True);
      aNodeCell:=aNodeCell.NextSibling;
      aPV:=ParseInteger(aNodeCell.TextContent, aParsed);
      aNodeCell:=aNodeCell.NextSibling;
      aBV:=ParseCurrency(aNodeCell.TextContent, aParsed);
      if aParsed then
      begin
        aNodeCell:=aNodeCell.NextSibling;
        aNodeCell:=aNodeCell.NextSibling;
        if not Assigned(aNodeCell) then
          Exit(True);
        aPPrice:=ParseCurrency(aNodeCell.TextContent, aParsed);
        if not aParsed then
          raise Exception.Create('Parse error');
        aNodeCell:=aNodeCell.NextSibling;
        aNodeCell:=aNodeCell.NextSibling;
        aCPrice:=ParseCurrency(aNodeCell.TextContent, aParsed);
        if not aParsed then
          raise Exception.Create('Parse error');
        AppendPriceItem(FOutJSON, aArticle, aName, aPV, aPPrice, aCPrice, aBV);
      end;
    end;
    aNodeRow:=aNodeRow.NextSibling;
  end;
  Result:=True;
end;

function TFrmMain.ParseCurrency(const aValue: String; out aParsed: Boolean): Currency;
var
  S: String;
begin
  Result:=0;
  aParsed:=True;
  S:=ParseText(aValue);
  if S.IsEmpty then
    Exit;
  S:=ReplaceStr(S, '.', EmptyStr);  
  S:=ReplaceStr(S, ' ', EmptyStr);
  if SameStr(S, '–') then
    Exit;
  aParsed:=TryStrToCurr(S, Result, FFormatSettings);
end;

function TFrmMain.ParseInteger(const aValue: String; out aParsed: Boolean): Integer;
var
  S: String;
begin
  Result:=0;
  S:=Trim(AValue);
  if SameStr(S, '–') then
    Exit;
  aParsed:=TryStrToInt(S, Result);
end;

function TFrmMain.ParseText(const aValue: String): String;
var
  S: String;
begin
  Result:=EmptyStr;
  S:=AdjustLineBreaks(aValue); 
  S:=ReplaceStr(Trim(S), LineEnding, EmptyStr);
  Result:=Trim(DelSpace1(S));
end;

end.

