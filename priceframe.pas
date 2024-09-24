unit priceframe;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, ComCtrls, EditBtn, ExtCtrls, fpjson
  ;

type

  { TFramePrice }

  TFramePrice = class(TFrame)
    CntrlBr: TControlBar;
    FlNmEdt: TFileNameEdit;
    LstVw: TListView;
    SttsBr: TStatusBar;
    procedure FlNmEdtChange(Sender: TObject);
  private
    function LoadFromJSON(const aFileName: String): TJSONData;
  public

  end;

implementation

{$R *.lfm}

uses
  jsonscanner, jsonparser
  ;

{ TFramePrice }

procedure TFramePrice.FlNmEdtChange(Sender: TObject);
var
  aArray: TJSONArray;
  aItem: TJSONEnum;
  aObject: TJSONObject;
begin
  aArray:=LoadFromJSON(TFileNameEdit(Sender).FileName) as TJSONArray;
  try
    for aItem in aArray do
    begin
      aObject:=aItem.Value as TJSONObject;
      with LstVw.Items.Add do
      begin
        Caption:=aObject.Strings['article'];
        SubItems.Add(aObject.Strings['name']);
        SubItems.Add(aObject.Integers['pv'].ToString);
        SubItems.Add(aObject.Floats['pprice'].ToString);
        SubItems.Add(aObject.Floats['cprice'].ToString);
        SubItems.Add(aObject.Floats['bv'].ToString);
      end;
    end;
  finally
    SttsBr.SimpleText:=Format('Count: %d', [aArray.Count]);
    aArray.Free;
  end;
end;

function TFramePrice.LoadFromJSON(const aFileName: String): TJSONData;
var
  S: TFileStream;
  P: TJSONParser;
begin
  Result:=nil;
  S:=TFileStream.Create(aFileName,fmOpenRead);
  try
    P:=TJSONParser.Create(S, DefaultOptions);
    try
      Result:=P.Parse;
    finally
      P.Free;
    end;
  finally
    S.Free;
  end;
end;

end.

