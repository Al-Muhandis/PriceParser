unit priceexplorerform;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ComCtrls, PairSplitter, priceframe;

type

  { TForm1 }

  TForm1 = class(TForm)
    FramePrice1: TFramePrice;
    FrmPrcLeft: TFramePrice;
    PrSpltr: TPairSplitter;
    PairSplitterSide1: TPairSplitterSide;
    PairSplitterSide2: TPairSplitterSide;
  private

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

end.

