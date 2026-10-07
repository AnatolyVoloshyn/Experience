unit Process_;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls, cxGraphics, cxControls,
  cxLookAndFeels, cxLookAndFeelPainters, cxContainer, cxEdit, cxProgressBar;

type
  TfrmProcess = class(TForm)
    lbl: TLabel;
    prb_: TProgressBar;
    prb: TcxProgressBar;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProcess: TfrmProcess;

implementation

{$R *.dfm}


end.
