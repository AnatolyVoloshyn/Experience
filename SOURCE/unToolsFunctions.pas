unit unToolsFunctions;

interface
uses Windows, Dialogs, Forms, Classes, SysUtils, cxGrid, Controls, Math, System.Variants, cxGridTableView, StrUtils, Winapi.Messages, cxGraphics,
     cxGridCustomTableView, Vcl.Graphics, cxGridDBBandedTableView, cxProgressBar, Vcl.StdCtrls;

{
procedure ShowSoob(AOwner: TComponent;MSG:String);
procedure DelSoob;
}
function ReadStringIniFiles(const Section:string; const Ident:string; const default_value:string=''):string;
procedure SaveStringIniFiles(const Section:string; const Ident:string; const Value:String);

function ReadIntegerIniFiles(const Section:string; const Ident:string; const default_value:integer=0):integer;
procedure SaveIntegerIniFiles(const Section:string; const Ident:string; const Value:Integer);

function ReadFloatIniFiles(const Section:string; const Ident:string; const default_value:Double=0):Double;
procedure SaveFloatIniFiles(const Section:string; const Ident:string; const Value:Double);

function ReadBooleanIniFiles(const Section:string; const Ident:string; const default_value:Boolean=False):Boolean;
procedure SaveBooleanIniFiles(const Section:string; const Ident:string; const Value:Boolean);

function ReadDateTimeIniFiles(const Section:string; const Ident:string; const default_value:TDateTime = 0):TDateTime;
procedure SaveDateTimeIniFiles(const Section:string; const Ident:string; const Value:TDateTime);

procedure EraseSection(const Section: string);
function existident(const Section:string;const Ident:string):boolean;

function ReadBinaryStreamIniFiles(const Section:string; const Ident:string; var s_value:TMemoryStream):Integer;
procedure SaveBinaryStreamIniFiles(const Section:string; const Ident:string; Value:TMemoryStream);


function RoundToMy(const Value: Double; Decimals: Integer): Double;


procedure ShowProgress(AOwner: TComponent;
                       sCaptionLablel:string;
                       nCountRecord:integer;
                       sCaptionForm:string='Пожалуйста, подождите...');
procedure UpdateProgressCaption(sUpdateCaptionLablel:string;sUpdateCaptionForm:string='Пожалуйста, подождите...');
procedure StepProgress;
procedure DelProgress;

////////////////////////////////////////////

procedure ShowProgress_(prog:TcxProgressBar;
                        lab: TLabel;
                        sCaptionLablel:string;
                        nCountRecord:integer
                        );



procedure UpdateProgressCaption_(lab: TLabel; sUpdateCaptionLablel:string);
procedure StepProgress_(prog:TcxProgressBar);
procedure DelProgress_(prog:TcxProgressBar; lab: TLabel);







{
procedure ShowOptionListReports(AOwner: TComponent);
procedure ShowSelectEmail(AOwner: TComponent;
                         var idOK:boolean; var output_memo: string);

procedure ShowViewMemo(AOwner: TComponent;
                       input_memo: variant);


function Get_Summa_String (column:integer):string;
function Get_Summa_String_report_5 (column:integer):string;
}
implementation

uses
//unSoob,
unIni,

Process_;

//unOptionListReports, unSelectEmails, unViewMemo;


{
procedure ShowSoob(AOwner: TComponent;MSG:String);
begin
 uShowSoob(AOwner,MSG);
end;



procedure DelSoob;
begin
  uDelSoob;
end;
}

function ReadStringIniFiles(const Section:string; const Ident:string; const default_value:string):string;
begin
 Result:=uReadStringIniFiles(Section, Ident,default_value);
end;

procedure SaveStringIniFiles(const Section:string; const Ident:string; const Value:String);
begin
 uSaveStringIniFiles(Section, Ident, Value);
end;

function ReadIntegerIniFiles(const Section:string; const Ident:string; const default_value:integer=0):integer;
begin
 Result:=uReadIntegerIniFiles(Section, Ident, default_value);
end;

procedure SaveIntegerIniFiles(const Section:string; const Ident:string; const Value:Integer);
begin
 uSaveIntegerIniFiles(Section, Ident, Value);
end;


function ReadFloatIniFiles(const Section:string; const Ident:string; const default_value:Double=0):Double;
begin
  Result:=uReadFloatIniFiles(Section, Ident, default_value);
end;


procedure SaveFloatIniFiles(const Section:string; const Ident:string; const Value:Double);
begin
  uSaveFloatIniFiles(Section, Ident, Value);
end;


function ReadBooleanIniFiles(const Section:string; const Ident:string; const default_value:Boolean=False):Boolean;
begin
 Result:=uReadBooleanIniFiles(Section, Ident, default_value);
end;


procedure SaveBooleanIniFiles(const Section:string; const Ident:string; const Value:Boolean);
begin
 uSaveBooleanIniFiles(Section, Ident, Value);
end;



function ReadDateTimeIniFiles(const Section:string; const Ident:string; const default_value:TDateTime = 0):TDateTime;
begin
   Result:=uReadDateTimeIniFiles(Section,Ident,default_value);
end;



procedure SaveDateTimeIniFiles(const Section:string; const Ident:string; const Value:TDateTime);
begin
  uSaveDateTimeIniFiles(Section,Ident,Value);
end;


function existident(const Section:string;const Ident:string):boolean;
begin
  Result:=uexistident(Section,Ident);
end;



procedure EraseSection(const Section: string);
begin
  uEraseSection(Section);
end;

function ReadBinaryStreamIniFiles(const Section:string; const Ident:string; var s_value:TMemoryStream):Integer;
begin
  Result:=uReadBinaryStreamIniFiles(Section,Ident,s_value);
end;

procedure SaveBinaryStreamIniFiles(const Section:string; const Ident:string; Value:TMemoryStream);
begin
  uSaveBinaryStreamIniFiles(Section, Ident, Value);
end;


function RoundToMy(const Value: Double; Decimals: Integer): Double;
begin
  Result:=RoundTo(Value + 0.0000001, Decimals);
end;




procedure ShowProgress(AOwner: TComponent;
                       sCaptionLablel:string;
                       nCountRecord:integer;
                       sCaptionForm:string='Пожалуйста, подождите...');
begin
 frmProcess := TfrmProcess.Create(AOwner);
// frmProcess.Position := poDesktopCenter;
 frmProcess.Position := poMainFormCenter;
 frmProcess.Show;
 frmProcess.lbl.Caption := sCaptionLablel;
 frmProcess.Caption := sCaptionForm;
 frmProcess.Update;
 //frmProcess.prb.PartsComplete:=0;
 frmProcess.prb.Position:=0;
 //frmProcess.prb.TotalParts:=nCountRecord;
 frmProcess.prb.Properties.Max:=nCountRecord;
 frmProcess.prb.Visible:=True;
end;


procedure UpdateProgressCaption(sUpdateCaptionLablel:string;sUpdateCaptionForm:string='Пожалуйста, подождите...');
begin
 frmProcess.lbl.Caption := sUpdateCaptionLablel;
 frmProcess.Caption := sUpdateCaptionForm;
 frmProcess.Update;
end;



procedure StepProgress;
begin
   //frmProcess.prb.IncPartsByOne;
   frmProcess.prb.Position:= frmProcess.prb.Position+1;
   frmProcess.prb.Update;
end;



procedure DelProgress;
begin
 frmProcess.prb.Visible:=False;
 frmProcess.Close;
 frmProcess.Destroy;
end;

///////////////////////////////////////
procedure ShowProgress_(prog:TcxProgressBar;
                        lab: TLabel;
                        sCaptionLablel:string;
                        nCountRecord:integer
                        );
begin
 prog.Position:=0;
 prog.Properties.Max:=nCountRecord;
 lab.Caption:=sCaptionLablel;
 lab.Visible:=True;
 prog.Visible:=True;
 Application.ProcessMessages;
end;

procedure UpdateProgressCaption_(lab: TLabel; sUpdateCaptionLablel:string);
begin
  lab.Caption:= sUpdateCaptionLablel;
  Application.ProcessMessages
end;

procedure StepProgress_(prog:TcxProgressBar);
begin
  prog.Position:= prog.Position+1;
  Application.ProcessMessages
end;

procedure DelProgress_(prog:TcxProgressBar; lab: TLabel);
begin
  prog.Visible:=False;
  lab.Visible:=False;
  Application.ProcessMessages
end;








 {
procedure ShowOptionListReports(AOwner: TComponent);
var
  opt:TfrmOptionListReports;
begin
  opt:=TfrmOptionListReports.Create(AOwner);
  opt.ShowModal;
  FreeAndNil(opt);
end;

procedure ShowSelectEmail(AOwner: TComponent;
                         var idOK:boolean; var output_memo: string);

var
  frm:TfrmSelectEmails;
begin
  frm:=TfrmSelectEmails.Create(AOwner);
  frm.ShowModal;
  idOK:=frm.pf_IDOK;

  frm.mdEmailsCh.DisableControls;
  frm.mdEmailsCh.First;
  output_memo:='';
  while not frm.mdEmailsCh.Eof do
  begin
    if frm.mdEmailsCh.FieldByName('is_ch').Value = 1 then
    output_memo:=output_memo+trim(frm.mdEmailsCh.FieldByName('fEmail').AsString)+'; ';

    frm.mdEmailsCh.Next;
  end;
  //  frm.edMemo.Lines.SaveToStream(output_memo);
  FreeAndNil(frm);
end;


procedure ShowViewMemo(AOwner: TComponent;
                       input_memo: variant);
var
  frm:TfrmViewMemo;
begin
  frm:=TfrmViewMemo.Create(AOwner, input_memo);
  frm.ShowModal;
  FreeAndNil(frm);
end;

function Get_Summa_String (column:integer):string;
var
temp_str:string;
n:integer;
begin
  n:=3;
  temp_str:='=';

   //Если есть что-то из FSS
   if ReadBooleanIniFiles('TfrmReportExcel_1', 'FILTER_FSS', True) = True then
   begin
     n:=n+1;
     temp_str:=temp_str+'R'+IntToStr(n)+'C'+IntToStr(column)+'+';
   end;

    //Если есть что-то из МСП
    if ReadBooleanIniFiles('TfrmReportExcel_1', 'FILTER_MSP', True) = True then
    begin
      n:=n+1;
      temp_str:=temp_str+'R'+IntToStr(n)+'C'+IntToStr(column)+'+';
      n:=n+3;
    end;


    if ReadBooleanIniFiles('TfrmReportExcel_1', 'FILTER_DCZ', True) = True then
    begin
      n:=n+1;
      temp_str:=temp_str+'R'+IntToStr(n)+'C'+IntToStr(column)+'+';
    end;

    if ReadBooleanIniFiles('TfrmReportExcel_1', 'FILTER_PFU', True) = True then
    begin
      n:=n+1;
      temp_str:=temp_str+'R'+IntToStr(n)+'C'+IntToStr(column)+'+';
    end;

    delete(temp_str, length(temp_str), 1);

    if (temp_str = '+') or (temp_str = '')  then temp_str:='';

    Result:= temp_str;
end;

function Get_Summa_String_report_5 (column:integer):string;
var
temp_str:string;
n:integer;
begin
  n:=3;
  temp_str:='=';

   //Если есть что-то из FSS
   if ReadBooleanIniFiles('TfrmReportExcel_1', 'FILTER_FSS_SUMMA_8_9', True) = True then
   begin
     n:=n+1;
     temp_str:=temp_str+'R'+IntToStr(n)+'C'+IntToStr(column)+'+';
   end;

    //Если есть что-то из МСП
    if ReadBooleanIniFiles('TfrmReportExcel_1', 'FILTER_MSP_SUMMA_8_9', True) = True then
    begin
      n:=n+1;
      temp_str:=temp_str+'R'+IntToStr(n)+'C'+IntToStr(column)+'+';
      n:=n+3;
    end;


    if ReadBooleanIniFiles('TfrmReportExcel_1', 'FILTER_DCZ_SUMMA_8_9', True) = True then
    begin
      n:=n+1;
      temp_str:=temp_str+'R'+IntToStr(n)+'C'+IntToStr(column)+'+';
    end;

    if ReadBooleanIniFiles('TfrmReportExcel_1', 'FILTER_PFU_SUMMA_8_9', True) = True then
    begin
      n:=n+1;
      temp_str:=temp_str+'R'+IntToStr(n)+'C'+IntToStr(column)+'+';
    end;

    delete(temp_str, length(temp_str), 1);

    if (temp_str = '+') or (temp_str = '')  then temp_str:='';

    Result:= temp_str;
end;
 }

end.
