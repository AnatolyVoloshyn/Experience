unit umMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls, cxClasses,
  cxShellBrowserDialog, cxGraphics, cxLookAndFeels, cxLookAndFeelPainters,
  Vcl.Menus, cxButtons, cxLocalization, System.ImageList, Vcl.ImgList,
  cxImageList, System.Actions, Vcl.ActnList, cxControls, cxContainer, cxEdit,
  cxLabel, Vcl.WinXPickers, dxGDIPlusClasses, cxImage, System.IOUtils, System.Types;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    cxShellBrowserDialog: TcxShellBrowserDialog;
    btnOpenFolder: TcxButton;
    cxImageList: TcxImageList;
    Localizer: TcxLocalizer;
    ActionList: TActionList;
    acAbout: TAction;
    acOpenFolder: TAction;
    lbFolder: TcxLabel;
    Label2: TLabel;
    DatePicker1: TDatePicker;
    acExecute: TAction;
    cxButton1: TcxButton;
    Label3: TLabel;
    cxImage1: TcxImage;
    Panel2: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure acOpenFolderExecute(Sender: TObject);
    procedure DatePicker1Change(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure acExecuteExecute(Sender: TObject);
  private
    { Private declarations }
    function GetFileCount(const APath: string; Mask: string): Integer;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses
 unExcelFunctions, Vcl.Themes, unToolsFunctions;

{$R *.dfm}
{$R RUS.RES}

function GetAppVersion(FileName: string): string;
var
  Dummy, InfoSize: DWORD;
  Info: Pointer;
  VerValue: PVSFixedFileInfo;
  VerSize: UINT;
begin
  Result := '';
  InfoSize := GetFileVersionInfoSize(PChar(FileName), Dummy);
  if InfoSize = 0 then Exit;

  GetMem(Info, InfoSize);
  try
    if GetFileVersionInfo(PChar(FileName), Dummy, InfoSize, Info) then
      if VerQueryValue(Info, '\', Pointer(VerValue), VerSize) then
        with VerValue^ do
          Result := Format('%d.%d.%d.%d',
            [HiWord(dwFileVersionMS), LoWord(dwFileVersionMS),
             HiWord(dwFileVersionLS), LoWord(dwFileVersionLS)]);
  finally
    FreeMem(Info);
  end;
end;

procedure TfrmMain.acOpenFolderExecute(Sender: TObject);
begin
  if cxShellBrowserDialog.Execute then
  begin
    lbFolder.Caption:=cxShellBrowserDialog.Path;
    SaveStringIniFiles('Experience','PATH_FILES', cxShellBrowserDialog.Path);

    if (DirectoryExists(lbFolder.Caption)) and (DatePicker1.Date<>0) then
      if acExecute.Enabled=False then acExecute.Enabled:=True;
  end;
end;

procedure TfrmMain.DatePicker1Change(Sender: TObject);
begin
 SaveDateTimeIniFiles('Experience', 'DATE', DatePicker1.Date);
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SaveDateTimeIniFiles('Experience', 'DATE', DatePicker1.Date);
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  rs: TResourceStream;
begin
  Caption:=Caption+' версия '+GetAppVersion(Application.ExeName);

  try
    rs := TResourceStream.Create(HInstance, 'RUSINI', RT_RCDATA);
    try
      Localizer.LoadFromStream(rs);
      Localizer.LanguageIndex := 1;
      Localizer.Translate;
    finally
      FreeAndNil(rs);
    end;
  except
    //
  end;

  lbFolder.Caption:=ReadStringIniFiles('Experience','PATH_FILES', '');
  DatePicker1.Date:= ReadDateTimeIniFiles('Experience', 'DATE', DATE);

  if (DirectoryExists(lbFolder.Caption)) and (DatePicker1.Date<>0) then acExecute.Enabled:=True;

end;

function TfrmMain.GetFileCount(const APath: string; Mask: string): Integer;
var
  Files: TStringDynArray;
begin
  // Получаем массив путей к файлам
  Files := TDirectory.GetFiles(APath, Mask);
  // Возвращаем размер массива
  Result := Length(Files);
end;

procedure TfrmMain.acExecuteExecute(Sender: TObject);
var
 FileName, Folder, Mask: string;
 Files: TStringDynArray;
begin
  Folder:=lbFolder.Caption;
  Mask:='*.xlsx';
  ShowProgress(Self, '', GetFileCount(Folder, Mask), 'Анализ указанной папки...');

  for FileName in TDirectory.GetFiles(Folder, Mask) do
  begin

    if Copy(ExtractFileName(FileName), 1, 2)<>'~$' then
    begin
       UpdateProgressCaption(ExtractFileName(FileName), 'Анализ указанной папки...');
       StepProgress;
       Sleep(50);

    end;
  end;
  DelProgress;
end;




end.
