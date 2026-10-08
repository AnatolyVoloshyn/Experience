unit umMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls, cxClasses,
  cxShellBrowserDialog, cxGraphics, cxLookAndFeels, cxLookAndFeelPainters,
  Vcl.Menus, cxButtons, cxLocalization, System.ImageList, Vcl.ImgList,
  cxImageList, System.Actions, Vcl.ActnList, cxControls, cxContainer, cxEdit,
  cxLabel, Vcl.WinXPickers, dxGDIPlusClasses, cxImage, System.IOUtils, System.Types, System.Masks, System.StrUtils,
  dxCore, dxCoreClasses, dxHashUtils, dxSpreadSheetCore,
  dxSpreadSheetCoreFormulas, dxSpreadSheetCoreHistory, dxSpreadSheetCoreStyles,
  dxSpreadSheetCoreStrs, dxSpreadSheetConditionalFormatting,
  dxSpreadSheetConditionalFormattingRules, dxSpreadSheetClasses,
  dxSpreadSheetContainers, dxSpreadSheetFormulas, dxSpreadSheetHyperlinks,
  dxSpreadSheetFunctions, dxSpreadSheetStyles, dxSpreadSheetGraphics,
  dxSpreadSheetPrinting, dxSpreadSheetTypes, dxSpreadSheetUtils,
  dxSpreadSheetFormattedTextUtils, dxBarBuiltInMenu, dxSpreadSheet,
  dxShellDialogs, Data.DB, dxmdaset, Vcl.Grids, Vcl.DBGrids;

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
    acSaveExcelFile: TAction;
    Panel3: TPanel;
    Panel8: TPanel;
    btnExportFile: TcxButton;
    cxButton3: TcxButton;
    SSR: TdxSpreadSheet;
    dxSaveFileDialog: TdxSaveFileDialog;
    MD: TdxMemData;
    MDFiirst: TDateField;
    MDSecond: TDateField;
    MDR: TIntegerField;
    MDC: TIntegerField;
    SST: TdxSpreadSheet;
    MDP: TdxMemData;
    DateField1: TDateField;
    DateField2: TDateField;
    IntegerField1: TIntegerField;
    IntegerField2: TIntegerField;
    procedure FormCreate(Sender: TObject);
    procedure acOpenFolderExecute(Sender: TObject);
    procedure DatePicker1Change(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure acExecuteExecute(Sender: TObject);
    procedure acSaveExcelFileExecute(Sender: TObject);
    procedure acAboutExecute(Sender: TObject);
  private
    { Private declarations }
    const Mask ='вислуга*.xlsx';
    //const Mask ='вислуга ВОЛОШИН.xlsx';
    //const Mask ='вислуга БАСЮК.xlsx';
    //var Folder:string;
    var DateRasc:TDateTime;
    function GetFileCount(const APath:string; Mask: string): Integer;
    function CopyFile_ (sFileName:string):string;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;
  Folder:string;

implementation

uses
 unExcelFunctions, Vcl.Themes, unToolsFunctions, unAbout;

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
 DateRasc:= DatePicker1.Date;
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SaveStringIniFiles('Experience','PATH_FILES', lbFolder.Caption);
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

  lbFolder.Caption:=ReadStringIniFiles('Experience','PATH_FILES', '\\A4785\Disc\ВИСЛУГА РОКІВ А4785\Вислуга років');
  DatePicker1.Date:= ReadDateTimeIniFiles('Experience', 'DATE', DATE);
  DateRasc:= DatePicker1.Date;

  if (DirectoryExists(lbFolder.Caption)) and (DatePicker1.Date<>0) then acExecute.Enabled:=True;
end;

function TfrmMain.GetFileCount(const APath:string; Mask: string): Integer;
var
  Files: TStringDynArray;
begin
  // Получаем массив путей к файлам
  Files := TDirectory.GetFiles(APath,
           function(const Path: string; const SearchRec: TSearchRec): Boolean
           begin
             Result := MatchesMask(WideUpperCase(SearchRec.Name), WideUpperCase(Mask));
           end);

  // Возвращаем размер массива
  Result := Length(Files);
 end;

procedure TfrmMain.acExecuteExecute(Sender: TObject);
var
 FileName, Folder, FileName_temp: string;
 col,row,i,current_Worksheet, current_row:integer;
 temp_datetime:TDateTime;
begin
  Folder:=lbFolder.Caption;
  if (DirectoryExists(Folder)) and (DatePicker1.Date<>0) and (GetFileCount(Folder, Mask)>0) then
  begin

//     //Рисуем заголовок в основном excele
//     SSR.ClearAll;
//     SSR.OptionsView.R1C1Reference:=True;
//     current_Worksheet:=AddWorksheet(SSR,'Результат анализа файлов выслуг');
//     SSR.Sheets[current_Worksheet].BeginUpdate;

     //Загружаем шаблон
     SSR.LoadFromFile(ExtractFileDir(Application.ExeName)+'\Вислуга-Вся.xlsx');
     SSR.OptionsView.R1C1Reference:=True;
     current_Worksheet:=SSR.Sheets[0].Index;
     ClearCells_(SSR,current_Worksheet, 1, 54, 4, 1000);
     SetFocusedCell_(SSR,current_Worksheet,1,4);
     SSR.Sheets[current_Worksheet].BeginUpdate;

     current_row:=1;
     DrawCell(SSR,
             'Список вислуг о/с на '+FormatDateTime('dd.mm.yyyy',DateRasc), current_Worksheet,current_row,1,True,ssahLeft,ssavCenter,$00,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

     DrawCell(SSR,
             'Сформирован: '+FormatDateTime('dd.mm.yyyy HH:NN:SS',now), current_Worksheet,current_row,5,True,ssahLeft,ssavCenter,$00,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);
     //Удаляем начинку



     //Уже есть в шаблоне
     {
     current_row:=2;
     SetColumnWidth(SSR, current_Worksheet, 1, 30);
     DrawCellMergeCells(SSR,current_Worksheet, 1, current_row, 1, current_row+1,
                        0, sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
                        True, '№ п/п',
                        current_row, 1, True, ssahCenter, ssavCenter, $31, True, False, 'Calibri', 11, [], 0,
                        $00ECFFD8 ,0, sscfsSolid, True);

     SetColumnWidth(SSR, current_Worksheet, 2, 250);
     DrawCellMergeCells(SSR,current_Worksheet, 2, current_row, 2, current_row+1,
                        0, sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
                        True, 'Назва файлу',
                        current_row, 2, True, ssahCenter, ssavCenter, $31, True, False, 'Calibri', 11, [], 0,
                        $00ECFFD8 ,0, sscfsSolid, True);

     SetColumnWidth(SSR, current_Worksheet, 3, 300);
     DrawCellMergeCells(SSR,current_Worksheet, 3, current_row, 3, current_row+1,
                        0, sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
                        True, 'ПІБ у файлі',
                        current_row, 3, True, ssahCenter, ssavCenter, $31, True, False, 'Calibri', 11, [], 0,
                        $00ECFFD8 ,0, sscfsSolid, True);


     SetColumnWidth(SSR, current_Worksheet, 4, 50);
     SetColumnWidth(SSR, current_Worksheet, 5, 50);
     SetColumnWidth(SSR, current_Worksheet, 6, 50);
     DrawCellMergeCells(SSR,current_Worksheet, 4, current_row, 6, current_row,
                        0, sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
                        True, 'Календарна вислуга',
                        current_row, 4, True, ssahCenter, ssavCenter, $31, True, False, 'Calibri', 11, [], 0,
                        $0097FCFC ,0, sscfsSolid, True);
     DrawCell(SSR,
             'Роки', current_Worksheet,current_row+1,4,True,ssahCenter,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              $0097FCFC,0,sscfsSolid);
     DrawCell(SSR,
             'Міс.', current_Worksheet,current_row+1,5,True,ssahCenter,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              $0097FCFC,0,sscfsSolid);
     DrawCell(SSR,
             'Дні', current_Worksheet,current_row+1,6,True,ssahCenter,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              $0097FCFC,0,sscfsSolid);

     SetColumnWidth(SSR, current_Worksheet, 7, 50);
     SetColumnWidth(SSR, current_Worksheet, 8, 50);
     SetColumnWidth(SSR, current_Worksheet, 9, 50);
     DrawCellMergeCells(SSR,current_Worksheet, 7, current_row, 9, current_row,
                        0, sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
                        True, 'Пільгова вислуга',
                        current_row, 7, True, ssahCenter, ssavCenter, $31, True, False, 'Calibri', 11, [], 0,
                        $00D9D9D9 ,0, sscfsSolid, True);
     DrawCell(SSR,
             'Роки', current_Worksheet,current_row+1,7,True,ssahCenter,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              $00D9D9D9,0,sscfsSolid);
     DrawCell(SSR,
             'Міс.', current_Worksheet,current_row+1,8,True,ssahCenter,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              $00D9D9D9,0,sscfsSolid);
     DrawCell(SSR,
             'Дні', current_Worksheet,current_row+1,9,True,ssahCenter,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              $00D9D9D9,0,sscfsSolid);

     SetColumnWidth(SSR, current_Worksheet, 10, 50);
     SetColumnWidth(SSR, current_Worksheet, 11, 50);
     SetColumnWidth(SSR, current_Worksheet, 12, 50);
     DrawCellMergeCells(SSR,current_Worksheet, 10, current_row, 12, current_row,
                        0, sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
                        True, 'Загальна вислуга',
                        current_row, 10, True, ssahCenter, ssavCenter, $31, True, False, 'Calibri', 11, [], 0,
                        $0000C0FF ,0, sscfsSolid, True);
     DrawCell(SSR,
             'Роки', current_Worksheet,current_row+1,10,True,ssahCenter,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              $0000C0FF,0,sscfsSolid);
     DrawCell(SSR,
             'Міс.', current_Worksheet,current_row+1,11,True,ssahCenter,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              $0000C0FF,0,sscfsSolid);
     DrawCell(SSR,
             'Дні', current_Worksheet,current_row+1,12,True,ssahCenter,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              $0000C0FF,0,sscfsSolid);
     }
     //Конец Уже есть в шаблоне



     //Пишем начинку
     current_row:=3;
     ShowProgress(Self, '', GetFileCount(Folder, Mask), 'Анализ указанной папки...');
     for FileName in TDirectory.GetFiles(Folder,
                     function(const Path: string; const SearchRec: TSearchRec): Boolean
                     begin
                       Result := MatchesMask(WideUpperCase(SearchRec.Name), WideUpperCase(Mask));
                     end)
    do
    begin
         try
           UpdateProgressCaption(ExtractFileName(FileName), 'Анализ указанной папки...');
           StepProgress;
           Application.ProcessMessages;
           FileName_temp:=ExtractFileDir(Application.ExeName)+'\'+ CopyFile_(FileName);

           current_row:=current_row+1;
           //Работаем с временным файлом
           SST.ClearAll;
           SST.BeginUpdate;
           SST.LoadFromFile(FileName_temp);
           SST.OptionsView.R1C1Reference:=True;

           ///////////////////////////////
           MD.DisableControls;
           if MD.Active then MD.Close;
           MD.Open;
           col:=2;
           i:=0;
           for row:= 4 to 119 do
           begin
             if TryStrToDate(Get_Cell(SST, 0, row, col).AsString, temp_datetime) then
             begin
               i:=i+1;
               if Odd(i) then
               begin
                MD.Append;
                MD.FieldByName('Fiirst').AsDateTime:= temp_datetime;
                MD.Post;
               end else
               begin
                MD.Edit;
                MD.FieldByName('Second').AsDateTime:= temp_datetime;
                MD.FieldByName('R').Value:= row;
                MD.FieldByName('C').Value:= col;
                MD.Post;
               end;
             end;
           end;
           MD.Last;
           DrawCellOnlyValue(SST,0,MD.FieldByName('R').AsInteger, MD.FieldByName('C').AsInteger, DateToStr(DateRasc), True);
           RecalcFormulsWorksheet(SST);
           Application.ProcessMessages;
           /////////////////////////////////////////////////////////

          DrawCell(SSR,
              IntToStr(current_row-3), current_Worksheet,current_row,1,True,ssahRight,ssavCenter,$01,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          DrawCell(SSR,
              ExtractFileName(FileName), current_Worksheet,current_row,2,True,ssahLeft,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          DrawCell(SSR,
              Get_Cell(SST, 0, 1, 2).AsString, current_Worksheet,current_row,3,True,ssahLeft,ssavCenter,$31,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          ///////
          DrawCell(SSR,
              Get_Cell(SST, 1, 89, 10).AsString, current_Worksheet,current_row,4,True,ssahRight,ssavCenter,$01,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          DrawCell(SSR,
              Get_Cell(SST, 1, 89, 11).AsString, current_Worksheet,current_row,5,True,ssahRight,ssavCenter,$01,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          DrawCell(SSR,
              Get_Cell(SST, 1, 89, 12).AsString, current_Worksheet,current_row,6,True,ssahRight,ssavCenter,$01,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          ///////
          DrawCell(SSR,
              Get_Cell(SST, 1, 91, 10).AsString, current_Worksheet,current_row,7,True,ssahRight,ssavCenter,$01,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          DrawCell(SSR,
              Get_Cell(SST, 1, 91, 11).AsString, current_Worksheet,current_row,8,True,ssahRight,ssavCenter,$01,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          DrawCell(SSR,
              Get_Cell(SST, 1, 91, 12).AsString, current_Worksheet,current_row,9,True,ssahRight,ssavCenter,$01,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          ///////
          DrawCell(SSR,
              Get_Cell(SST, 1, 93, 10).AsString, current_Worksheet,current_row,10,True,ssahRight,ssavCenter,$01,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          DrawCell(SSR,
              Get_Cell(SST, 1, 93, 11).AsString, current_Worksheet,current_row,11,True,ssahRight,ssavCenter,$01,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);

          DrawCell(SSR,
              Get_Cell(SST, 1, 93, 12).AsString, current_Worksheet,current_row,12,True,ssahRight,ssavCenter,$01,False,False,'Calibri', 11, [], 0,
              0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
              0,0,sscfsSolid);
           SST.EndUpdate;


           ///////////////////////////////
           MDP.DisableControls;
           if MDP.Active then MDP.Close;
           MDP.Open;
           col:=15;
           i:=0;
           for row:= 4 to 119 do
           begin
             if TryStrToDate(Get_Cell(SST, 0, row, col).AsString, temp_datetime) then
             begin
               i:=i+1;
               if Odd(i) then
               begin
                MDP.Append;
                MDP.FieldByName('Fiirst').AsDateTime:= temp_datetime;
                MDP.Post;
               end else
               begin
                MDP.Edit;
                MDP.FieldByName('Second').AsDateTime:= temp_datetime;
                MDP.FieldByName('R').Value:= row;
                MDP.FieldByName('C').Value:= col;
                MDP.Post;
               end;
             end;
           end;
           MDP.First;

           col:=14;
           while not MDP.Eof do
           begin
             col:=col+1;
             DrawCell(SSR,
             MDP.FieldByName('Fiirst').Value, current_Worksheet,current_row,col,True,ssahRight,ssavCenter,$0E,False,False,'Calibri', 11, [], 0,
             0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
             0,0,sscfsSolid);
             SetColumnWidth(SSR, current_Worksheet, col, 75);

             col:=col+1;
             DrawCell(SSR,
             MDP.FieldByName('Second').Value, current_Worksheet,current_row,col,True,ssahRight,ssavCenter,$0E,False,False,'Calibri', 11, [], 0,
             0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,0,sscbsDefault,
             0,0,sscfsSolid);
             SetColumnWidth(SSR, current_Worksheet, col, 75);
             MDP.Next;
           end;
           /////////////////////////////////////////////////////////


           Application.ProcessMessages;
         finally
           if FileExists(FileName_temp) then
           DeleteFile(FileName_temp);
         end;
    end;
    DelProgress;

         // Конец Пишем начинку

    FreezeRows(SSR, current_Worksheet, 3);
    SetFocusedCell_(SSR,current_Worksheet,1,4);
    SSR.Sheets[current_Worksheet].EndUpdate;
    RecalcFormulsWorksheet(SSR);
  end;
end;

function TfrmMain.CopyFile_(sFileName:string):string;
begin
  if FileExists(sFileName) then
  begin
   CopyFile(PWideChar(sFileName), PWideChar(ChangeFileExt(ExtractFileName(sFileName),'')+'_временный.xlsx'), False);
   Result:=ChangeFileExt(ExtractFileName(sFileName),'')+'_временный.xlsx';
  end;
end;

procedure TfrmMain.acSaveExcelFileExecute(Sender: TObject);
begin
 if dxSaveFileDialog.Execute then
 begin
   if ExtractFileExt(dxSaveFileDialog.FileName)='' then
   SSR.SaveToFile(dxSaveFileDialog.FileName+'.xlsx') else
   SSR.SaveToFile(dxSaveFileDialog.FileName);
 end;

end;

procedure TfrmMain.acAboutExecute(Sender: TObject);
begin
  frmAbout := TfrmAbout.Create(frmMain);
  try
    frmAbout.ShowModal;
  finally
    frmAbout.Free;
  end;
end;


end.
