program Experience;

uses
  Vcl.Forms,
  umMain in '..\SOURCE\umMain.pas' {frmMain},
  Process_ in '..\SOURCE\Process_.pas' {frmProcess},
  unToolsFunctions in '..\SOURCE\unToolsFunctions.pas',
  unIni in '..\SOURCE\unIni.pas',
  unExcelFunctions in '..\SOURCE\unExcelFunctions.pas',
  Vcl.Themes,
  Vcl.Styles,
  unAbout in '..\SOURCE\unAbout.pas' {frmAbout};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Smokey Quartz Kamri');
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
