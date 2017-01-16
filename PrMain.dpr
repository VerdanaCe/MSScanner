program PrMain;

uses
  Forms,
  Unit_Main in 'Unit_Main.pas' {ScanPolis},
  Unit_Scan in 'Unit_Scan.pas',
  commonClasses in 'commonClasses.pas',
  aClass_TIME in 'aClass_TIME.pas',
  CharCurrency in 'CharCurrency.pas',
  MTMainForm in 'MTMainForm.pas' {MainForm};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TScanPolis, ScanPolis);
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
