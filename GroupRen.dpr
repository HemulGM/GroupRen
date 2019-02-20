program GroupRen;

uses
  Forms,
  Main in 'Main.pas' {FormMain},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  TStyleManager.TrySetStyle('Aqua Light Slate');
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
