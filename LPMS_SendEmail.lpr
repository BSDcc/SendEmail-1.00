program LPMS_SendEmail;

{$MODE Delphi}

uses
  Forms, Interfaces,
  ldSendEmail in 'ldSendEmail.pas' {FldSendEmail};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFldSendEmail, FldSendEmail);
  Application.Run;
end.
