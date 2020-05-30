program BSD_SendEmail;

{$MODE Delphi}

uses
  Forms, Interfaces,
  bsdsendemail in 'BSDSendEmail.pas' {FBSDSendEmail};

{$R *.res}

begin
  Application.Initialize;
   Application. CreateForm( TFBSDSendEmail, FBSDSendEmail);
  Application.Run;
end.
