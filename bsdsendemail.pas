//------------------------------------------------------------------------------
// Date.......: 29 May 2020
// System.....: Standalone Email Sender
// Program ID.: BSD_SendEmail
// Platform...: Lazarus (Winblows, Linux, Raspbian & macOS)
// Author.....: Francois De Bruin Meyer (BlueCrane Software Development CC)
//------------------------------------------------------------------------------
// History....: 29 May 2020 - Create first version
//------------------------------------------------------------------------------

unit bsdsendemail;

{$MODE Delphi}

//------------------------------------------------------------------------------
// Uses clause
//------------------------------------------------------------------------------
interface

uses
  LCLIntf, LCLType, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, ExtCtrls, ActnList, StdActns, ButtonPanel,
  Buttons, RichMemo;

//------------------------------------------------------------------------------
// Declarations
//------------------------------------------------------------------------------
type

  { TFBSDSendEmail }

   TFBSDSendEmail = class(TForm)
  BitBtn1: TBitBtn;
  BitBtn10: TBitBtn;
  BitBtn2: TBitBtn;
  BitBtn3: TBitBtn;
  BitBtn4: TBitBtn;
  BitBtn5: TBitBtn;
  BitBtn6: TBitBtn;
  BitBtn7: TBitBtn;
  BitBtn8: TBitBtn;
  BitBtn9: TBitBtn;
  bntSend: TButton;
  btnAttachments: TButton;
  btnCancel: TButton;
  btnDelete: TButton;
  cbDeliver: TCheckBox;
  cbRead: TCheckBox;
  EditLeft: TAction;
  EditCentre: TAction;
  EditRight: TAction;
  EditBlock: TAction;
  EditItalic: TAction;
  EditUnder: TAction;
  EditBold: TAction;
  alAction: TActionList;
  EditCopy: TEditCopy;
  EditCut: TEditCut;
  EditPaste: TEditPaste;
  edtBcc: TEdit;
  edtBody: TRichMemo;
      dlgOpen: TOpenDialog;
      edtCc: TEdit;
      edtSubject: TEdit;
      edtTo: TEdit;
      Image1: TImage;
      Label1: TLabel;
      Label2: TLabel;
      Label3: TLabel;
      Label4: TLabel;
      lvAttachments: TListView;
      Panel1: TPanel;
      Panel2: TPanel;
      Panel3: TPanel;
      procedure btnCancelClick(Sender: TObject);
      procedure bntSendClick(Sender: TObject);
      procedure btnDeleteClick(Sender: TObject);
      procedure btnAttachmentsClick(Sender: TObject);
      procedure FormCreate( Sender: TObject);
      procedure FormShow( Sender: TObject);
      procedure lvAttachmentsEditing(Sender: TObject; Item: TListItem; var AllowEdit: Boolean);
      procedure lvAttachmentsClick(Sender: TObject);

private  { Private declarations }

   StrCaption    : string;        // Caption to be display at the top of the form
   SMTPParms     : string;        // '|' Delimited list of SMTP Parameters - Host, User and Password
   StrAttachLst  : string;        // '|' Delimited list of attachments
   StrEmailBody  : string;        // '|' Delimited content lines of the email
   StrFrom       : string;        // From email address
   StrTo         : string;        // Comma delimited list of 'To' recipients
   StrCc         : string;        // Comma delimited list of 'Cc' recipients
   StrBcc        : string;        // Comma delimited list of 'Bcc' recipients
   StrSubject    : string;        // Subject line of the email

public   { Public declarations }

end;

//------------------------------------------------------------------------------
// Global variables
//------------------------------------------------------------------------------
var
   FBSDSendEmail: TFBSDSendEmail;

{$IFDEF DARWIN}
   function  cmdlOptions(OptList : string; CmdLine, ParmStr : TStringList): integer; cdecl; external 'libbsd_utilities.dylib';
   function SendMimeMail(From, ToStr, CcStr, BccStr, Subject, Body, Attach, SMTPStr : string): boolean; cdecl; external 'libbsd_utilities.dylib';
{$ENDIF}
{$IFDEF LINUX}
   function  cmdlOptions(OptList : string; CmdLine, ParmStr : TStringList): integer; cdecl; external 'libbsd_utilities.so';
   function SendMimeMail(From, ToStr, CcStr, BccStr, Subject, Body, Attach, SMTPStr : string): boolean; cdecl; external 'libbsd_utilities.so';
{$ENDIF}
{$IFDEF WINDOWS}
   function  cmdlOptions(OptList : string; CmdLine, ParmStr : TStringList): integer; cdecl; external 'BSD_Utilities.dll';
   function SendMimeMail(From, ToStr, CcStr, BccStr, Subject, Body, Attach, SMTPStr : string): boolean; cdecl; external 'BSD_Utilities.dll';
{$ENDIF}

implementation

{$R *.lfm}

//------------------------------------------------------------------------------
// Executed when the form is created
//------------------------------------------------------------------------------
procedure TFBSDSendEmail. FormCreate( Sender: TObject);
var
   idx, NumParms  : integer;
   Params, Args   : TStringList;

begin

//--- Check whether any paramters were passed and retrieve if so

   try

//--- Params is modified by the DLL and as such will be freed/destroyed by
//--- the DLL. If it is freed here then it will result in an error when the DLL
//--- tries to free the allocated memory

      Params  := TStringList.Create;
      Args    := TStringList.Create;

      for idx := 1 to ParamCount do
         Args.Add(ParamStr(idx));

//--- Call and execute the cmdlOptions function in the BSD_Utilities DLL

      NumParms := cmdlOptions('F:P:A:B:E:T:C:S:O:f:p:a:b:e:t:c:s:o:', Args, Params);

      if NumParms > 0 then begin

         idx      := 0;
         NumParms := NumParms * 2;

         while idx < Params.Count do begin

            if ((Params.Strings[idx] = 'F') or (Params.Strings[idx] = 'f')) then
               StrCaption := Params.Strings[idx + 1];

            if ((Params.Strings[idx] = 'P') or (Params.Strings[idx] = 'p')) then
               SMTPParms := Params.Strings[idx + 1];

            if ((Params.Strings[idx] = 'A') or (Params.Strings[idx] = 'a')) then
               StrAttachLst := Params.Strings[idx + 1];

            if ((Params.Strings[idx] = 'B') or (Params.Strings[idx] = 'b')) then
               StrBcc := Params.Strings[idx + 1];

            if ((Params.Strings[idx] = 'E') or (Params.Strings[idx] = 'e')) then
               StrEmailBody := Params.Strings[idx + 1];

            if ((Params.Strings[idx] = 'T') or (Params.Strings[idx] = 't')) then
               StrTo := Params.Strings[idx + 1];

            if ((Params.Strings[idx] = 'C') or (Params.Strings[idx] = 'c')) then
               StrCc := Params.Strings[idx + 1];

            if ((Params.Strings[idx] = 'S') or (Params.Strings[idx] = 's')) then
               StrSubject := Params.Strings[idx + 1];

            if ((Params.Strings[idx] = 'O') or (Params.Strings[idx] = 'o')) then
               StrFrom := Params.Strings[idx + 1];

            idx := idx + 2;

         end;

      end;

   finally

      Args.Free;

   end;

end;

//------------------------------------------------------------------------------
// Executed before the form is Displayed
//------------------------------------------------------------------------------
procedure TFBSDSendEmail. FormShow( Sender: TObject);
var
   idx         : integer;
   AttachList  : TStringList;  // Final list of attahcments
   ThisItem    : TListItem;    // Used for adding attachements to the Attachment ListView

begin

   btnDelete.Enabled := false;

   if Trim(StrCaption) <> '' then
      FBSDSendEmail.Caption := StrCaption;

   if Trim(StrTo) <> '' then
      edtTo.Text := StrTo;

   if Trim(StrCc) <> '' then
      edtCc.Text := StrCc;

   if Trim(StrBcc) <> '' then
      edtBcc.Text := StrBcc;

   if Trim(StrSubject) <> '' then
      edtSubject.Text := StrSubject;

//--- Extract the List of files to be attached

   AttachList := TStringList.Create;
   ExtractStrings(['|'],[' '],PChar(StrAttachLst),AttachList);

   lvAttachments.Clear;

   for idx := 0 to AttachList.Count - 1 do begin

      ThisItem := lvAttachments.Items.Add;
      ThisItem.Caption := AttachList.Strings[idx];

   end;

//--- Extract the body of the email

   ExtractStrings(['|'],['*'],PChar(StrEmailBody),edtBody.Lines);

//--- Clean up

   AttachList.Free;

end;

//---------------------------------------------------------------------------
// User clicked on the Close button
//---------------------------------------------------------------------------
procedure TFBSDSendEmail.btnCancelClick(Sender: TObject);
begin

   Close;

end;

//---------------------------------------------------------------------------
// User clicked on the Send button
//---------------------------------------------------------------------------
procedure TFBSDSendEmail.bntSendClick(Sender: TObject);
var
   idx                         : integer;
   AttachList, Body, Delimiter : string;

begin

//--- Do some basic checking of required fields before sending

   if ((Trim(edtTo.Text) = '') and (Trim(edtCc.Text) = '') and (Trim(edtBcc.Text) = '')) then begin

      Application.MessageBox('At least 1 of ''To:'', ''Cc:'' or ''Bcc:'' must be specified.','BSD Send Email',(MB_OK + MB_ICONWARNING));
      edtTo.SetFocus;
      Exit;

   end;

   if Trim(edtSubject.Text) = '' then begin

      if Application.MessageBox('Note: ''Subject:'' is empty. You can: ' + #10 + #10 + #10 + 'Click [Ok] to proceed with an empty subject line; or' + #10 + #10 + 'Click [Cancel] to return.','BSD Send Email',(MB_OKCANCEL + MB_ICONWARNING)) = ID_CANCEL then begin

         edtSubject.SetFocus;
         Exit;

      end;

   end;

   if edtBody.Lines.Count = 0 then begin

      if Application.MessageBox('Note: Email text is empty. You can: ' + #10 + #10 + #10 + 'Click [Ok] to proceed with an empty Email; or' + #10 + #10 + 'Click [Cancel] to return.','BSD Send Email',(MB_OKCANCEL + MB_ICONWARNING)) = ID_CANCEL then begin

         edtBody.SetFocus;
         Exit;

      end;

   end;

//--- Package the Attachments. Only files that actually exist are added

   if lvAttachments.Items.Count > 0 then begin

      AttachList := '';
      Delimiter  := '';

      for idx := 0 to lvAttachments.Items.Count - 1 do begin

         if FileExists(lvAttachments.Items.Item[idx].Caption) = True then
            AttachList := AttachList + Delimiter + lvAttachments.Items.Item[idx].Caption;

         Delimiter := '|';

      end;

   end;

//--- Package the body of the email

   if edtBody.Lines.Count > 0 then begin

      Body      := '';
      Delimiter := '';

      for idx := 0 to edtBody.Lines.Count - 1 do begin

         if Trim(edtBody.Lines[idx]) = '' then
            Body := Body + Delimiter + ' '
         else
            Body := Body + Delimiter + edtBody.Lines[idx];

         Delimiter := '|';

      end;

   end;

//--- Send the email

   if SendMimeMail(StrFrom, edtTo.Text, edtCc.Text, edtBcc.Text, edtSubject.Text, Body, AttachList, SMTPParms) = False then
      Application.MessageBox('Sending Email failed! Please check Email set-up details.','BSD Send Email',(MB_OK + MB_ICONSTOP))
   else begin

      Application.MessageBox('Send Email completed.','BSD Send Email',(MB_OK + MB_ICONINFORMATION));
      btnCancelClick(Sender);

   end;

end;

//---------------------------------------------------------------------------
// User clicked on the Delete button
//---------------------------------------------------------------------------
procedure TFBSDSendEmail.btnDeleteClick(Sender: TObject);
var
   idx : integer;

begin

//--- Step through the items in the listview and delete the checked items

   for idx := lvAttachments.Items.Count - 1 downto 0 do begin

      if (lvAttachments.Items.Item[idx].Checked = True) then
         lvAttachments.Items.Item[idx].Delete;

   end;

   btnDelete.Enabled := False;

end;

//---------------------------------------------------------------------------
// User clicked on the Attachments button
//---------------------------------------------------------------------------
procedure TFBSDSendEmail.btnAttachmentsClick(Sender: TObject);
var
   idx      : integer;
   ThisList : TListItem;

begin

   if dlgOpen.Execute = True then begin

      for idx := 0 to dlgOpen.Files.Count - 1 do begin

         ThisList := lvAttachments.Items.Add;
         ThisList.Caption := dlgOpen.Files[idx];

      end;

   end;

end;

//---------------------------------------------------------------------------
// Function to prevent an inline edit of the ListView items
//---------------------------------------------------------------------------
procedure TFBSDSendEmail.lvAttachmentsEditing(Sender: TObject; Item: TListItem; var AllowEdit: Boolean);
begin
   AllowEdit := false;
end;


//---------------------------------------------------------------------------
// User selected/deselected an atachment in the Listview
//---------------------------------------------------------------------------
procedure TFBSDSendEmail.lvAttachmentsClick(Sender: TObject);
var
   idx, SelCount : integer;

begin
//--- Check how many items are selected then Enable Delete button if 1 or more

   SelCount := 0;

   for idx := 0 to lvAttachments.Items.Count - 1 do begin
      if (lvAttachments.Items.Item[idx].Checked = true) then
         inc(SelCount);
   end;

   if (SelCount > 0) then
      btnDelete.Enabled := True
   else
      btnDelete.Enabled := False;

end;

//---------------------------------------------------------------------------
end.
