//------------------------------------------------------------------------------
// Date.......: 29 May 2020
// System.....: Standalone Email Sender
// Program ID.: BSD_SendEmail
// Platform...: Lazarus (Winblows, Linux, Raspbian & macOS)
// Author.....: Francois De Bruin Meyer (BlueCrane Software Development CC)
//------------------------------------------------------------------------------
// History....: 29 May 2020 - Create first version
//------------------------------------------------------------------------------

unit ldSendEmail;

{$MODE Delphi}

//------------------------------------------------------------------------------
// Uses clause
//------------------------------------------------------------------------------
interface

uses
  LCLIntf, LCLType, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, ExtCtrls, RichMemo;

//------------------------------------------------------------------------------
// Declarations
//------------------------------------------------------------------------------
type

  { TFldSendEmail }

   TFldSendEmail = class(TForm)
      edtSubject: TEdit;
      Image1: TImage;
      Label1: TLabel;
      edtTo: TEdit;
      edtCc: TEdit;
      edtBcc: TEdit;
      btnAttachments: TButton;
      Label2: TLabel;
      Label3: TLabel;
      Label4: TLabel;
      lvAttachments: TListView;
      btnCancel: TButton;
      bntSend: TButton;
      btnDelete: TButton;
      cbRead: TCheckBox;
      cbDeliver: TCheckBox;
      edtBody: TRichMemo;
      dlgOpen: TOpenDialog;
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
   StrTo         : string;        // Comma delimited list of 'To' recipients
   StrCc         : string;        // Comma delimited list of 'Cc' recipients
   StrBcc        : string;        // Comma delimited list of 'Bcc' recipients
   StrSubject    : string;        // Subject line of the email
   SMTPHost      : string;        // Extracted SMTP Hostname
   SMTPUser      : string;        // Extracted SMTP User ID
   SMTPPass      : string;        // Extracted SMTP Password


public   { Public declarations }

end;

//------------------------------------------------------------------------------
// Global variables
//------------------------------------------------------------------------------
var
   FldSendEmail: TFldSendEmail;

{$IFDEF WINDOWS}
   function  cmdlOptions(OptList : string; CmdLine, ParmStr : TStringList): integer; cdecl; external 'BSD_Utilities';
{$ELSE}
   function  cmdlOptions(OptList : string; CmdLine, ParmStr : TStringList): integer; cdecl; external 'libbsd_utilities';
{$ENDIF}

implementation

{$R *.lfm}

//------------------------------------------------------------------------------
// Executed when the form is created
//------------------------------------------------------------------------------
procedure TFldSendEmail. FormCreate( Sender: TObject);
var
   idx, NumParms  : integer;
   Params, Args   : TStringList;

begin

//--- Check whether any paramters were passed and retrieve if so

   try

      Params  := TStringList.Create;
      Args    := TStringList.Create;

      for idx := 1 to ParamCount do
         Args.Add(ParamStr(idx));

//--- Call and execute the cmdlOptions function in the BSD_Utilities DLL

      NumParms := cmdlOptions('F:P:A:B:E:T:C:S:f:p:a:b:e:t:c:s:', Args, Params);

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

            idx := idx + 2;

         end;

      end;

   finally

//      Params.Free;
      Args.Free;

   end;

end;

//------------------------------------------------------------------------------
// Executed before the form is Displayed
//------------------------------------------------------------------------------
procedure TFldSendEmail. FormShow( Sender: TObject);
var
   idx          : integer;
   SMTPParams   : TStringList;  // Final list of SMTP parameters
   AttachList   : TStringList;  // Final list of attahcments
   ThisItem     : TListItem;    // Used for adding attachements to the Attachment ListView

begin

   btnDelete.Enabled := false;

   if Trim(StrCaption) <> '' then
      FldSendEmail.Caption := StrCaption;

   if Trim(StrTo) <> '' then
      edtTo.Text := StrTo;

   if Trim(StrCc) <> '' then
      edtCc.Text := StrCc;

   if Trim(StrBcc) <> '' then
      edtBcc.Text := StrBcc;

   if Trim(StrSubject) <> '' then
      edtSubject.Text := StrSubject;

//--- Extract the SMTP Parameters

   SMTPParams := TStringList.Create;
   ExtractStrings(['|'],[' '],PChar(SMTPParms),SMTPParams);

   SMTPHost := SMTPParams.Strings[0];
   SMTPUser := SMTPParams.Strings[1];
   SMTPPass := SMTPParams.Strings[2];

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

   SMTPParams.Free;
   AttachList.Free;

end;

//---------------------------------------------------------------------------
// User clicked on the Close button
//---------------------------------------------------------------------------
procedure TFldSendEmail.btnCancelClick(Sender: TObject);
begin

   Close;

end;

//---------------------------------------------------------------------------
// User clicked on the Send button
//---------------------------------------------------------------------------
procedure TFldSendEmail.bntSendClick(Sender: TObject);
var
   idx                         : integer;
   AttachList, Body, Delimiter : string;

begin

//--- Package the Attachments

   if lvAttachments.Items.Count > 0 then begin

      AttachList := '';
      Delimiter  := '';

      for idx := 0 to lvAttachments.Items.Count - 1 do begin

         if lvAttachments.Items.Item[idx].Caption <> '' then
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

   btnCancelClick(Sender);

end;

//---------------------------------------------------------------------------
// User clicked on the Delete button
//---------------------------------------------------------------------------
procedure TFldSendEmail.btnDeleteClick(Sender: TObject);
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
procedure TFldSendEmail.btnAttachmentsClick(Sender: TObject);
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
procedure TFldSendEmail.lvAttachmentsEditing(Sender: TObject; Item: TListItem; var AllowEdit: Boolean);
begin
   AllowEdit := false;
end;


//---------------------------------------------------------------------------
// User selected/deselected an atachment in the Listview
//---------------------------------------------------------------------------
procedure TFldSendEmail.lvAttachmentsClick(Sender: TObject);
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
