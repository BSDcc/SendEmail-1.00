//==============================================================================
// Program ID...: SendEmail
// Author.......: Francois De Bruin Meyer
// Copyright....: BlueCrane Software Development CC
// Date.........: 06 May 2020
//------------------------------------------------------------------------------
// Description..: Command line program to send an email with attachments
//==============================================================================

program SendEmail;

{$MODE Delphi}

//{$APPTYPE GUI}

uses
   Forms, Dialogs, Interfaces, SysUtils, SMTPSend, MimeMess, MimePart,
   SynaUtil, Classes;

//------------------------------------------------------------------------------
// Function to use MIME to send an email
//------------------------------------------------------------------------------
function DoMimeMail(From, ToStr, CcStr, BccStr, Subject, Body, Attach, SMTPStr : string): boolean;
var
   idx          : integer;     // Used for adding attachments
   NumErrors    : integer;     // Keesp track whether there were errors during send

   SendResult   : boolean;     // Result returned by the SMTP send function

   SMTPHost     : string;      // SMTP Host name
   SMTPUser     : string;      // SMTP User name
   SMTPPasswd   : string;      // SMTP Password
   AddrId       : string;      // Holds an extracted email address
   ThisStr      : string;      // Used in the transformation of the address lists

   SMTPParms    : TStringList; // Final list of SMTP parameters
   Content      : TStringList; // Final string list containing the body of the email
   AttachList   : TStringList; // Final list of attahcments
   Mime         : TMimeMess;   // Mail message in MIME format
   MimePtr      : TMimePart;   // Pointer to the MIME messsage being constructed

begin

//--- Extract the SMTP Parameters

   SMTPParms := TStringList.Create;
   ExtractStrings(['|'],[' '],PChar(SMTPStr),SMTPParms);

   SMTPHost   := SMTPParms.Strings[0];
   SMTPUser   := SMTPParms.Strings[1];
   SMTPPasswd := SMTPParms.Strings[2];

//--- Extract the List of files to be attached

   AttachList := TStringList.Create;
   ExtractStrings(['|'],[' '],PChar(Attach),AttachList);

//--- Extract the body of the email

   Content := TStringList.Create;
   ExtractStrings(['|'],['*'],PChar(Body),Content);

   Mime := TMimeMess.Create;

   try

//--- Set the headers. The various address lists (To, Cc and Bcc) can contain
//--- more than 1 address but these must be seperted by commas and there should
//--- be no spaces between addresses. The Bcc address list is not added to the
//--- Header as the data in the Header is not used to determine who to send the
//--- email to.

      ThisStr := ToStr;

      repeat

         AddrId := Trim(FetchEx(ThisStr, ',', '"'));

         if AddrId <> '' then
            Mime.Header.ToList.Append(AddrId);

      until ThisStr = '';

      ThisStr := CcStr;

      repeat

         AddrId := Trim(FetchEx(ThisStr, ',', '"'));

         if AddrId <> '' then
            Mime.Header.CCList.Append(AddrId);

      until ThisStr = '';

      Mime.Header.Subject := Subject;
      Mime.Header.From    := From;
      Mime.Header.ReplyTo := From;

//--- Create a MultiPart part

      MimePtr := Mime.AddPartMultipart('mixed',Nil);

//--- Add the mail text as the first part

      Mime.AddPartText(Content,MimePtr);

//--- Add all atachments

      if AttachList.Count > 0 then begin

         for idx := 0 to AttachList.Count - 1 do
            Mime.AddPartBinaryFromFile(AttachList[idx],MimePtr);

      end;

//--- Compose the message to comply with MIME requirements

      Mime.EncodeMessage;

//--- Now the messsage can be sent to each of the recipients (To, Cc and Bcc)
//--- using SendToRaw

      NumErrors := 0;
      ThisStr  := ToStr + ',' + CcStr + ',' + BccStr;

      repeat

         AddrId := Trim(FetchEx( ThisStr, ',', '"'));

         if AddrId <> '' then begin
            SendResult := SendToRaw(From, AddrId, SMTPHost, Mime.Lines, SMTPUser, SMTPPasswd);

            if SendResult = False then
               Inc(NumErrors);

         end;

      until ThisStr = '';

   finally
      Mime.Free;
   end;


//--- Delete the lists we used

   AttachList.Free;
   SMTPParms.Free;
   Content.Free;

  if NumErrors > 0 then
     Result := False
  else
     Result := True;

end;

//------------------------------------------------------------------------------
// Entry Point
//------------------------------------------------------------------------------
// Parameters passed on the command line:
//  Parameter  1  = ToStr - List of To email addresses separared by ','                                                    //
//  Parameter  2  = CcStr - List of Cc email addresses separated by ','
//  Parameter  3  = BccStr - List of Bcc email addresses separaetd by ','
//  Parameter  4  = Subject - Subject line of the Email
//  Parameter  5  = Content - Body of the email, each line separated by '|'
//  Parameter  6  = List of attachments - Full file names separated by '|'
//  Parameter  7  = SMTP Parameters - Host, User, Password separated by '|'
//  Parameter  8  = From address - Email address of the sender
//  Parameter  9  = Info Message: Silent = 0, Display = 1
//------------------------------------------------------------------------------
var

   NumParms     : integer;     // Number of Parameters passed

   SendResult   : boolean;     // Receives the result of the send attempt
   ShowResult   : boolean;     // Conbtrols whether a success/failure message is shown

   ToStr        : string;      // To addresses
   CcStr        : string;      // Cc addresses
   BccStr       : string;      // Bcc addresses
   Subject      : string;      // Email Subject
   Body         : string;      // Body of the email
   Attach       : string;      // Delimited string containing files to be attached
   SMTPStr      : string;      // Delimited string containing the SMTP parameters
   From         : string;      // Email address of the Sender

begin

//--- Check whether the correct number of parameters were passed

   NumParms := ParamCount;

//--- Verify the number of parameters

   if (NumParms <> 9) then begin
      MessageDlg('Invalid use of SendMail - RC = ' + IntToStr(NumParms), mtWarning, [mbOK], 0);
      Application.Terminate;
      Exit;
   end;

//--- Transfer the Parameters

   ToStr      := ParamStr(1);
   CcStr      := ParamStr(2);
   BccStr     := ParamStr(3);
   Subject    := ParamStr(4);
   Body       := ParamStr(5);
   Attach     := ParamStr(6);
   SMTPStr    := ParamStr(7);
   From       := ParamStr(8);
   ShowResult := ParamStr(9).ToBoolean();

//--- Send the email

   SendResult := DoMimeMail(From, ToStr, CcStr, BccStr, Subject, Body, Attach, SMTPStr);

   if ShowResult = True then begin

      if SendResult = False then
         MessageDlg('Email message failed to send!', mtWarning, [mbOK], 0)
      else
         MessageDlg('Email message successfully sent.', mtInformation, [mbOK], 0)

   end;

   if SendResult = True then
      System.ExitCode := 0
   else
      System.ExitCode := 1;

   Application.Terminate;
   Exit;

end.
