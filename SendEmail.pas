//===========================================================================
// Program ID...: SendEmail
// Author.......: Francois De Bruin Meyer
// Copyright....: BlueCrane Software Development CC
// Date.........: 06 May 2020
//---------------------------------------------------------------------------
// Description..: Command line program to send an email with attachments
//===========================================================================

//---------------------------------------------------------------------------
// Entry Point
//---------------------------------------------------------------------------
program SendEmail;

{$MODE Delphi}

//{$APPTYPE GUI}

uses
  LCLIntf, LCLType, LMessages, Messages, Registry, Forms, Dialogs, Interfaces,
  SysUtils, SMTPSend, MimeMess, MimePart, SynaUtil, Classes;

//-----------------------------------------------------------------------------
// Variable Definitions
//-----------------------------------------------------------------------------
var

   idx          : integer;     // Used for adding attachments
   NumParms     : integer;     // Number of Parameters passed

   SendResult   : boolean;     // Result returned by the SMTP send function

   ToStr        : string;      // To addresses
   CcStr        : string;      // Cc addresses
   BccStr       : string;      // Bcc addresses
   Subject      : string;      // Email Subject
   Body         : string;      // Body of the email
   Attach       : string;      // Delimited string containing files to be attached
   SMTPStr      : string;      // Delimited string containing the SMTP parameters
   Sender       : string;      // Email address of the Sender
   SMTPHost     : string;      // SMTP Host name
   SMTPUser     : string;      // SMTP User name
   SMTPPasswd   : string;      // SMTP Password
   ToList       : string;      // Final To address list
   CcList       : string;      // Final Cc address list
   AddrList     : string;      // Holds addresses that the SMTP server must send to
   AddrId       : string;      // Holds an extracted email address
   ThisList     : string;      // Used in the transformation of the address lists

   AttachList   : TStringList; // Final list of attahcments
   SMTPParms    : TStringList; // Final list of SMTP parameters
   Content      : TStringList; // Final string list containing the body of the ema il
   Mime         : TMimeMess;   // Mail message in MIME format
   MimePtr      : TMimePart;   // Pointer to the MIME messsage being constructed

begin

//---------------------------------------------------------------------------//
// Parameters passed on the command line:                                    //
//  Parameter  1  = ToStr                                                    //
//  Parameter  2  = CcStr                                                    //
//  Parameter  3  = BccStr                                                   //
//  Parameter  4  = Subject                                                  //
//  Parameter  5  = Content                                                  //
//  Parameter  6  = List of attachments                                      //
//  Parameter  7  = SMTP Parameters                                          //
//  Parameter  8  = From address                                             //
//---------------------------------------------------------------------------//

//--- Check whether the correct number of parameters were passed

   NumParms := ParamCount;

//--- Verify the number of parameters

   if (NumParms <> 8) then begin
      MessageDlg('Invalid use of SendMail - RC = ' + IntToStr(NumParms), mtWarning, [mbOK], 0);
      Application.Terminate;
      Exit;
   end;

//--- Transfer the Parameters

   ToStr    := ParamStr(1);
   CcStr    := ParamStr(2);
   BccStr   := ParamStr(3);
   Subject  := ParamStr(4);
   Body     := ParamStr(5);
   Attach   := ParamStr(6);
   SMTPStr  := ParamStr(7);
   Sender   := ParamStr(8);

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

//--- Send the email

   if (Attachlist.Count > 0) then begin

      Mime := TMimeMess.Create;

      try

// Set the headers. The various address lists (To, Cc and Bcc) can contain more
// than 1 address but these must be seperted by commas and there should be no
// spaces between addresses. The Bcc address list is not added to the Header.

         ThisList := ToStr;

         repeat

            ToList := Trim( FetchEx( ThisList, ',', '"' ) );

            if ( ToList <> '' ) then
               Mime.Header.ToList.Append(ToList);

         until ThisList = '';

         ThisList := CcStr;

         repeat

            CcList := Trim( FetchEx( ThisList, ',', '"' ) );

            if ( CcList <> '' ) then
               Mime.Header.CCList.Append(CcList);

         until ThisList = '';

         AddrList := '';
         ThisList := ToStr + ',' + CcStr + ',' + BccStr;

         repeat

            AddrId := Trim( FetchEx( ThisList, ',', '"' ) );

            if ( AddrId <> '' ) then
               AddrList := AddrList + ',' + AddrId;

         until ThisList = '';

         Mime.Header.Subject     := Subject;
         Mime.Header.From        := Sender;

// Create a MultiPart part

         MimePtr := Mime.AddPartMultipart('mixed',Nil);

// Add the mail text as the first part

         Mime.AddPartText(Content,MimePtr);

// Add all atachments

         for idx := 0 to AttachList.Count - 1 do
            Mime.AddPartBinaryFromFile(AttachList[idx],MimePtr);

// Compose the message to comply with MIME requirements

         Mime.EncodeMessage;

         Mime;

// Now the messsage can be sent using SendToRaw

         SendResult := SendToRaw(Sender, ToStr, SMTPHost, Mime.Lines, SMTPUser, SMTPPasswd);

         if SendResult = False then
            MessageDlg('Email message could not be sent', mtWarning, [mbOK], 0)
         else
            MessageDlg('Email message successfully sent', mtWarning, [mbOK], 0)

      finally
         Mime.Free;
      end;
      finally
      end;

   end else begin

      SendResult := SendToEx(Sender, ToStr, Subject, SMTPHost, Content, SMTPUser, SMTPPasswd);

      if SendResult = False then
         MessageDlg('Email message could not be sent', mtWarning, [mbOK], 0)
      else
         MessageDlg('Email message successfully sent', mtWarning, [mbOK], 0)

   end;

//--- Delete the lists we used

   Content.Free;
   AttachList.Free;
   SMTPParms.Free;

   Application.Terminate;
   Exit;

end.
