//==============================================================================
// Program ID...: clOpt
// Author.......: Francois De Bruin Meyer
// Copyright....: BlueCrane Software Development CC
// Date.........: 07 May 2020
//------------------------------------------------------------------------------
// Description..: App to show the use of clOpt - a function to extract and
// .............: return parameters passed when a program is invoked
//==============================================================================

unit cloptshow;

{$mode objfpc}{$H+}

interface

uses
   Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ComCtrls;

type

   { TFclOptShow }

   TFclOptShow = class( TForm)
      lv1: TListView;
      procedure FormShow( Sender: TObject);

   private

      function cmdlOpt(OptList : string; Options, Parms : TStringList) : integer;

   public

   end;

var
   FclOptShow: TFclOptShow;

implementation

{$R *.lfm}

{ TFclOptShow }

//------------------------------------------------------------------------------
// Executed when the form is shown
//------------------------------------------------------------------------------
procedure TFclOptShow. FormShow( Sender: TObject);
var
   idx             : integer;
   Options, Params : TStringList;
   ThisItem        : TListItem;

begin

   Options := TStringList.Create;
   Params  := TStringList.Create;

   if cmdlOpt('MD:',Options,Params) > 0 then begin

      for idx := 0 to Options.Count - 1 do begin

         ThisItem := lv1.Items.Add;
         ThisItem.Caption := Options.Strings[idx];

         if Options.Strings[idx] = 'M' then
            ThisItem.SubItems.Add('No value expected')
         else if Options.Strings[idx] = 'D' then
            ThisItem.SubItems.Add(Params.Strings[idx])
         else
            ThisItem.SubItems.Add(Params.Strings[idx]);

      end;

   end;

   Options.Free;
   Params.Free;

end;

//---------------------------------------------------------------------------
// Function to extract and return the parameters that were passed on the
// command line when the application was invoked.
//
// An Option list of "H:L:mu:" would expect a command line similar to:
//   -H followed by a parameter e.g. -Hwww.sourcingmethods.com
//   -L followed by a parameter e.g. -L1
//   -m
//   -u followed by a parameter e.g. -uFrancois
//
// The calling function must create the following TStringList variables:
//
//   Options
//   Parms
//
// The function returns:
//
//    0 if no command line parameters were passed
//    $ in the Parms string if a value was expected but not found
//    # in the Parms string if an unknown parameter was found
//   -1 if the switch '-' could not be found
//   The number of parameters found if no error were found
//---------------------------------------------------------------------------
function TFclOptShow.cmdlOpt(OptList : string; Options, Parms : TStringList) : integer;
var
   idx1, idx2           : integer;
   Found                : boolean;
   ThisParm, ThisOption : string;

begin

   if ParamCount < 1 then begin

      Result := 0;
      Exit;

   end;

//--- Extract the parameters that were passed from ParamStr

   for idx1 := 1 to ParamCount do begin

      ThisParm := ParamStr(idx1);

//--- First character of the argument must be the switch character ('-')

      if ThisParm.SubString(0,1) <> '-' then begin

         Result := -1;
         Exit;

      end;

//--- Extract the second character and search for it in OptList

      ThisOption := ThisParm.SubString(1,1);
      Found := False;

      for idx2 := 0 to OptList.Length do begin

         if OptList.SubString(idx2,1) = ThisOption then begin

            Found := True;
            Options.Add(ThisOption);

//--- If this Option is followed by ":" in the OptList then a parameter is
//    expected. Extract the parameter if it is expected

            if OptList.SubString(idx2 + 1,1) = ':' then begin

               if ThisParm.Length < 3 then
                  Parms.Add('$')
               else
                  Parms.Add(ThisParm.SubString(2, ThisParm.Length - 1));

            end else
               Parms.Add('$');

         end;

      end;

      if Found = False then begin
         Options.Add(ThisOption);
         Parms.Add('#');
      end;

   end;

   Result := Parms.Count;

end;

//------------------------------------------------------------------------------

end.

