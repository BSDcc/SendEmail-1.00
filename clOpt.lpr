program clOptApp;

{$mode objfpc}{$H+}

uses
   {$IFDEF UNIX}{$IFDEF UseCThreads}
   cthreads,
   {$ENDIF}{$ENDIF}
   Interfaces, // this includes the LCL widgetset
   Forms, cloptshow
   { you can add units after this };

{$R *.res}

begin
   RequireDerivedFormResource := True;
   Application. Title :='clOptShow';
   Application. Scaled := True;
   Application. Initialize;
   Application. CreateForm( TFclOptShow, FclOptShow);
   Application. Run;
end.

