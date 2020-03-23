program PowerTools;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  System.SysUtils,
  System.Diagnostics,
  ActiveX,
  PowerShell in 'Units\PowerShell.pas',
  Tasking in 'Units\Tasking.pas';

var
  P: TPowerShell;
  S: TStopwatch;

begin
  try
    CoInitialize(Nil);
    P := TPowerShell.Create;
    Writeln(P.Get('"Hello"'));
    //Tasking.Shell := P;
    Tasks.Exists('Omage');
    P.Free;
    CoUninitialize;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
  Readln;
end.
