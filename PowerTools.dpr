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
  T: TTasks;
  J: TArray<TTask>;
  I: Integer;

procedure LogTask(X: TTask);
Begin
  Writeln('[Task]  ', X.Name);
  Writeln(#9, 'Path = ', X.Path);
  Writeln(#9, 'Enabled = ', BoolToStr(X.Enabled, True));
  Writeln(#9, 'State = ', X.State);
  Writeln;
End;

begin
  ReportMemoryLeaksOnShutdown := True;
  try
    CoInitialize(Nil);
    T := TTasks.Create;
    J := T.List;
    Writeln('There are ', Length(J), ' running tasks.');
    Writeln;
    for I := 0 to High(J) do
      LogTask(J[I]);
    T.Free;
    CoUninitialize;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
  Readln;
end.
