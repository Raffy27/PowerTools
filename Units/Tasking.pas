unit Tasking;

interface

uses
  System.SysUtils,
  PowerShell;

type
  TTrigger = (trMinute, trHour, trDay, trWeek, trMonth, trLogon, trStartup, trIdle);

  TTask = record
    Name:   String;
    Desc:   String;
    Cmd:    String;
    Param:  String;
    Dir:    String;
    Trigger:TTrigger;
  end;

  Tasks = class
  public
    class function Exists(Name: String): Boolean;
    class procedure Create;
    class procedure Remove(Name: String);
    class function List: TArray<TTask>;
  end;

  EShellError  = class(Exception);
  EAccessError = class(Exception);

var
  Shell: TPowerShell;

implementation

procedure CheckPS;
begin
  if Not Assigned(Shell) then
    raise EShellError.Create('No PowerShell instance was passed to ' + Tasks.UnitName + '.Shell');
end;

class function Tasks.Exists(Name: string): Boolean;
begin
  CheckPS;
end;

class procedure Tasks.Create;
begin
  CheckPS;
end;

class procedure Tasks.Remove(Name: string);
begin
  CheckPS;
end;

class function Tasks.List: TArray<TTask>;
begin
  CheckPS;
end;

end.
