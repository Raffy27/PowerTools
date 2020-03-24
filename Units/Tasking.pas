unit Tasking;

interface

uses
  System.SysUtils,
  ComObj,
  ActiveX,
  Variants;

type

  TTask = record
    Name: String;
    Path: String;
    Enabled: Boolean;
    State: Integer;
    
    constructor Create(V: OleVariant);
  end;

  TTasks = class
  private
    FServ: OleVariant;
  public
    constructor Create;
    destructor Destroy; override;
    function List: TArray<TTask>;
    function ListRunning: TArray<TTask>;
  end;

implementation

constructor TTask.Create(V: OleVariant);
begin
  Name := V.Name;
  Path := V.Path;
  Enabled := V.Enabled;
  State := V.State;
end;

constructor TTasks.Create;
begin
  FServ := CreateOleObject('Schedule.Service');
  FServ.Connect;
end;

destructor TTasks.Destroy;
begin
  FServ := Unassigned;
end;

function TTasks.List;
var
  L, V: OleVariant;
  Enum: IEnumVariant;
  I, S: Cardinal;
begin
  S := 0;
  L := FServ.GetFolder('\').GetTasks(0);
  SetLength(Result, Cardinal(L.Count));
  Enum := IUnknown(L._NewEnum) as IEnumVariant;
  while Enum.Next(1, V, I) = S_OK do Begin
    Result[S] := TTask.Create(V);
    Inc(S);  
  End;
end;

function TTasks.ListRunning;
var
  L, V: OleVariant;
  E: IEnumVariant;
  I, S: Cardinal;
begin
  S := 0;
  L := FServ.GetRunningTasks(0);
  SetLength(Result, Cardinal(L.Count));
  E := IUnknown(L._NewEnum) as IEnumVariant;
  while E.Next(1, V, I) = S_OK do Begin
    Result[S] := TTask.Create(V);
    Inc(S);
  End;
end;

end.
