unit PowerShell;

interface

uses
  System.SysUtils,
  ComObj,
  Variants;

type

  TStrCallback = procedure(Line: String);

  /// <summary>A class that allows communication with a single instance of PowerShell</summary>
  TPowerShell = class
  const
    /// <summary>Unique string used for the detection of the end of command output</summary>
    CmdEnd = 'CmdEnd';
    CmdJoin = '%s;"%s"';
    CmdJoinErr = 'try{%s}catch{$_}"%s"';
    CmdJoinNull = '$Null=&{%s}';
    WshShell = 'WScript.Shell';
    PSCmd = 'powershell -NoLogo -NoProfile';
  private
    FShell: OleVariant;
    FExec: OleVariant;

    function FullCommand(Cmd: String; NoRes, Err: Boolean): String;
  public
    /// <summary>Creates an instance of PowerShell</summary>
    constructor Create;
    /// <summary>Kills the assigned instance of PowerShell</summary>
    destructor Destroy; override;
    /// <summary>Executes a PowerShell command</summary>
    /// <remarks>Does not wait for the command to finish</remarks>
    procedure Exec(Cmd: String);
    /// <summary>Executes a PowerShell command and return its output</summary>
    /// <returns>Trimmed output of the given PowerShell command</returns>
    function Get(Cmd: String; Err: Boolean): String; overload;
    /// <summary>Executes a PowerShell command and return its output line by line using the provided callback</summary>
    procedure Get(Cmd: String; Err: Boolean; Call: TStrCallback); overload;
  end;

implementation

constructor TPowerShell.Create;
begin
  FShell := CreateOleObject(WshShell);
  FExec := FShell.Exec(PSCmd);
end;

destructor TPowerShell.Destroy;
Begin
  FExec := Unassigned;
  FShell := Unassigned;
End;

function TPowerShell.FullCommand(Cmd: string; NoRes, Err: Boolean): String;
begin
  if NoRes then
    Result := Format(CmdJoinNull, [Cmd])
  else if Err then
    Result := Format(CmdJoinErr, [Cmd, CmdEnd])
  else
    Result := Format(CmdJoin, [Cmd, CmdEnd]);
end;

procedure TPowerShell.Exec(Cmd: string);
begin
  FExec.StdIn.WriteLine(FullCommand(Cmd, True, False));
end;

function TPowerShell.Get(Cmd: String; Err: Boolean): String;
var
  L: String;
Begin
  FExec.StdIn.WriteLine(FullCommand(Cmd, False, Err));
  Repeat
    L := FExec.StdOut.ReadLine;
  Until L.Contains(CmdEnd);
  L := String.Empty;
  Repeat
    Result := Result + L + #10;
    L := FExec.StdOut.ReadLine;
  Until L = CmdEnd;
  Result := Result.Trim;
End;

procedure TPowerShell.Get(Cmd: String; Err: Boolean; Call: TStrCallback);
var
  L: String;
Begin
  if not Assigned(Call) then
    raise Exception.Create('Get callback is unassigned');
  FExec.StdIn.WriteLine(FullCommand(Cmd, False, Err));
  Repeat
    L := FExec.StdOut.ReadLine;
  Until L.Contains(CmdEnd);
  Repeat
    L := FExec.StdOut.ReadLine;
    if L = CmdEnd then
      Break
    else
      Call(L);
  Until False;
End;

end.
