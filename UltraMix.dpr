library UltraMix;

{$IFNDEF DEBUG}
  {$WEAKLINKRTTI ON}
  {$RTTI EXPLICIT METHODS([]) FIELDS([]) PROPERTIES([])}
{$ENDIF}

uses
  Windows,
  apiPlugin,
  UltraMixMain in 'UltraMixMain.pas';

{$IFNDEF DEBUG}
  {$SetPEFlags IMAGE_FILE_DEBUG_STRIPPED}
  {$SetPEFlags IMAGE_FILE_LINE_NUMS_STRIPPED}
  {$SetPEFlags IMAGE_FILE_LOCAL_SYMS_STRIPPED}
  {$SetPEFlags IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP}
  {$SetPEFlags IMAGE_FILE_NET_RUN_FROM_SWAP}
  {$SetPEFlags IMAGE_FILE_EXECUTABLE_IMAGE}
{$ENDIF}

function AIMPPluginGetHeader(out Header: IAIMPPlugin): HRESULT; stdcall;
begin
{$IFDEF DEBUG}
    ReportMemoryLeaksOnShutdown := True;
{$ENDIF}
  try
    Header := TUMPlugin.Create;
    Result := S_OK;
  except
    Result := E_UNEXPECTED;
  end;
end;

exports
  AIMPPluginGetHeader;

begin
  //
end.
