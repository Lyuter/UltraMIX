{~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

             UltraMIX - AIMP3 plugin
                  Version: 1.4.1
              Copyright (c) Lyuter
           Mail : pro100lyuter@mail.ru

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~}
unit UltraMixMain;

interface

uses  Windows, SysUtils, Classes, StrUtils,
      AIMPCustomPlugin,
      apiWrappers, apiPlaylists, apiCore,
      apiMenu, apiActions, apiObjects, apiPlugin,
      apiMessages, apiFileManager;

{$R MenuIcon.res}

const
    UM_PLUGIN_NAME              = 'UltraMIX v1.4.1';
    UM_PLUGIN_AUTHOR            = 'Author: Lyuter';
    UM_PLUGIN_SHORT_DESCRIPTION = 'Provide improved algorithm of playlist mixing';
	  UM_PLUGIN_FULL_DESCRIPTION  = 'This plugin allows you to evenly distribute ' +
                                  'all the artists in the playlist.';
    //
    UM_CAPTION                  = 'UltraMIX';
    //
    UM_HOTKEY_ID                = 'UltraMix.Hotkey';
    UM_HOTKEY_GROUPNAME_KEYPATH = 'HotkeysGroups\PlaylistSorting';
    UM_HOTKEY_DEFAULT_MOD       = AIMP_ACTION_HOTKEY_MODIFIER_SHIFT;
    UM_HOTKEY_DEFAULT_KEY       = 77; // 'M' key
    //
    UM_CONTEXTMENU_ID           = 'UltraMix.Menu';
    UM_CONTEXTMENU_ICON         = 'MENU_ICON';

type

  TUMMessageHook = class(TInterfacedObject, IAIMPMessageHook)
  public
    procedure CoreMessage(Message: DWORD; Param1: Integer; Param2: Pointer;
                                                    var Result: HRESULT); stdcall;
  end;

  TUMPlugin = class(TAIMPCustomPlugin)
  private
    function MakeLocalDefaultHotkey: Integer;
    procedure CreateActionAndMenu;
    function GetBuiltInMenu(ID: Integer): IAIMPMenuItem;
    function LoadMenuIcon(const ResName: string): IAIMPImage;
  protected
    UMMessageHook: TUMMessageHook;
    function InfoGet(Index: Integer): PWideChar; override; stdcall;
    function InfoGetCategories: Cardinal; override; stdcall;
    function Initialize(Core: IAIMPCore): HRESULT; override; stdcall;
    procedure Finalize; override; stdcall;
  end;

  TUMPlaylistListener = class(TInterfacedObject, IAIMPPlaylistListener)
  public
    procedure Activated; stdcall;
    procedure Changed(Flags: DWORD); stdcall;
    procedure Removed; stdcall;
  end;

  TUMPlaylistManagerListener = class(TInterfacedObject, IAIMPExtensionPlaylistManagerListener)
  protected
    UMPlaylistListener: IAIMPPlaylistListener;
    UMAttachedPlaylist: IAIMPPlaylist;
  public
    procedure PlaylistActivated(Playlist: IAIMPPlaylist); stdcall;
    procedure PlaylistAdded(Playlist: IAIMPPlaylist); stdcall;
    procedure PlaylistRemoved(Playlist: IAIMPPlaylist); stdcall;
    destructor Destroy; override;
  end;

  TUMExecuteHandler = class(TInterfacedObject, IAIMPActionEvent)
  public
    procedure OnExecute(Data: IInterface); stdcall;
  end;

{--------------------------------------------------------------------}
    TExIndex = record
      Index : Integer;
      RealIndex : Real;
    end;

    TAuthorSong = record
      Name : String;
      Songs : array of TExIndex;
    end;

    TUltraMixer = class(TObject)
     protected
      procedure SortAuthors(var Arr: array of TAuthorSong; Low, High: Integer);
      procedure RandomizeList(var Arr: array of TAuthorSong);
     public
      procedure Execute(Playlist: IAIMPPlaylist);
    end;
{--------------------------------------------------------------------}

implementation

{--------------------------------------------------------------------}
procedure ShowErrorMessage(ErrorMessage: String);
var
  DLLName: array[0..MAX_PATH - 1] of Char;
  FullMessage: String;
begin
  FillChar(DLLName, MAX_PATH, #0);
  GetModuleFileName(HInstance, DLLName, MAX_PATH);
  FullMessage := 'Exception in module "' + DLLName + '".'#13#13 + ErrorMessage;
  MessageBox(0, PChar(FullMessage), UM_CAPTION, MB_ICONERROR);
end;
{--------------------------------------------------------------------}
procedure UpdateActionStatusForActivePlaylist;
var
  PLManager: IAIMPServicePlaylistManager;
  ActivePL: IAIMPPlaylist;
  PLProperties: IAIMPPropertyList;
  PLIsReadOnly: Integer;
  ActionManager: IAIMPServiceActionManager;
  Action: IAIMPAction;
begin
 try
  // Get active playlist
  CheckResult(CoreIntf.QueryInterface(IID_IAIMPServicePlaylistManager, PLManager));
  CheckResult(PLManager.GetActivePlaylist(ActivePL));

  // Check "Read only" property
  CheckResult(ActivePL.QueryInterface(IID_IAIMPPropertyList, PLProperties));
  CheckResult(PLProperties.GetValueAsInt32(AIMP_PLAYLIST_PROPID_READONLY, PLIsReadOnly));
  PLIsReadOnly := Integer(not Bool(PLIsReadOnly));

  // Update Action status for the Playlist
  CheckResult(CoreIntf.QueryInterface(IID_IAIMPServiceActionManager, ActionManager));
  CheckResult(ActionManager.GetByID(MakeString(UM_HOTKEY_ID), Action));
  CheckResult(Action.SetValueAsInt32(AIMP_ACTION_PROPID_ENABLED, PLIsReadOnly));
 except
  ShowErrorMessage('"UpdateActionStatusForActivePlaylist" failure!');
 end;
end;

{=========================================================================)
                                 TUMPlugin
(=========================================================================}
function TUMPlugin.InfoGet(Index: Integer): PWideChar;
begin
  case Index of
    AIMP_PLUGIN_INFO_NAME               : Result := UM_PLUGIN_NAME;
    AIMP_PLUGIN_INFO_AUTHOR             : Result := UM_PLUGIN_AUTHOR;
    AIMP_PLUGIN_INFO_SHORT_DESCRIPTION  : Result := UM_PLUGIN_SHORT_DESCRIPTION;
    AIMP_PLUGIN_INFO_FULL_DESCRIPTION   : Result := UM_PLUGIN_FULL_DESCRIPTION;
  else
    Result := nil;
  end;
end;

function TUMPlugin.InfoGetCategories: Cardinal;
begin
  Result := AIMP_PLUGIN_CATEGORY_ADDONS;
end;
{--------------------------------------------------------------------
Initialize}
function TUMPlugin.Initialize(Core: IAIMPCore): HRESULT;
var
  APlaylistManager: IAIMPServicePlaylistManager;
  UMServiceMessageDispatcher: IAIMPServiceMessageDispatcher;
begin
  Result := Core.QueryInterface(IID_IAIMPServicePlaylistManager, APlaylistManager);
  if Succeeded(Result)
  then
    begin
      Result := inherited Initialize(Core);
      if Succeeded(Result)
      then
        try
          CreateActionAndMenu;
          // Creating the message hook
          CheckResult(CoreIntf.QueryInterface(IID_IAIMPServiceMessageDispatcher,
                                                UMServiceMessageDispatcher));
          UMMessageHook := TUMMessageHook.Create;
          CheckResult(UMServiceMessageDispatcher.Hook(UMMessageHook));
        except
          Result := E_UNEXPECTED;
        end;
    end;
end;
{--------------------------------------------------------------------
Finalize}
procedure TUMPlugin.Finalize;
var
  UMServiceMessageDispatcher: IAIMPServiceMessageDispatcher;
begin
 try
  // Removing the message hook
  CheckResult(CoreIntf.QueryInterface(IID_IAIMPServiceMessageDispatcher,
                                                UMServiceMessageDispatcher));
  CheckResult(UMServiceMessageDispatcher.Unhook(UMMessageHook));
 except
  ShowErrorMessage('"Plugin.Finalize" failure!');
 end;
  inherited;
end;
{--------------------------------------------------------------------}
function TUMPlugin.MakeLocalDefaultHotkey: Integer;
var
  ServiceActionManager: IAIMPServiceActionManager;
begin
  CheckResult(CoreIntf.QueryInterface(IID_IAIMPServiceActionManager,
                                                    ServiceActionManager));
  Result := ServiceActionManager.MakeHotkey(UM_HOTKEY_DEFAULT_MOD,
                                                    UM_HOTKEY_DEFAULT_KEY);
end;

function TUMPlugin.LoadMenuIcon(const ResName: string): IAIMPImage;
var
  AContainer: IAIMPImageContainer;
  AResStream: TResourceStream;
begin
  CheckResult(CoreIntf.CreateObject(IID_IAIMPImageContainer, AContainer));
  AResStream := TResourceStream.Create(HInstance, ResName, RT_RCDATA);
  try
    CheckResult(AContainer.SetDataSize(AResStream.Size));
    AResStream.ReadBuffer(AContainer.GetData^, AContainer.GetDataSize);
    CheckResult(AContainer.CreateImage(Result));
  finally
    AResStream.Free;
  end;
end;

function TUMPlugin.GetBuiltInMenu(ID: Integer): IAIMPMenuItem;
var
  AMenuService: IAIMPServiceMenuManager;
begin
  CheckResult(CoreIntf.QueryInterface(IAIMPServiceMenuManager, AMenuService));
  CheckResult(AMenuService.GetBuiltIn(ID, Result));
end;
{--------------------------------------------------------------------
CreateActionAndMenu}
procedure TUMPlugin.CreateActionAndMenu;
var
  UMHotkey: IAIMPAction;
  UMContextMenu: IAIMPMenuItem;
begin
 try
  // Create hotkey action
  CheckResult(CoreIntf.CreateObject(IID_IAIMPAction, UMHotkey));
  CheckResult(UMHotkey.SetValueAsObject(AIMP_ACTION_PROPID_ID,
                                          MakeString(UM_HOTKEY_ID)));
  CheckResult(UMHotkey.SetValueAsObject(AIMP_ACTION_PROPID_NAME,
                                          MakeString(UM_CAPTION)));
  CheckResult(UMHotkey.SetValueAsInt32(AIMP_ACTION_PROPID_DEFAULTLOCALHOTKEY,
                                          MakeLocalDefaultHotkey));
  // Geting localized string for the HOTKEY_GROUPNAME
  CheckResult(UMHotkey.SetValueAsObject(AIMP_ACTION_PROPID_GROUPNAME,
                      MakeString(LangLoadString(UM_HOTKEY_GROUPNAME_KEYPATH))));
  CheckResult(UMHotkey.SetValueAsObject(AIMP_ACTION_PROPID_EVENT,
                                          TUMExecuteHandler.Create));
  // Register the local hotkey in manager
  CheckResult(CoreIntf.RegisterExtension(IID_IAIMPServiceActionManager, UMHotkey));

  // Create menu item
  CheckResult(CoreIntf.CreateObject(IID_IAIMPMenuItem, UMContextMenu));
  CheckResult(UMContextMenu.SetValueAsObject(AIMP_MENUITEM_PROPID_ID,
                                          MakeString(UM_CONTEXTMENU_ID)));
  CheckResult(UMContextMenu.SetValueAsObject(AIMP_MENUITEM_PROPID_ACTION, UMHotkey));
  CheckResult(UMContextMenu.SetValueAsInt32(AIMP_MENUITEM_PROPID_STYLE,
                                          AIMP_MENUITEM_STYLE_NORMAL));
  CheckResult(UMContextMenu.SetValueAsObject(AIMP_MENUITEM_PROPID_PARENT,
                           GetBuiltInMenu(AIMP_MENUID_PLAYER_PLAYLIST_SORTING)));
  CheckResult(UMContextMenu.SetValueAsObject(AIMP_MENUITEM_PROPID_GLYPH,
                                          LoadMenuIcon(UM_CONTEXTMENU_ICON)));
  // Register the menu item in manager
  CheckResult(CoreIntf.RegisterExtension(IID_IAIMPServiceMenuManager, UMContextMenu));

  // Register PlaylistManagerListener
  CheckResult(CoreIntf.RegisterExtension(IID_IAIMPServicePlaylistManager,
                                          TUMPlaylistManagerListener.Create));
 except
  ShowErrorMessage('"CreateActionAndMenu" failure!');
 end;
end;

{=========================================================================)
                                TUMMessageHook
(=========================================================================}
procedure TUMMessageHook.CoreMessage(Message: DWORD; Param1: Integer;
  Param2: Pointer; var Result: HRESULT);

var
  UMServiceActionManager: IAIMPServiceActionManager;
  UMAction: IAIMPAction;
begin
  case Message  of
    AIMP_MSG_EVENT_LANGUAGE:
      try
        // Update the name of group in hotkey settings tab
        CheckResult(CoreIntf.QueryInterface(IID_IAIMPServiceActionManager,
                                                    UMServiceActionManager));
        CheckResult(UMServiceActionManager.GetByID(MakeString(UM_HOTKEY_ID),
                                                    UMAction));
        CheckResult(UMAction.SetValueAsObject(AIMP_ACTION_PROPID_GROUPNAME,
                      MakeString(LangLoadString(UM_HOTKEY_GROUPNAME_KEYPATH))));
      except
        ShowErrorMessage('"MessageHook.CoreMessage" failure!');
      end;
  end;
end;

{=========================================================================)
                        TUMPlaylistManagerListener
(=========================================================================}
destructor TUMPlaylistManagerListener.Destroy;
begin
 try
  if UMPlaylistListener <> nil
  then
    CheckResult(UMAttachedPlaylist.ListenerRemove(UMPlaylistListener));
 except
  ShowErrorMessage('"PlaylistManagerListener.Destroy" failure!');
 end;
  inherited;
end;

procedure TUMPlaylistManagerListener.PlaylistActivated(Playlist: IAIMPPlaylist);
begin
 try
  // Register PlaylistListener
  if UMPlaylistListener = nil
  then
    UMPlaylistListener :=  TUMPlaylistListener.Create
  else
    begin
      if UMAttachedPlaylist <> nil
      then
        CheckResult(UMAttachedPlaylist.ListenerRemove(UMPlaylistListener));
    end;
  CheckResult(Playlist.ListenerAdd(UMPlaylistListener));
  UMAttachedPlaylist := Playlist;
 except
  ShowErrorMessage('"PlaylistManagerListener.PlaylistActivated" failure!');
 end;
end;

procedure TUMPlaylistManagerListener.PlaylistAdded(Playlist: IAIMPPlaylist);
begin
  //
end;

procedure TUMPlaylistManagerListener.PlaylistRemoved(Playlist: IAIMPPlaylist);
begin
  //
end;

{=========================================================================)
                             TUMPlaylistListener
(=========================================================================}
procedure TUMPlaylistListener.Activated;
begin
  UpdateActionStatusForActivePlaylist;
end;

procedure TUMPlaylistListener.Changed(Flags: DWORD);
begin
  if (AIMP_PLAYLIST_NOTIFY_READONLY and Flags) <> 0
  then
    UpdateActionStatusForActivePlaylist;
end;

procedure TUMPlaylistListener.Removed;
begin
//
end;

{=========================================================================)
                              TUMExecuteHandler
(=========================================================================}
procedure TUMExecuteHandler.OnExecute(Data: IInterface);
var
  PLManager: IAIMPServicePlaylistManager;
  ActivePL: IAIMPPlaylist;
  UMixer: TUltraMixer;
begin
 try
  CheckResult(CoreIntf.QueryInterface(IID_IAIMPServicePlaylistManager, PLManager));
  CheckResult(PLManager.GetActivePlaylist(ActivePL));

  // Mixing the active playlist
  UMixer := TUltraMixer.Create;
  try
    UMixer.Execute(ActivePL);
  finally
    UMixer.Free;
  end;
 except
  ShowErrorMessage('"ExecuteHandler.OnExecute" failure!');
 end;
end;

{=========================================================================)
                                TUltraMixer
(=========================================================================}
procedure TUltraMixer.Execute(Playlist: IAIMPPlaylist);
var
  PLPropertyList: IAIMPPropertyList;
  PLItem: IAIMPPlaylistItem;
  PLItemInfo: IAIMPFileInfo;
  PLItemAuthor: IAIMPString;
  PLItemAuthorStr: String;
  PLItemCount: Integer;
  i, j, k ,
  SortedListLength: Integer;
  RealIndex : Real;
  AuthorsList : array of TAuthorSong;
  SortedList : array of TExIndex;
begin
  // Checking the Playlist READONLY status
  CheckResult(Playlist.QueryInterface(IID_IAIMPPropertyList, PLPropertyList));
  CheckResult(PLPropertyList.GetValueAsInt32(AIMP_PLAYLIST_PROPID_READONLY, i));
  if i <> 0  then  exit;

 try
  CheckResult(Playlist.BeginUpdate);

  // Initialization of variables
  PLItemCount := Playlist.GetItemCount;
  SetLength(SortedList, 0);
  SetLength(AuthorsList, 0);

  // Filling list of autors
  for i := 0 to PLItemCount - 1
  do  begin
        CheckResult(Playlist.GetItem(i, IID_IAIMPPlaylistItem, PLItem));
        CheckResult(PLItem.GetValueAsObject(AIMP_PLAYLISTITEM_PROPID_FILEINFO,
                                              IID_IAIMPFileInfo, PLItemInfo));
        CheckResult(PLItemInfo.GetValueAsObject(AIMP_FILEINFO_PROPID_ARTIST,
                                              IID_IAIMPString, PLItemAuthor));
        PLItemAuthorStr := AnsiLowerCase(IAIMPStringToString(PLItemAuthor));

        for j := 0 to Length(AuthorsList) - 1 do
          begin
            if  PLItemAuthorStr = AuthorsList[j].Name
            then
              begin
                SetLength(AuthorsList[j].Songs, Length(AuthorsList[j].Songs)+1);
                AuthorsList[j].Songs[Length(AuthorsList[j].Songs)-1].Index := i;
                Break;
              end
            else  Continue;
          end;
        if  j > Length(AuthorsList) - 1
        then
          begin
            k := Length(AuthorsList);
            SetLength(AuthorsList, k +1);
            AuthorsList[k].Name := PLItemAuthorStr;
            SetLength(AuthorsList[k].Songs, 1);
            AuthorsList[k].Songs[0].Index := i;
          end;
      end;

  // Sorting & rardomizing
  SortAuthors(AuthorsList, 0, Length(AuthorsList)-1);
  RandomizeList(AuthorsList);

  // Main cycle of mixing
  for  i := 0 to Length(AuthorsList) - 1 do
    begin
      for j := 0  to  Length(AuthorsList[i].Songs) - 1 do
        begin
          RealIndex :=   i + j * PLItemCount / Length(AuthorsList[i].Songs);
          SortedListLength := Length(SortedList);
          k := SortedListLength - 1;
          while (k > 0) and (RealIndex <= SortedList[k].RealIndex)
          do  Dec(k);
          if  k = SortedListLength - 1
          then
            begin
              SetLength(SortedList, SortedListLength + 1);
              SortedList[SortedListLength].Index := AuthorsList[i].Songs[j].Index;
              SortedList[SortedListLength].RealIndex := RealIndex;
            end
          else
            begin
              SetLength(SortedList, SortedListLength + 1);
              Move(SortedList[k], SortedList[k + 1],
                                 (SortedListLength - k) * SizeOf(SortedList[k]));
              SortedList[k+1].Index := AuthorsList[i].Songs[j].Index;
              SortedList[k+1].RealIndex := RealIndex;
            end;
          end;
      end;

  // Moving playlist entries
  for i := 0 to PLItemCount - 1 do
    begin
      CheckResult(Playlist.GetItem(SortedList[i].Index, IID_IAIMPPlaylistItem, PLItem));
      CheckResult(PLItem.SetValueAsInt32(AIMP_PLAYLISTITEM_PROPID_INDEX, PLItemCount - 1));

      for j := i to  PLItemCount - 1  do
        begin
          if  (SortedList[j].Index > SortedList[i].Index)
            and  (SortedList[j].Index <= PLItemCount - 1 - i)
          then  SortedList[j].Index := SortedList[j].Index - 1;
        end;
      SortedList[i].Index := PLItemCount - 1;
  end;
 finally
  Playlist.EndUpdate;
 end;
end;
{--------------------------------------------------------------------}
procedure TUltraMixer.RandomizeList(var Arr: array of TAuthorSong);

  procedure SwapSongs(Auth ,Index1, Index2: Integer);
    var   Tmp: TExIndex;
  begin
    Tmp := Arr[Auth].Songs[Index1];
    Arr[Auth].Songs[Index1] := Arr[Auth].Songs[Index2];
    Arr[Auth].Songs[Index2] := Tmp;
  end;

var
  i, j, k: Integer;
  Len: Integer;
  Randomized : array of Boolean;
begin
  Randomize;
  for i := 0 to Length(Arr) - 1
  do
    begin
      Len := Length(Arr[i].Songs);
      SetLength(Randomized, 0);
      SetLength(Randomized, Len);
      for j := 0 to Len - 1
      do
        begin
          repeat
            k := Random(Len);
          until not Randomized[k];
          SwapSongs(i, j, k);
          Randomized[k] := True;
        end;
    end;
end;
{--------------------------------------------------------------------}
procedure TUltraMixer.SortAuthors(var Arr: array of TAuthorSong; Low,
  High: Integer);

    procedure Swap(Index1, Index2: Integer);
     var   Tmp: TAuthorSong;
    begin
     Tmp := Arr[Index1];
     Arr[Index1] := Arr[Index2];
     Arr[Index2] := Tmp;
    end;

var
    Mid: Integer;
    Item : TAuthorSong;
    ScanUp, ScanDown: Integer;
begin
  if High - Low <= 0
    then exit;
  if High - Low = 1
  then
    begin
      if Length(Arr[High].Songs) > Length(Arr[Low].Songs)
      then  Swap(Low, High);
      Exit;
    end;
  Mid := (High + Low) shr 1;
  Item := Arr[Mid];
  Swap(Mid, Low);
  ScanUp := Low + 1;
  ScanDown := High;
  repeat
    while (ScanUp <= ScanDown)
        and (Length(Arr[ScanUp].Songs) >= Length(Item.Songs))
      do  Inc(ScanUp);
    while (Length(Arr[ScanDown].Songs) < Length(Item.Songs))
      do  Dec(ScanDown);
    if (ScanUp < ScanDown)
      then  Swap(ScanUp, ScanDown);
  until (ScanUp >= ScanDown);
  Arr[Low] := Arr[ScanDown];
  Arr[ScanDown] := Item;
  if (Low < ScanDown - 1)
    then  SortAuthors(Arr, Low, ScanDown - 1);
  if (ScanDown + 1 < High)
    then  SortAuthors(Arr, ScanDown + 1, High);
end;

{=========================================================================)
                                  THE END
(=========================================================================}

end.
