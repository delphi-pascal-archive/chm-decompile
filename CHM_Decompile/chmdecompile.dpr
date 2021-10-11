////////////////////////////////////////////////////////////////////////////////
//
//  ****************************************************************************
//  * Project   : chmdecompile sample
//  * Purpose   : Демонстрационный пример извлечения данных из CHM файлов
//  * Author    : Александр (Rouse_) Багель
//  * Copyright : © Fangorn Wizards Lab 1998 - 2009.
//  * Version   : 1.00
//  * Home Page : http://rouse.drkb.ru
//  ****************************************************************************
//

program chmdecompile;

{$APPTYPE CONSOLE}

uses
  Windows,
  Classes,
  ActiveX,
  AxCtrls,
  ComObj,
  SysUtils;

const
  CLSID_ITStorage: TGUID = (D1: $5D02926A; D2: $212E; D3: $11D0;
    D4: ($9D, $F9, $00, $A0, $C9, $22, $E6, $EC));
  IID_ITStorage: TGUID = (D1: $88CC31DE; D2: $27AB; D3: $11D0;
    D4: ($9D, $F9, $00, $A0, $C9, $22, $E6, $EC));

var
  SaveFolderPathLength: Integer;

type
  TCompactionLev = (COMPACT_DATA, COMPACT_DATA_AND_PATH);

  PItsControlData = ^TItsControlData;
  _ITS_Control_Data = record
    cdwControlData: UINT;
    adwControlData: array [0..0] of UINT;
  end;
  TItsControlData=_ITS_Control_Data;

  IItsStorage = interface(IUnknown)
    ['{88CC31DE-27AB-11D0-9DF9-00A0C922E6EC}']
    function StgCreateDocFile(const pwcsName: PWideChar; grfMode: DWORD;
      reserved: DWORD; var ppstgOpen: IStorage): HRESULT; stdcall;
    function StgCreateDocFileOnILockBytes(plkbyt: ILockBytes; grfMode: DWORD;
      reserved: DWORD; var ppstgOpen: IStorage): HRESULT; stdcall;
    function StgIsStorageFile(const pwcsName:  PWideChar): HRESULT; stdcall;
    function StgIsStorageILockBytes(plkbyt: ILockBytes): HRESULT; stdcall;
    function StgOpenStorage(const pwcsName: PWideChar; pstgPriority: IStorage;
      grfMode: DWORD; snbExclude: TSNB; reserved: DWORD;
      var ppstgOpen: IStorage): HRESULT; stdcall;
    function StgOpenStorageOnILockBytes(plkbyt: ILockBytes;
      pStgPriority: IStorage; grfMode: DWORD; snbExclude: TSNB;
      reserved: DWORD; var ppstgOpen: IStorage): HRESULT; stdcall;
    function StgSetTimes(const lpszName: PWideChar;
      const pctime, patime, pmtime: TFileTime): HRESULT; stdcall;
    function SetControlData(pControlData: PItsControlData): HRESULT; stdcall;
    function DefaultControlData(
      var ppControlData: PItsControlData): HRESULT; stdcall;
    function Compact(const pwcsName: PWideChar;
      iLev: TCompactionLev): HRESULT; stdcall;
  end;

  procedure ShowHelp;
  begin
    Writeln('Use: chmdecompile.exe [path to CHM file]');
  end;

  procedure ExtractLog(IsStorage: Boolean; const Value: string);
  begin
    if IsStorage then
      Writeln('Folder created: ',
        Copy(Value, SaveFolderPathLength, Length(Value)))
    else
      Writeln('Stream extracted: ',
        Copy(Value, SaveFolderPathLength, Length(Value)))
  end;

  function WaitKeyboardInput(const Promt: string): Char;
  var
    Done: Boolean;
    IR: INPUT_RECORD;
    hCon: THandle;
    NumOfEvents,
    NumOfEventsRead: DWORD;
    I: Integer;
  begin
    Writeln;
    Writeln(Promt);
    Result := #0;
    Done := False;
    hCon := GetStdHandle(STD_INPUT_HANDLE);
    try
      while not Done do
      begin
         if not GetNumberOfConsoleInputEvents(hCon, NumOfEvents) then
           raise EInOutError.CreateFmt(
            'GetNumberOfConsoleInputEvents failed %s',
            [SysErrorMessage(GetLastError)]);
         if NumOfEvents <= 0 then Continue;
         for I := 0 to NumOfEvents - 1 do
         begin
           if (not ReadConsoleInput(hCon, ir, 1, NumOfEventsRead)) then
             raise EInOutError.CreateFmt(
              'ReadConsoleInput failed %s', [SysErrorMessage(GetLastError)]);
           Done :=
            (NumOfEventsRead = 1) and
            (IR.EventType = KEY_EVENT) and
            TKeyEventRecord(IR.Event).bKeyDown;
           if Done then
             Result := TKeyEventRecord(IR.Event).AsciiChar;
         end;
      end;
    except
      on E : Exception do
       Writeln(Format('Exception: %s', [E.Message]));
    end;
  end;

  procedure ClearScreen;
  var
    ActualCoord, ZeroCoord: TCoord;
    cWritten: DWORD;
    hStdout: THandle;
    chFillChar: Char;
  begin
    hStdout := GetStdHandle(STD_OUTPUT_HANDLE);
    ActualCoord := GetLargestConsoleWindowSize(hStdout);
    ZeroMemory(@ZeroCoord, SizeOf(TCoord));
    chFillChar := ' ';
    FillConsoleOutputCharacter(hStdout, chFillChar,
      ActualCoord.X * ActualCoord.Y, ZeroCoord, cWritten);
    SetConsoleCursorPosition(hStdout, ZeroCoord);
  end;

  function ValidStorage(const Path: string): Boolean;
  var
    ITS: IItsStorage;
  begin
    Result := False;
    OleCheck(CoCreateInstance(CLSID_ITStorage, nil,
      CLSCTX_INPROC_SERVER, IID_ITStorage, ITS));
    if FileExists(Path) then
      Result := ITS.StgIsStorageFile(StringToOleStr(Path)) = S_OK;
  end;

  procedure ExtractStream(const RootPath: string; Root: IStorage;
    const StreamName: string);
  var
    TmpStream: IStream;
    OS: TOleStream;
    FS: TFileStream;
    FilePath: string;
  begin
    OleCheck(Root.OpenStream(StringToOleStr(StreamName),
      nil, STGM_READ or STGM_SHARE_EXCLUSIVE, 0, TmpStream));
    OS := TOleStream.Create(TmpStream);
    try
      FilePath := IncludeTrailingPathDelimiter(RootPath) + StreamName;
      FS := TFileStream.Create(FilePath, fmCreate);
      try
        OS.Position := 0;
        FS.CopyFrom(OS, OS.Size);
        ExtractLog(False, FilePath);
      finally
        FS.Free;
      end;
    finally
      OS.Free;
    end;
  end;

  procedure ExtractFolder(const RootPath: string; Root: IStorage);
  var
    ShellMalloc: IMalloc;
    Enum: IEnumStatStg;
    Fetched: Int64;
    TmpElement: TStatStg;
    ChildFolder: IStorage;
    ChildPath: string;
  begin
    if (CoGetMalloc(1, ShellMalloc) <> S_OK) or (ShellMalloc = nil) then
      raise EComponentError.Create('CoGetMalloc failed.');
    OleCheck(Root.EnumElements(0, nil, 0, Enum));
    Fetched := 1;
    while Fetched > 0 do
      if Enum.Next(1, TmpElement, @Fetched) = S_OK then
        if ShellMalloc.DidAlloc(TmpElement.pwcsName) = 1 then
        try
          if TmpElement.dwType = STGTY_STORAGE then
          begin
            OleCheck(Root.OpenStorage(TmpElement.pwcsName, nil,
              STGM_READ or STGM_SHARE_EXCLUSIVE, nil, 0, ChildFolder));
            ChildPath := IncludeTrailingPathDelimiter(RootPath) +
              string(TmpElement.pwcsName);
            ForceDirectories(ChildPath);
            ExtractLog(True, ChildPath);
            ExtractFolder(ChildPath, ChildFolder);
          end
          else
          begin
            ExtractStream(RootPath, Root, string(TmpElement.pwcsName));
          end;
        finally
          ShellMalloc.Free(TmpElement.pwcsName);
        end;
  end;

  procedure Decompile(const Path: string);
  var
    ITS: IItsStorage;
    Root: IStorage;
    SaveFolderPath: string;
  begin
    OleCheck(CoCreateInstance(CLSID_ITStorage, nil,
      CLSCTX_INPROC_SERVER, IID_ITStorage, ITS));
    OleCheck(ITS.StgOpenStorage(StringToOleStr(Path), nil,
      STGM_READ or STGM_SHARE_EXCLUSIVE, nil, 0, Root));
    SaveFolderPath := ExtractFileName(Path);
    SaveFolderPath := Copy(SaveFolderPath, 1, Length(SaveFolderPath) -
      Length(ExtractFileExt(Path)));
    SaveFolderPath := ExtractFilePath(Path) +
      UpperCase(SaveFolderPath) + '_DUMP';
    ForceDirectories(SaveFolderPath);
    Writeln('Create dump foledr: ', SaveFolderPath);
    SaveFolderPathLength := Length(SaveFolderPath) + 1;
    ExtractFolder(SaveFolderPath, Root);
    Writeln;
    Writeln('All jobs done');
  end;

begin
  try
    ClearScreen;
    CoInitialize(nil);
    if ParamCount = 0 then
      ShowHelp
    else
    begin
      if ValidStorage(ParamStr(1)) then
        Decompile(ParamStr(1))
      else
        Writeln(ExtractFileName(ParamStr(1)), ' has wrong format.');
    end;
  except
    on E: Exception do
      Writeln(E.Message);
  end;   
  WaitKeyboardInput('Press any key to continue...')
end.
