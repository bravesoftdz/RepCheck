unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  ActiveX, Classes, ComCtrls,
  ComObj, Dialogs, Forms, jwatlhelp32, lazutf8, ShellApi, StdCtrls, SysUtils, Windows;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1:   TButton;
    Button2:   TButton;
    Button3:   TButton;
    Button4:   TButton;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    Edit1:     TEdit;
    Edit2:     TEdit;
    Memo1:     TMemo;
    ProgressBar1: TProgressBar;
    ProgressBar2: TProgressBar;
    SelectDirectoryDialog1: TSelectDirectoryDialog;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FindFileInFolder(path, ext: string);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure KillTask(TaskFileName: string);
    function fileSize(const fname: string): int64;
  private
    { private declarations }
  public
    { public declarations }
  end;

type
  TMyThread = class(TThread)
  protected
    prog: integer;
    Run:  boolean;
    Err:  boolean;
    procedure UpdateProgress;
    procedure Execute; override;
  public
  end;

type
  TClearThread = class(TThread)
  protected
    procedure Execute; override;
    function DelDir(dir: string): boolean;
  public
  end;

type
  TXLSThread = class(TThread)
  protected
    procedure Execute; override;
  public
  end;

var
  Form1: TForm1;
  AdvAdr, Params: WideString;
  WinState: word;
  StringToFind, FindCat: WideString;
  RangeFrom, RangeTo: DWORD;
  FList, OKList: TStringList;
  A:   TMyThread;
  XLS: TXLSThread;
  Cl:  TClearThread;

implementation

{$R *.lfm}


procedure TForm1.KillTask(TaskFileName: string);
  const
    PROCESS_TERMINATE = $0001;
  var
    ContinueLoop:    BOOL;
    FSnapshotHandle: THandle;
    FProcessEntry32: TProcessEntry32;
  begin
    FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
    ContinueLoop    := Process32First(FSnapshotHandle, FProcessEntry32);
    while integer(ContinueLoop) <> 0 do
      begin
      if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) =
        UpperCase(TaskFileName)) or (UpperCase(FProcessEntry32.szExeFile) =
        UpperCase(TaskFileName))) then
        TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0),
          FProcessEntry32.th32ProcessID), 0);
      ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
      end;
    CloseHandle(FSnapshotHandle);
  end;

{ TForm1 }
procedure TForm1.FindFileInFolder(path, ext: string);
  var
    SR:   TSearchRec;
    Ress: integer;
  begin
    FList := TStringList.Create;
    if path[Length(path)] <> '\' then
      path := path + '\';
    Ress   := FindFirst(path + ext, faAnyFile, SR);
    while (Ress = 0) and (Sr.Attr <> faHidden) do
      begin
      FList.Add(path + Sr.Name);
      Ress := FindNext(SR);
      end;
    SysUtils.FindClose(SR);
  end;

procedure TForm1.FormCreate(Sender: TObject);
  begin
    Edit2.Text := FormatDateTime('yyyy', Now);
    if fileSize('\\euwinkiefsv001\RetailerServices\RepCheck.exe') <>
      fileSize(Application.Exename) then
      ShowMessage('Please download new version !' + #13#10 +
        'Link: \\euwinkiefsv001\RetailerServices\RepCheck.exe');
  end;

procedure TForm1.FormResize(Sender: TObject);
  begin
    Memo1.Height := Form1.Height - Memo1.Top - 5;
    Memo1.Width  := Form1.Width - 5;
  end;

procedure TMyThread.Execute;
  var
    StartInfo: TStartupInfo;
    ProcInfo: TProcessInformation;
    CmdLine: ShortString;
    // hProcess: DWORD;
    BytesRead: DWORD;
    BytesToRead: DWORD;
    I:      DWORD;
    Buffer: PWideChar;
    OldProtect: PDWORD;
    ps:     smallint = 0;
    ps2:    smallint = 0;
    ps3:    smallint = 0;
    Buf:    string = '';
    per1, per2: ShortString;
    Res:    boolean = False;
    BufferSize: integer = 512;
  begin
    Form1.KillTask('wsp.exe');
    per1    := Form1.ComboBox1.Text + ' ' + Form1.Edit2.Text;
    per2    := Form1.ComboBox2.Text + ' ' + Form1.Edit2.Text;
    CmdLine := '"' + AdvAdr + '" ' + '"' + Params + '"';
    FillChar(StartInfo, SizeOf(StartInfo), #0);
    with StartInfo do
      begin
      cb      := SizeOf(StartInfo);
      dwFlags := STARTF_USESHOWWINDOW;
      wShowWindow := WinState;
      end;
    Run := CreateProcess(nil, PChar(string(CmdLine)), nil, nil, False,
      CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil,
      PChar(ExtractFilePath(AdvAdr)), StartInfo, ProcInfo);
    { Ожидаем завершения приложения }
    WaitForInputIdle(ProcInfo.hProcess, INFINITE);
    if (Run = True) and (ProcInfo.hProcess <> 0) then
      begin
        try
        Buffer := AllocMem(BufferSize);
          try
          I := RangeFrom;
          while (I < RangeTo) do
            begin
            BytesToRead := BufferSize;
            if (RangeTo - I) < BytesToRead then
              begin
              BytesToRead := RangeTo - I;
              end;
            prog := i;
            Synchronize(@UpdateProgress);
            // if VirtualQueryEx(ProcInfo.hProcess, Pointer(i), Mbi, SizeOf(Mbi)) <> 0 then
            // VirtualProtectEx(ProcInfo.hProcess, Pointer(i), BytesToRead, PAGE_READWRITE,
            //    OldProtect);
            ReadProcessMemory(ProcInfo.hProcess, Pointer(i), Buffer, BytesToRead,
              BytesRead);
            Buf := UpperCase(WideCharToString(Buffer));
            if (UTF8Pos('.WSV', Buf) = 0) and (UTF8Pos('#INV', Buf) = 0) and
              (UTF8Pos('#MKT', Buf) = 0) and (Buf <> '') then
              begin
              if ps = 0 then
                ps := UTF8Pos(UTF8UpperCase(StringToFind), Buf);
              if ps = 0 then
                begin
                if (UTF8Pos(UTF8UpperCase(
                  StringReplace(StringToFind, '-', '', [rfReplaceAll, rfIgnoreCase])),
                  Buf) > 0) or
                  (UTF8Pos(UTF8UpperCase(StringReplace(StringToFind,
                  '-', ' ', [rfReplaceAll, rfIgnoreCase])), Buf) > 0) then
                  ps := 1;                                             //Название сети
                end;
              if ps2 = 0 then
                ps2 := UTF8Pos(UTF8UpperCase(FindCat), Buf);   //Категория
              if ps3 = 0 then
                if (UTF8Pos(UTF8UpperCase(per1), Buf) > 0) or
                  (UTF8Pos(UTF8UpperCase(per2), Buf) > 0) then
                  ps3 := 1;                                         //Период
              end;
            if (ps > 0) and (ps2 > 0) and (ps3 > 0) then
              begin
              Res := True;
              // VirtualProtectEx(ProcInfo.hProcess, Pointer(I), BytesToRead,
              //   longword(OldProtect),
              //   OldProtect);
              Break;
              end;
            Inc(I, BytesToRead);
            end;
          finally
          Buffer := nil;
          end;
        finally
        end;
      //WaitForSingleObject(ProcInfo.hProcess, INFINITE);
      if (Res = False) and (Err = True) then
        Form1.Memo1.Lines.Add('[' + UTF8Encode(Params) + ']' +
          ' - FAIL ' + 'Назв.Сети:' + IntToStr(ps) + ':' + 'Категория:' +
          IntToStr(ps2) + ':' + 'Период:' + IntToStr(ps3));
      if Res = True then
        OKList.Add('[' + UTF8Encode(Params) + ']' + ' ----- OK ');
      Form1.ProgressBar1.Position := 0;
      { Free the Handles }
      TerminateProcess(ProcInfo.hProcess, 0);
      CloseHandle(ProcInfo.hProcess);
      CloseHandle(ProcInfo.hThread);
      Form1.KillTask('wsp.exe');
      Sleep(100);
      if (res = False) and (Err = False) then
        begin
        Err := True;
        BufferSize := 360;
        Execute;
        end;
      end;
  end;

procedure TXLSThread.Execute;
  var
    XLApp: olevariant;
    Ret:   WideString;
    Res:   boolean = False;
    ps:    smallint = 0;
    ps2:   smallint = 0;
    hit:   smallint = 0;
    i:     integer;
    lis:   boolean = False;
  begin
    Form1.KillTask('EXCEL.exe');
    CoInitialize(nil);
    XLApp := CreateOleObject('Excel.Application');
      try
      XLApp.Visible := False;
      XLApp.DisplayAlerts := False;
      XLApp.Workbooks.Open(Params);
        try
        for i := 1 to XLApp.Worksheets.Count do
          if UTF8Encode(XLApp.Worksheets[i].Name) = 'Лист1' then
            lis := True;
        if (UTF8Encode(XLApp.worksheets['WSP_TOC'].cells[3, 2].Value) =
          'ТОП анализ') and (lis = False) then
          begin
          if UTF8Encode(XLApp.worksheets['WSP_Sheet4'].cells[5, 2].Value) =
            'Ранг' then
            begin
            Ret := Utf8Encode(XLApp.worksheets['WSP_Sheet4'].Shapes.item(
              'TextBox 1').TextFrame.Characters.Text);
            Delete(Ret, 1, Pos(':', Ret) + 1);
            if UTF8UpperCase(StringReplace(StringToFind, '"', '',
              [rfReplaceAll, rfIgnoreCase])) = UTF8UpperCase(
              StringReplace(Ret, ' ', '', [rfReplaceAll, rfIgnoreCase])) then
              ps := 1;
            if UTF8Pos(UTF8UpperCase(Utf8Encode(FindCat)), UTF8UpperCase(
              UTF8Encode(XLApp.worksheets['WSP_TOC'].cells[5, 3].Value))) > 0 then
              ps2 := 1;
            end;
          end;
        if UTF8Encode(XLApp.worksheets['WSP_TOC'].cells[3, 2].Value) =
          'Хит-парад' then
          hit := 1;
        except
        Res := True;
        end;
      finally
      if (ps > 0) and (ps2 > 0) and (res = False) then
        OKList.Add('[' + UTF8Encode(Params) + ']' + ' - OK ')
      else
        begin
        if hit > 0 then
          OKList.Add('[' + UTF8Encode(Params) + ']' + ' - Хит-парад ')
        else
          Form1.Memo1.Lines.Add('[' + UTF8Encode(Params) + ']' +
            ' - FAIL ' + 'Назв.Сети:' + IntToStr(ps) + ':' + 'Категория:' +
            IntToStr(ps2));
        end;
      XLApp.ActiveWorkBook.Close;
      XLApp.Quit;
      XLAPP := Unassigned;
      end;
    CoUninitialize;
  end;

procedure TMyThread.UpdateProgress;
  begin
    Form1.ProgressBar1.Position := prog;
  end;

procedure TForm1.Button2Click(Sender: TObject);
  var
    tmp:  shortstring;
    i, r: smallint;
    iRus: smallint;
  begin
    Button3.Enabled := True;
    OKList := TStringList.Create;
    Memo1.Clear;
    WinState  := SW_HIDE;
    RangeFrom := $00000000;
    RangeTo   := $1fffffff;
    ProgressBar1.Max := integer(Pointer(RangeTo));
    AdvAdr    := 'C:\Program Files\ACNielsen\Advisor Interactive\wsp.exe';
    FindFileInFolder(Edit1.Text, '*.wsv');
    ProgressBar2.Max := FList.Count;
    for i := 0 to FList.Count - 1 do
      begin
      ////Очистка TEMP
      Cl := TClearThread.Create(True);
      Cl.Priority := tpHighest;
      Cl.FreeOnTerminate := True;
      Cl.Resume;
      /////
      iRus   := 0;
      A      := TMyThread.Create(True);
      A.Priority := tpHighest;
      A.FreeOnTerminate := True;
      Params := Flist[i];
      tmp    := Params;
      while pos('\', tmp) > 0 do
        Delete(tmp, 1, pos('\', tmp));
      for r := 1 to Length(tmp) do
        if Ord(tmp[r]) in [192..239, 240..255, 167, 183] then
          begin
          iRus := r - 2;
          break;
          end;
      if iRus < 1 then
        StringToFind := '"' + Copy(tmp, 1, Pos('_', tmp) - 1)
      else
        begin
        StringToFind := '"' + Copy(tmp, 1, iRus);
        StringToFind := StringReplace(StringToFind, '_', ' ',
          [rfReplaceAll, rfIgnoreCase]);
        end;
      while pos('_', tmp) > 0 do
        Delete(tmp, 1, pos('_', tmp));
      FindCat := Copy(tmp, 1, Pos('.', tmp) - 1);
      ProgressBar2.Position := i + 1;
      A.Resume;
      Application.ProcessMessages;
      A.WaitFor;
      if i >= FList.Count - 1 then
        ProgressBar2.Position := 0;
      end;

    //////////////EXCEL CHECK//////////////////////////////////////////////////
    FindFileInFolder(Edit1.Text, '*.xls');
    ProgressBar2.Max := FList.Count;
    for i := 0 to FList.Count - 1 do
      begin
      iRus   := 0;
      XLS    := TXLSThread.Create(True);
      XLS.Priority := tpHighest;
      XLS.FreeOnTerminate := True;
      Params := Flist[i];
      tmp    := Params;
      while pos('\', tmp) > 0 do
        Delete(tmp, 1, pos('\', tmp));
      for r := 1 to Length(tmp) do
        if Ord(tmp[r]) in [192..239, 240..255, 167, 183] then
          begin
          iRus := r - 2;
          break;
          end;
      if iRus < 1 then
        StringToFind := '"' + Copy(tmp, 1, Pos('_', tmp) - 1)
      else
        begin
        StringToFind := '"' + Copy(tmp, 1, iRus);
        StringToFind := StringReplace(StringToFind, '_', ' ',
          [rfReplaceAll, rfIgnoreCase]);
        end;
      while pos('_', tmp) > 0 do
        Delete(tmp, 1, pos('_', tmp));
      FindCat := Copy(tmp, 1, Pos('.', tmp) - 1);
      ProgressBar2.Position := i + 1;
      XLS.Resume;
      Application.ProcessMessages;
      XLS.WaitFor;
      if i >= FList.Count - 1 then
        ProgressBar2.Position := 0;
      end;
  end;

procedure TForm1.Button3Click(Sender: TObject);
  var
    i: smallint;
  begin
    Memo1.Lines.Add('');
    Memo1.Lines.Add(
      '----------------------------------SUCCESSFUL LIST------------------------------');
    for i := 0 to OKlist.Count - 1 do
      Memo1.Lines.Add(OKlist[i]);
    OKList.Free;
  end;

procedure TForm1.Button4Click(Sender: TObject);
  begin
    Cl := TClearThread.Create(True);
    Cl.Priority := tpHighest;
    Cl.FreeOnTerminate := True;
    Cl.Resume;
    Button4.Enabled := False;
  end;

function TForm1.fileSize(const fname: string): int64;
  var
    h: integer;
  begin
    h := FileOpen(fname, fmOpenRead);
    if (INVALID_HANDLE_VALUE <> DWORD(h)) then
        try
        Result := FileSeek(h, 0, 2);
        finally
        FileClose(h);
        end
    else
      Result := -1;
  end;

function TClearThread.DelDir(dir: string): boolean;
  var
    fos: TSHFileOpStruct;
  begin
    ZeroMemory(@fos, SizeOf(fos));
    with fos do
      begin
      wFunc  := FO_DELETE;
      fFlags := FOF_SILENT or FOF_NOCONFIRMATION;
      pFrom  := PChar(dir + #0);
      end;
    Result := (0 = ShFileOperation(fos));
  end;

procedure TClearThread.Execute;
  var
    SR:     TSearchRec;
    FindRes: integer;
    N:      string;
    ClCou:  integer = 0;
    NClCou: integer = 0;
  begin
    N := SysUtils.GetEnvironmentVariable('TEMP');
    findRes := FindFirst(N + '\~ACN*', faDirectory, SR);
    while findres = 0 do
      begin
      if DelDir(N + '\' + SR.Name) then
        ClCou  := ClCou + 1
      else
        NClCou := NClCou + 1;
      FindRes := FindNext(SR);
      end;
    SysUtils.FindClose(SR);
    // if ClCou > 1 then
    //   ShowMessage('Deleted [' + IntToStr(ClCou) + '] Files' + #13#10 +
    //     'Failed to Delete [' + IntToStr(NClCou) + '] files');
    Form1.Button4.Enabled := True;
  end;

procedure TForm1.Button1Click(Sender: TObject);
  begin
    if SelectDirectoryDialog1.Execute then
      begin
      Edit1.Text      := SelectDirectoryDialog1.FileName;
      Button2.Enabled := True;
      end;
  end;

end.
