unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XPMan, ComCtrls, Vcl.ExtCtrls, Vcl.Buttons, System.Actions,
  Vcl.ActnList, Vcl.Ribbon, Vcl.ActnCtrls, Vcl.ToolWin, Vcl.ActnMan,
  Vcl.ActnMenus, Vcl.RibbonActnMenus, Vcl.RibbonLunaStyleActnCtrls,
  Vcl.RibbonObsidianStyleActnCtrls, Vcl.ImgList, Vcl.Menus, System.ImageList;

type
  TMode = (mdGroupRen, mdFixListRen);
  TFormMain = class(TForm)
    OpenDialogFiles: TOpenDialog;
    ProgressBar1: TProgressBar;
    ImageList16: TImageList;
    ImageList32: TImageList;
    ActionList: TActionList;
    ActionAddFiles: TAction;
    ActionAddFilesDir: TAction;
    ActionAddHiddenSkip: TAction;
    ActionAddRecDir: TAction;
    ActionMGroupRen: TAction;
    ActionMFixListRen: TAction;
    ActionCLRList: TAction;
    ActionStopRen: TAction;
    ActionExit: TAction;
    GridPanelFL: TGridPanel;
    MemoFiles: TMemo;
    Panel2: TPanel;
    Label6: TLabel;
    ButtonPrefDir: TSpeedButton;
    EditPrefDir: TEdit;
    ListViewFiles: TListView;
    PanelMode: TPanel;
    PageControl: TPageControl;
    TabSheetModeNum: TTabSheet;
    ButtonRename: TButton;
    ListBoxPrev: TListBox;
    LabelSample: TLabel;
    LabelPrefix: TLabel;
    LabelInc: TLabel;
    LabelCharNum: TLabel;
    Label5: TLabel;
    Label4: TLabel;
    Label3: TLabel;
    Label2: TLabel;
    Label1: TLabel;
    EditPrefix: TEdit;
    EditExt: TEdit;
    EditCharNum: TEdit;
    EditBegin: TEdit;
    ComboBoxInc: TComboBox;
    CheckBoxFN: TCheckBox;
    CheckBoxEF: TCheckBox;
    Bevel6: TBevel;
    Bevel5: TBevel;
    Bevel4: TBevel;
    Bevel3: TBevel;
    Bevel2: TBevel;
    Bevel1: TBevel;
    ActionModeNum: TAction;
    ActionModeFNNormal: TAction;
    TabSheetModeFNNormal: TTabSheet;
    ButtonRenameNorm: TButton;
    Bevel7: TBevel;
    Bevel8: TBevel;
    CheckBoxCharRep: TCheckBox;
    EditSpaceChar: TEdit;
    GroupBox1: TGroupBox;
    RadioButton1: TRadioButton;
    RadioButtonNameDesc: TRadioButton;
    GroupBox2: TGroupBox;
    RadioButtonCharFirstUp: TRadioButton;
    MainMenu: TMainMenu;
    MenuItemFile: TMenuItem;
    MenuItemExit: TMenuItem;
    MenuItemAdd: TMenuItem;
    MenuItemAddFiles1: TMenuItem;
    MenuItemAddFilesDir: TMenuItem;
    MenuItemN1: TMenuItem;
    MenuItemAddRecDir: TMenuItem;
    MenuItemAddHiddenSkip: TMenuItem;
    MenuItemList: TMenuItem;
    MenuItemMGroupRen: TMenuItem;
    MenuItemMFixListRen: TMenuItem;
    MenuItemN2: TMenuItem;
    MenuItemCLRList: TMenuItem;
    MenuItemMode: TMenuItem;
    MenuItemModeNum: TMenuItem;
    MenuItemModeFNNormal: TMenuItem;
    MenuItemStopRen: TMenuItem;
    CheckBoxCharEach: TCheckBox;
    RadioButtonCharDown: TRadioButton;
    RadioButton2: TRadioButton;
    CheckBoxRepChars: TCheckBox;
    EditCharRepTo: TEdit;
    EditCharRepTarget: TEdit;
    Label8: TLabel;
    procedure ButtonRefreshClick(Sender: TObject);
    procedure ButtonRenameClick(Sender: TObject);
    procedure ButtonDeleteClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ListViewFilesMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ButtonPrefDirClick(Sender: TObject);
    procedure ActionStopRenExecute(Sender: TObject);
    procedure ActionCLRListExecute(Sender: TObject);
    procedure ActionMFixListRenExecute(Sender: TObject);
    procedure ActionMGroupRenExecute(Sender: TObject);
    procedure ActionAddRecDirExecute(Sender: TObject);
    procedure ActionAddHiddenSkipExecute(Sender: TObject);
    procedure ActionAddFilesDirExecute(Sender: TObject);
    procedure ActionAddFilesExecute(Sender: TObject);
    procedure ActionExitExecute(Sender: TObject);
    procedure ActionModeNumExecute(Sender: TObject);
    procedure ActionModeFNNormalExecute(Sender: TObject);
    procedure ButtonRenameNormClick(Sender: TObject);
  private
    originalLVWindowProc:TWndMethod;
    FMode:TMode;
    procedure LVWindowProc(var Msg: TMessage);
    procedure LVImageDrop(var Msg: TMessage);
  public
    procedure Proc;
    procedure StopProc;
    procedure ChangeMode(aMode:TMode);
  end;

var
  FormMain: TFormMain;
  Stop:Boolean = False;

  function GetName(Numb:Word; Move:Integer):string; overload;
  procedure AddFile(LV:TListView; FN:string);

implementation
 uses ShlObj, ActiveX, ShellAPI, System.Win.ComObj;

{$R *.dfm}

//Перевернуть строку
function Reverse(s:string):string;
var i:Word;
begin
 if Length(s) <= 1 then Exit(s);
 for i:= Length(s) downto 1 do Result:=Result+s[i];
end;

//Имя файла без разширения
function GetFileNameWoE(FileName:TFileName):string;
var PPos:Integer;
    str:string;
begin
 str:=ExtractFileName(FileName);
 if Length(str) < 3 then Exit;
 PPos:=Pos('.', Reverse(str));
 if PPos > 0 then Result:=Copy(str, 1, Length(str) - PPos);
end;

function GetFileDescription(const FileName, ExceptText:string):string;
type TLangRec = array[0..1] of Word;
var
  InfoSize, zero:Cardinal;
  pbuff: Pointer;
  pk: Pointer;
  nk: Cardinal;
  lang_hex_str: String;
  LangID:Word;
  LangCP:Word;
begin
 pbuff:=nil;
 Result:='';
 InfoSize:=GetFileVersionInfoSize(PChar(FileName), zero);
 if InfoSize <> 0 then
  try
   GetMem(pbuff, InfoSize);
   if GetFileVersionInfo(PChar(FileName), 0, InfoSize, pbuff) then
    begin
     if VerQueryValue(pbuff, '\VarFileInfo\Translation', pk, nk) then
      begin
       LangID:= TLangRec(pk^)[0];
       LangCP:= TLangRec(pk^)[1];
       lang_hex_str:= Format('%.4x',[LangID]) + Format('%.4x', [LangCP]);  //FileDescription
       if VerQueryValue(pbuff, PChar('\\StringFileInfo\\'+lang_hex_str+'\\FileDescription'), pk, nk)
       then Result:=String(PChar(pk))
       else
        if VerQueryValue(pbuff, PChar('\\StringFileInfo\\'+lang_hex_str+'\\CompanyName'), pk, nk)
        then Result:=String(PChar(pk));
      end;
    end;
  finally
   if pbuff <> nil then FreeMem(pbuff);
  end;
 if Result <> '' then
  while (Length(Result) > 0) and (Result[Length(Result)] = ' ') do Delete(Result, Length(Result), 1);
 if Result <> '' then
  while (Length(Result) > 0) and (Result[1] = ' ') do Delete(Result, 1, 1);
 if Result = '' then
  begin
   if (ExceptText <> '') then if (ExceptText <> '/') then Result:=ExceptText else Exit('')
   else Result:=GetFileNameWoE(FileName);
  end;
end;

function AdvSelectDirectory(const Caption: string; const Root: WideString;
 var Directory: string; EditBox: Boolean = False; ShowFiles: Boolean = False;
 AllowCreateDirs: Boolean = True): Boolean;

function SelectDirCB(Wnd: HWND; uMsg: UINT; lParam, lpData: lParam): Integer; stdcall;
begin
 case uMsg of
   BFFM_INITIALIZED: SendMessage(Wnd, BFFM_SETSELECTION, Ord(True), Integer(lpData));
 end;
 Result:= 0;
end;

var
 WindowList: Pointer;
 BrowseInfo: TBrowseInfo;
 Buffer: PChar;
 RootItemIDList, ItemIDList: PItemIDList;
 ShellMalloc: IMalloc;
 IDesktopFolder: IShellFolder;
 Eaten, Flags: LongWord;
const
 BIF_USENEWUI = $0040;
 BIF_NOCREATEDIRS = $0200;
begin
 Result:= False;
 if not DirectoryExists(Directory) then Directory:= '';
 FillChar(BrowseInfo, SizeOf(BrowseInfo), 0);
 if (ShGetMalloc(ShellMalloc) = S_OK) and (ShellMalloc <> nil) then
  begin
   Buffer:= ShellMalloc.Alloc(MAX_PATH);
   try
    RootItemIDList:= nil;
    if Root <> '' then
     begin
      SHGetDesktopFolder(IDesktopFolder);
      IDesktopFolder.ParseDisplayName(Application.Handle, nil, POleStr(Root), Eaten, RootItemIDList, Flags);
     end;
    OleInitialize(nil);
    with BrowseInfo do
     begin
      hwndOwner:= Application.Handle;
      pidlRoot:= RootItemIDList;
      pszDisplayName := Buffer;
      lpszTitle:= PChar(Caption);
      ulFlags:= BIF_RETURNONLYFSDIRS or BIF_USENEWUI or
         BIF_EDITBOX * Ord(EditBox) or BIF_BROWSEINCLUDEFILES * Ord(ShowFiles) or
         BIF_NOCREATEDIRS * Ord(not AllowCreateDirs);
      lpfn:= @SelectDirCB;
      if Directory <> '' then lParam := Integer(PChar(Directory));
     end;
    WindowList:= DisableTaskWindows(0);
    try
     ItemIDList:= ShBrowseForFolder(BrowseInfo);
    finally
     EnableTaskWindows(WindowList);
    end;
    Result:= ItemIDList <> nil;
    if Result then
     begin
      ShGetPathFromIDList(ItemIDList, Buffer);
      ShellMalloc.Free(ItemIDList);
      Directory:= Buffer;
     end;
   finally
    ShellMalloc.Free(Buffer);
   end;
  end;
end;

procedure TFormMain.ChangeMode(aMode:TMode);
begin
 FMode:=aMode;
 case FMode of
  mdGroupRen:
   begin
    ListViewFiles.Visible:=True;
    GridPanelFL.Visible:=False;
    ActionMGroupRen.Checked:=True;
    ActionMFixListRen.Checked:=False;
   end;
  mdFixListRen:
   begin
    GridPanelFL.Visible:=True;
    ListViewFiles.Visible:=False;
    ActionMGroupRen.Checked:=False;
    ActionMFixListRen.Checked:=True;
   end;
 end;
end;

procedure TFormMain.Proc;
begin
 Stop:=False;
 ProgressBar1.Visible:=True;
 ActionStopRen.Enabled:=True;
end;

procedure TFormMain.StopProc;
begin
 Stop:=True;
 ProgressBar1.Visible:=False;
 ActionStopRen.Enabled:=False;
end;

procedure Scan(SDir:string);
var SearchRec:TSearchRec;
begin
 Application.ProcessMessages;
 if Stop then Exit;
 if SDir[Length(SDir)] <> '\' then SDir:=SDir + '\';
 if FindFirst(SDir + '*.*', faAnyFile, SearchRec) = 0 then
  begin
   with FormMain do
    repeat
     Application.ProcessMessages;
     if Stop then Break;
     if (SearchRec.Attr and faDirectory) <> faDirectory then
      begin
       if (not ActionAddHiddenSkip.Checked) or
          ((SearchRec.Attr and faHidden) <> faHidden) or
          ((SearchRec.Attr and faSysFile) <> faSysFile)
       then
        begin
         AddFile(FormMain.ListViewFiles, SDir + SearchRec.Name);
        end
      end
     else
      if (SearchRec.Name <> '..') and (SearchRec.Name <> '.')then
       begin
        if ActionAddRecDir.Checked then Scan(SDir + SearchRec.Name + '\');
       end;
    until FindNext(SearchRec) <> 0;
   FindClose(SearchRec);
  end;
end;

procedure TFormMain.LVWindowProc(var Msg: TMessage) ;
begin
 if Msg.Msg = WM_DROPFILES then LVImageDrop(Msg) else originalLVWindowProc(Msg);
end;

procedure ShowSysPopup(aFile: string; x, y: integer; HND: HWND);
var Root: IShellFolder;
    ShellParentFolder: IShellFolder;
    chEaten,dwAttributes: ULONG;
    FilePIDL,ParentFolderPIDL: PItemIDList;
    CM: IContextMenu;
    Menu: HMenu;
    Command: LongBool;
    ICM2: IContextMenu2;
    ICI: TCMInvokeCommandInfo;
    ICmd: integer;
    P: TPoint;
begin
 OleCheck(SHGetDesktopFolder(Root));//Get the Desktop IShellFolder interface
 OleCheck(Root.ParseDisplayName(HND, nil,
    PWideChar(WideString(ExtractFilePath(aFile))),
    chEaten, ParentFolderPIDL, dwAttributes)); // Get the PItemIDList of the parent folder

 OleCheck(Root.BindToObject(ParentFolderPIDL, nil, IShellFolder,
    ShellParentFolder)); // Get the IShellFolder Interface  of the Parent Folder

 OleCheck(ShellParentFolder.ParseDisplayName(HND, nil,
    PWideChar(WideString(ExtractFileName(aFile))),
    chEaten, FilePIDL, dwAttributes)); // Get the relative  PItemIDList of the File

 ShellParentFolder.GetUIObjectOf(HND, 1, FilePIDL, IID_IContextMenu, nil, CM); // get the IContextMenu Interace for the file

 if CM = nil then Exit;
 P.X := X;
 P.Y := Y;

 //Windows.ClientToScreen(HND, P);
 Menu:= CreatePopupMenu;

 try
  CM.QueryContextMenu(Menu, 0, 1, $7FFF, CMF_EXPLORE or CMF_EXTENDEDVERBS);
  CM.QueryInterface(IID_IContextMenu2, ICM2); //To handle submenus.
  try
   Command:= TrackPopupMenu(Menu, TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RIGHTBUTTON or
    TPM_RETURNCMD, p.X, p.Y, 0, HND, nil);
  finally
   ICM2:= nil;
  end;

  if Command then
   begin
    ICmd:= LongInt(Command) - 1;
    FillChar(ICI, SizeOf(ICI), #0);
    with ICI do
     begin
      cbSize:= SizeOf(ICI);
      hWND:= 0;
      lpVerb:= MakeIntResourceA(ICmd);
      nShow:= SW_SHOWNORMAL;
     end;
    CM.InvokeCommand(ICI);
   end;
 finally
  DestroyMenu(Menu)
 end;
End;

procedure TFormMain.ListViewFilesMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
 if Button = mbRight then
  begin
   if ListViewFiles.Selected = nil then Exit;
   ShowSysPopup(ListViewFiles.Selected.SubItems[0], Mouse.CursorPos.X, Mouse.CursorPos.Y, Handle);
  end;
end;

procedure TFormMain.LVImageDrop(var Msg: TMessage) ;
var i, amount, size: integer;
    Filename:PChar;
    FN:string;
begin
 Proc;
 inherited;
 Amount:=DragQueryFile(Msg.WParam, $FFFFFFFF, Filename, 255);
 for i:= 0 to (Amount - 1) do
  begin
   size:=DragQueryFile(Msg.WParam, i, nil, 0) + 1;
   Filename:=StrAlloc(size);
   DragQueryFile(Msg.WParam, i, Filename, size);
   FN:=StrPas(Filename);
   if FileExists(FN) then AddFile(ListViewFiles, FN);
   if DirectoryExists(FN) then Scan(FN);
   StrDispose(Filename);
  end;
 DragFinish(Msg.WParam);
 StopProc;
end;

function GetExt:string;
begin
 with FormMain do
  if EditExt.Text <> '' then Result:='.'+EditExt.Text else Result:='';
end;

function GetMove:Integer;
begin
 with FormMain do
  if EditBegin.Text <> '' then
   TryStrToInt(EditBegin.Text, Result)
  else Result:=0;
end;

function GetNums:Integer;
begin
 with FormMain do
  if EditCharNum.Text <> '' then
   TryStrToInt(EditCharNum.Text, Result)
  else Result:=0;
end;

function GetName(Numb:Word; Move:Integer):string; overload;
var Num:string;
    Nums:Word;
begin
 Randomize;
 with FormMain do
  begin
   try
    Nums:=GetNums;
   except
    raise Exception.Create('Ошибка преобразования значения в поле "Кол-во знаков"');
   end;
   case ComboBoxInc.ItemIndex of
    0:Numb:=Numb+Move;
    1:Numb:=Numb*2+Move;
   end;
   Num:=IntToStr(Numb);
   while Length(Num) < Nums do Num:='0'+Num;
   Result:=EditPrefix.Text+Num;
  end;
end;

procedure TFormMain.ButtonRefreshClick(Sender: TObject);
var i:Byte;
    Num:Integer;
begin
 ListBoxPrev.Clear;
 Num:=GetMove;
 for i:= 1 to 5 do
  begin
   ListBoxPrev.Items.Add(ChangeFileExt(GetName(i-1, Num), GetExt));
  end;
end;

procedure TFormMain.ButtonPrefDirClick(Sender: TObject);
var dir:string;
begin
 if AdvSelectDirectory('Выберите каталог...', '', dir, False, False, False) then
  begin
   EditPrefDir.Text:=dir;
  end;
end;

procedure AddFile(LV:TListView; FN:string);
var ListItem:TListItem;
 i:Word;
 Exists:Boolean;
begin
 Exists:=False;
 if LV.Items.Count>0 then
  for i:=0 to LV.Items.Count-1 do
   if FN = LV.Items.Item[i].SubItems[0] then
    begin
     Exists:=True;
     Break;
    end;
 if not Exists then
  with LV.Items do
   begin
    ListItem:=Add;
    ListItem.Caption:=ExtractFileName(FN);
    ListItem.SubItems.Add(FN);
    ListItem.SubItems.Add('Ожидание');
   end;
end;

function BTS(Value:Boolean):string;
begin
 if Value then Exit('ОК') else Exit('Не ОК');
end;

procedure TFormMain.ButtonRenameClick(Sender: TObject);
var i, j, Num:Word;
  TName, CName, DName:string;
  ResRen:Boolean;
begin
 if ListViewFiles.Visible then
  begin
   if ListViewFiles.Items.Count <= 0 then
    begin
     MessageBox(Handle, 'Список файлов пуст!', '', MB_ICONWARNING or MB_OK);
     Exit;
    end;
   if MessageBox(Handle, 'Вы действительно хотите выполнить операцию переименования?', '', MB_ICONWARNING or MB_YESNO)<>ID_YES then Exit;
   Proc;
   //Начинаем счет с нужного значения
   Num:=GetNums;
   //Идём по таблице
   for i:=0 to ListViewFiles.Items.Count-1 do
    begin
     //Текущее имя файла
     CName:=ListViewFiles.Items.Item[i].SubItems[0];
     //Новое имя - пока текущее
     TName:=CName;
     //Изменить имя
     if CheckBoxFN.Checked then TName:=ExtractFilePath(CName)+GetName(i, Num)+ExtractFileExt(CName);
     //Изменить расширение
     if CheckBoxEF.Checked then TName:=ChangeFileExt(TName, GetExt);
     //Переименновываем
     ResRen:=RenameFile(CName, TName);
     //Покажем успех переименнования в таблице
     ListViewFiles.Items.Item[i].SubItems[1]:=BTS(ResRen);
     //Если всё ОК, то меняем имена в таблице на реальные
     if ResRen then
      begin
       //Первый столбец - только имя
       ListViewFiles.Items.Item[i].Caption:=ExtractFileName(TName);
       //Второй столбец - полное имя
       ListViewFiles.Items.Item[i].SubItems[0]:=TName;
      end;
     //Возможность остановить процесс и предотвратить заблуждение о зависании
     Application.ProcessMessages;
     if Stop then Break;
    end;
   StopProc;
  end
 else
  begin
   if MemoFiles.Lines.Count <= 0 then
    begin
     MessageBox(Handle, 'Список файлов пуст!', '', MB_ICONWARNING or MB_OK);
     Exit;
    end;
   if MessageBox(Handle, 'Вы действительно хотите выполнить операцию переименования?', '', MB_ICONWARNING or MB_YESNO)<>ID_YES then Exit;
   Proc;
   //Начинаем счет с нужного значения
   Num:=GetNums;
   //Идём по таблице
   for i:=0 to MemoFiles.Lines.Count-1 do
    begin
     //Каталог файлов
     DName:=EditPrefDir.Text;
     if DName <> '' then DName:=DName+'\';
     //Текущее имя файла
     CName:=DName+MemoFiles.Lines.Strings[i];
     //Новое имя - пока текущее
     TName:=CName;
     //Изменить имя
     if CheckBoxFN.Checked then TName:=ExtractFilePath(CName)+GetName(i, Num)+ExtractFileExt(CName);
     //Изменить расширение
     if CheckBoxEF.Checked then TName:=ChangeFileExt(TName, GetExt);
     //Переименновываем
     ResRen:=RenameFile(CName, TName);
     //Покажем успех переименнования в таблице
     //MemoFiles.Lines.Strings[i].SubItems[1]:=BTS(ResRen);
     //Если всё ОК, то меняем имена в таблице на реальные
     if ResRen then
      begin
       //Первый столбец - только имя
       MemoFiles.Lines.Strings[i]:=ExtractFileName(TName);
       //Второй столбец - полное имя
       //MemoFiles.Lines.Strings[i].SubItems[0]:=TName;
      end;
     //Возможность остановить процесс и предотвратить заблуждение о зависании
     Application.ProcessMessages;
     if Stop then Break;
    end;
   StopProc;
  end;
end;

function CharReplace(Str:string; TargetChar, NewChar:Char):string;
var i:Integer;
    CP:Integer;
begin
 Result:=Str;
 if Length(Str) <= 0 then Exit;
 repeat
  CP:=Pos(TargetChar, Str);
  if CP <> 0 then
   begin
    if NewChar <> #0 then Str[CP]:=NewChar
    else Delete(Str, CP, 1);
   end;
 until (CP = 0) or (Length(Str) <= 0);
 Result:=Str;
end;

procedure TFormMain.ButtonRenameNormClick(Sender: TObject);
var i, j, Num:Word;
  TName, CName, DName:string;
  ResRen:Boolean;
  C1, C2:Char;
begin
 if ListViewFiles.Visible then
  begin
   if ListViewFiles.Items.Count <= 0 then
    begin
     MessageBox(Handle, 'Список файлов пуст!', '', MB_ICONWARNING or MB_OK);
     Exit;
    end;
   if MessageBox(Handle, 'Вы действительно хотите выполнить операцию переименования?', '', MB_ICONWARNING or MB_YESNO)<>ID_YES then Exit;
   Proc;
   //Начинаем счет с нужного значения
   Num:=GetNums;
   //Идём по таблице
   for i:=0 to ListViewFiles.Items.Count-1 do
    begin
     //Текущее имя файла
     CName:=ListViewFiles.Items.Item[i].SubItems[0];

     //Новое имя
     if RadioButtonNameDesc.Checked then
      TName:=GetFileDescription(CName, GetFileNameWoE(CName))
     else TName:=GetFileNameWoE(CName);

     //Удаляем или заменяем пробелы
     if CheckBoxCharRep.Checked then
      begin
       if EditSpaceChar.Text <> '' then TName:=CharReplace(TName, ' ', EditSpaceChar.Text[1])
       else TName:=CharReplace(TName, ' ', #0);
      end;

     //Заменяем символы
     if CheckBoxRepChars.Checked then
      begin
       if EditCharRepTarget.Text <> '' then C1:=EditCharRepTarget.Text[1] else C1:=#0;
       if EditCharRepTo.Text <> '' then C2:=EditCharRepTo.Text[1] else C2:=#0;
       TName:=CharReplace(TName, C1, C2);
      end;

     //Переименновываем
     ResRen:=RenameFile(CName, ExtractFilePath(CName)+TName+ExtractFileExt(CName));
     //Покажем успех переименнования в таблице
     ListViewFiles.Items.Item[i].SubItems[1]:=BTS(ResRen);
     //Если всё ОК, то меняем имена в таблице на реальные
     if ResRen then
      begin
       //Первый столбец - только имя
       ListViewFiles.Items.Item[i].Caption:=ExtractFileName(TName);
       //Второй столбец - полное имя
       ListViewFiles.Items.Item[i].SubItems[0]:=TName;
      end;
     //Возможность остановить процесс и предотвратить заблуждение о зависании
     Application.ProcessMessages;
     if Stop then Break;
    end;
   StopProc;
  end
 else
  begin
   if MemoFiles.Lines.Count <= 0 then
    begin
     MessageBox(Handle, 'Список файлов пуст!', '', MB_ICONWARNING or MB_OK);
     Exit;
    end;
   if MessageBox(Handle, 'Вы действительно хотите выполнить операцию переименования?', '', MB_ICONWARNING or MB_YESNO)<>ID_YES then Exit;
   Proc;
   //Начинаем счет с нужного значения
   Num:=GetNums;
   //Идём по таблице
   for i:=0 to MemoFiles.Lines.Count-1 do
    begin
     //Каталог файлов
     DName:=EditPrefDir.Text;
     if DName <> '' then DName:=DName+'\';
     //Текущее имя файла
     CName:=DName+MemoFiles.Lines.Strings[i];
     //Новое имя - пока текущее
     TName:=CName;
     //Изменить имя
     if CheckBoxFN.Checked then TName:=ExtractFilePath(CName)+GetName(i, Num)+ExtractFileExt(CName);
     //Изменить расширение
     if CheckBoxEF.Checked then TName:=ChangeFileExt(TName, GetExt);
     //Переименновываем
     ResRen:=RenameFile(CName, TName);
     //Покажем успех переименнования в таблице
     //MemoFiles.Lines.Strings[i].SubItems[1]:=BTS(ResRen);
     //Если всё ОК, то меняем имена в таблице на реальные
     if ResRen then
      begin
       //Первый столбец - только имя
       MemoFiles.Lines.Strings[i]:=ExtractFileName(TName);
       //Второй столбец - полное имя
       //MemoFiles.Lines.Strings[i].SubItems[0]:=TName;
      end;
     //Возможность остановить процесс и предотвратить заблуждение о зависании
     Application.ProcessMessages;
     if Stop then Break;
    end;
   StopProc;
  end;
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
 StopProc;
 originalLVWindowProc:=ListViewFiles.WindowProc;
 ListViewFiles.WindowProc:=LVWindowProc;
 DragAcceptFiles(ListViewFiles.Handle, True);
 ChangeMode(mdGroupRen);
end;

procedure TFormMain.ButtonDeleteClick(Sender: TObject);
begin
 ListViewFiles.DeleteSelected;
end;

procedure TFormMain.ActionAddFilesDirExecute(Sender: TObject);
var dir:string;
begin
 if AdvSelectDirectory('Выберите каталог...', '', dir, False, False, False) then
  begin
   Proc;
   Scan(dir);
   StopProc;
  end;
end;

procedure TFormMain.ActionAddFilesExecute(Sender: TObject);
var i:Word;
begin
 if OpenDialogFiles.Execute then
  for i:=0 to OpenDialogFiles.Files.Count-1 do AddFile(ListViewFiles, OpenDialogFiles.Files.Strings[i]);
end;

procedure TFormMain.ActionAddHiddenSkipExecute(Sender: TObject);
begin
 //
end;

procedure TFormMain.ActionAddRecDirExecute(Sender: TObject);
begin
 //
end;

procedure TFormMain.ActionCLRListExecute(Sender: TObject);
begin
 case FMode of
  mdGroupRen: ListViewFiles.Clear;
  mdFixListRen: MemoFiles.Clear;
 end;
end;

procedure TFormMain.ActionExitExecute(Sender: TObject);
begin
 Application.Terminate;
end;

procedure TFormMain.ActionMFixListRenExecute(Sender: TObject);
begin
 ChangeMode(mdFixListRen);
end;

procedure TFormMain.ActionMGroupRenExecute(Sender: TObject);
begin
 ChangeMode(mdGroupRen);
end;

procedure TFormMain.ActionModeFNNormalExecute(Sender: TObject);
begin
 PageControl.ActivePage:=TabSheetModeFNNormal;
end;

procedure TFormMain.ActionModeNumExecute(Sender: TObject);
begin
 PageControl.ActivePage:=TabSheetModeNum;
end;

procedure TFormMain.ActionStopRenExecute(Sender: TObject);
begin
 StopProc;
end;

initialization
  OleInitialize(nil);
finalization
  OleUninitialize;

end.
