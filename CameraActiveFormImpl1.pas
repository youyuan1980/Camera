unit CameraActiveFormImpl1;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ActiveX, AxCtrls, Camera_TLB, StdVcl,EncdDecd, ExtCtrls,Registry,jpeg;

type
  TCameraActiveFormX = class(TActiveForm, ICameraActiveFormX)
    procedure ActiveFormCreate(Sender: TObject);
  private
    ReadSysDirectory:string;
    { Private declarations }
    FEvents: ICameraActiveFormXEvents;
    procedure ActivateEvent(Sender: TObject);
    procedure ClickEvent(Sender: TObject);
    procedure CreateEvent(Sender: TObject);
    procedure DblClickEvent(Sender: TObject);
    procedure DeactivateEvent(Sender: TObject);
    procedure DestroyEvent(Sender: TObject);
    procedure KeyPressEvent(Sender: TObject; var Key: Char);
    procedure PaintEvent(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

  protected
    { Protected declarations }
    procedure DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage); override;
    procedure EventSinkChanged(const EventSink: IUnknown); override;
    function Get_Active: WordBool; safecall;
    function Get_AlignDisabled: WordBool; safecall;
    function Get_AutoScroll: WordBool; safecall;
    function Get_AutoSize: WordBool; safecall;
    function Get_AxBorderStyle: TxActiveFormBorderStyle; safecall;
    function Get_Caption: WideString; safecall;
    function Get_Color: OLE_COLOR; safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    function Get_DropTarget: WordBool; safecall;
    function Get_Enabled: WordBool; safecall;
    function Get_Font: IFontDisp; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_KeyPreview: WordBool; safecall;
    function Get_PixelsPerInch: Integer; safecall;
    function Get_PrintScale: TxPrintScale; safecall;
    function Get_Scaled: WordBool; safecall;
    function Get_ScreenSnap: WordBool; safecall;
    function Get_SnapBuffer: Integer; safecall;
    function Get_Visible: WordBool; safecall;
    function Get_VisibleDockClientCount: Integer; safecall;
    procedure _Set_Font(var Value: IFontDisp); safecall;
    procedure Set_AutoScroll(Value: WordBool); safecall;
    procedure Set_AutoSize(Value: WordBool); safecall;
    procedure Set_AxBorderStyle(Value: TxActiveFormBorderStyle); safecall;
    procedure Set_Caption(const Value: WideString); safecall;
    procedure Set_Color(Value: OLE_COLOR); safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    procedure Set_DropTarget(Value: WordBool); safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure Set_Font(const Value: IFontDisp); safecall;
    procedure Set_HelpFile(const Value: WideString); safecall;
    procedure Set_KeyPreview(Value: WordBool); safecall;
    procedure Set_PixelsPerInch(Value: Integer); safecall;
    procedure Set_PrintScale(Value: TxPrintScale); safecall;
    procedure Set_Scaled(Value: WordBool); safecall;
    procedure Set_ScreenSnap(Value: WordBool); safecall;
    procedure Set_SnapBuffer(Value: Integer); safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    procedure Capture_Photo; safecall;
    function Get_Acquire_PhotoArrayTo64Str: WideString; safecall;
    procedure Set_Acquire_PhotoArrayTo64Str(const Value: WideString); safecall;
    function BitmapToString(img:TJPegImage):string;
    procedure pic_bmptojpg(FileName:string);
    procedure Close_Camera; safecall;
  public
    { Public declarations }
    procedure Initialize; override;
  end;

  const WM_CAP_START = WM_USER;
  const WM_CAP_STOP = WM_CAP_START + 68;
  const WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10;
  const WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11;
  const WM_CAP_SAVEDIB = WM_CAP_START + 25;
  const WM_CAP_GRAB_FRAME = WM_CAP_START + 60;
  const WM_CAP_SEQUENCE = WM_CAP_START + 62;
  const WM_CAP_FILE_SET_CAPTURE_FILEA = WM_CAP_START + 20;
  const WM_CAP_SEQUENCE_NOFILE =WM_CAP_START+ 63;
  const WM_CAP_SET_OVERLAY =WM_CAP_START+ 51;
  const WM_CAP_SET_PREVIEW =WM_CAP_START+ 50;
  const WM_CAP_SET_CALLBACK_VIDEOSTREAM = WM_CAP_START +6;
  const WM_CAP_SET_CALLBACK_ERROR=WM_CAP_START +2;
  const WM_CAP_SET_CALLBACK_STATUSA= WM_CAP_START +3;
  const WM_CAP_SET_CALLBACK_FRAME= WM_CAP_START +5;
  const WM_CAP_SET_SCALE=WM_CAP_START+ 53;
  const WM_CAP_SET_PREVIEWRATE=WM_CAP_START+ 52;
  function capCreateCaptureWindowA(lpszWindowName : PCHAR; dwStyle : longint; x : integer;
  y : integer;nWidth : integer;nHeight : integer;ParentWin : HWND;
  nId : integer): HWND;STDCALL EXTERNAL 'AVICAP32.DLL';

  var
  hWndC:HWND;
implementation

uses ComObj, ComServ;

{$R *.DFM}

{ TCameraActiveFormX }

procedure TCameraActiveFormX.DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage);
begin
  
  { Define property pages here.  Property pages are defined by calling
    DefinePropertyPage with the class id of the page.  For example,
      DefinePropertyPage(Class_CameraActiveFormXPage); }
end;

procedure TCameraActiveFormX.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as ICameraActiveFormXEvents;
  inherited EventSinkChanged(EventSink);
end;

procedure TCameraActiveFormX.Initialize;
begin
  inherited Initialize;
  OnActivate := ActivateEvent;
  OnClick := ClickEvent;
  OnCreate := CreateEvent;
  OnDblClick := DblClickEvent;
  OnDeactivate := DeactivateEvent;
  OnDestroy := DestroyEvent;
  OnKeyPress := KeyPressEvent;
  OnPaint := PaintEvent;
end;

function TCameraActiveFormX.Get_Active: WordBool;
begin
  Result := Active;
end;

function TCameraActiveFormX.Get_AlignDisabled: WordBool;
begin
  Result := AlignDisabled;
end;

function TCameraActiveFormX.Get_AutoScroll: WordBool;
begin
  Result := AutoScroll;
end;

function TCameraActiveFormX.Get_AutoSize: WordBool;
begin
  Result := AutoSize;
end;

function TCameraActiveFormX.Get_AxBorderStyle: TxActiveFormBorderStyle;
begin
  Result := Ord(AxBorderStyle);
end;

function TCameraActiveFormX.Get_Caption: WideString;
begin
  Result := WideString(Caption);
end;

function TCameraActiveFormX.Get_Color: OLE_COLOR;
begin
  Result := OLE_COLOR(Color);
end;

function TCameraActiveFormX.Get_DoubleBuffered: WordBool;
begin
  Result := DoubleBuffered;
end;

function TCameraActiveFormX.Get_DropTarget: WordBool;
begin
  Result := DropTarget;
end;

function TCameraActiveFormX.Get_Enabled: WordBool;
begin
  Result := Enabled;
end;

function TCameraActiveFormX.Get_Font: IFontDisp;
begin
  GetOleFont(Font, Result);
end;

function TCameraActiveFormX.Get_HelpFile: WideString;
begin
  Result := WideString(HelpFile);
end;

function TCameraActiveFormX.Get_KeyPreview: WordBool;
begin
  Result := KeyPreview;
end;

function TCameraActiveFormX.Get_PixelsPerInch: Integer;
begin
  Result := PixelsPerInch;
end;

function TCameraActiveFormX.Get_PrintScale: TxPrintScale;
begin
  Result := Ord(PrintScale);
end;

function TCameraActiveFormX.Get_Scaled: WordBool;
begin
  Result := Scaled;
end;

function TCameraActiveFormX.Get_ScreenSnap: WordBool;
begin
  Result := ScreenSnap;
end;

function TCameraActiveFormX.Get_SnapBuffer: Integer;
begin
  Result := SnapBuffer;
end;

function TCameraActiveFormX.Get_Visible: WordBool;
begin
  Result := Visible;
end;

function TCameraActiveFormX.Get_VisibleDockClientCount: Integer;
begin
  Result := VisibleDockClientCount;
end;

procedure TCameraActiveFormX._Set_Font(var Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TCameraActiveFormX.ActivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnActivate;
end;

procedure TCameraActiveFormX.ClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnClick;
end;

procedure TCameraActiveFormX.CreateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnCreate;
end;

procedure TCameraActiveFormX.DblClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDblClick;
end;

procedure TCameraActiveFormX.DeactivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDeactivate;
end;

procedure TCameraActiveFormX.DestroyEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDestroy;
end;

procedure TCameraActiveFormX.KeyPressEvent(Sender: TObject; var Key: Char);
var
  TempKey: Smallint;
begin
  TempKey := Smallint(Key);
  if FEvents <> nil then FEvents.OnKeyPress(TempKey);
  Key := Char(TempKey);
end;

procedure TCameraActiveFormX.PaintEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnPaint;
end;

procedure TCameraActiveFormX.Set_AutoScroll(Value: WordBool);
begin
  AutoScroll := Value;
end;

procedure TCameraActiveFormX.Set_AutoSize(Value: WordBool);
begin
  AutoSize := Value;
end;

procedure TCameraActiveFormX.Set_AxBorderStyle(
  Value: TxActiveFormBorderStyle);
begin
  AxBorderStyle := TActiveFormBorderStyle(Value);
end;

procedure TCameraActiveFormX.Set_Caption(const Value: WideString);
begin
  Caption := TCaption(Value);
end;

procedure TCameraActiveFormX.Set_Color(Value: OLE_COLOR);
begin
  Color := TColor(Value);
end;

procedure TCameraActiveFormX.Set_DoubleBuffered(Value: WordBool);
begin
  DoubleBuffered := Value;
end;

procedure TCameraActiveFormX.Set_DropTarget(Value: WordBool);
begin
  DropTarget := Value;
end;

procedure TCameraActiveFormX.Set_Enabled(Value: WordBool);
begin
  Enabled := Value;
end;

procedure TCameraActiveFormX.Set_Font(const Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TCameraActiveFormX.Set_HelpFile(const Value: WideString);
begin
  HelpFile := String(Value);
end;

procedure TCameraActiveFormX.Set_KeyPreview(Value: WordBool);
begin
  KeyPreview := Value;
end;

procedure TCameraActiveFormX.Set_PixelsPerInch(Value: Integer);
begin
  PixelsPerInch := Value;
end;

procedure TCameraActiveFormX.Set_PrintScale(Value: TxPrintScale);
begin
  PrintScale := TPrintScale(Value);
end;

procedure TCameraActiveFormX.Set_Scaled(Value: WordBool);
begin
  Scaled := Value;
end;

procedure TCameraActiveFormX.Set_ScreenSnap(Value: WordBool);
begin
  ScreenSnap := Value;
end;

procedure TCameraActiveFormX.Set_SnapBuffer(Value: Integer);
begin
  SnapBuffer := Value;
end;

procedure TCameraActiveFormX.Set_Visible(Value: WordBool);
begin
  Visible := Value;
end;

procedure TCameraActiveFormX.ActiveFormCreate(Sender: TObject);
var
 reg:Tregistry;
 KeyList: TStringList;
 x:Integer;
begin
    reg:=Tregistry.create;
    KeyList := TStringList.Create;
    reg.RootKey := HKEY_CLASSES_ROOT;
    reg.openkey('CLSID\{B80E027B-4903-4242-8CD4-D39308A9ED1C}\InprocServer32', False);
    reg.GetValueNames(KeyList);
    for x := 0 to KeyList.Count - 1 do
    begin
      ReadSysDirectory:=extractfilepath(reg.ReadString(KeyList[0]))+'\photo.jpg';
    end;
    reg.CloseKey;
    reg.free;
    KeyList.Free;
hWndC:=capCreateCaptureWindowA('My Own Capture Window',WS_CHILD or WS_VISIBLE ,0,0,320,240,Handle,0);
if hWndC<>0 then
  begin
  SendMessage(hWndC, WM_CAP_DRIVER_DISCONNECT, 0, 0);
  SendMessage(hWndC, WM_CAP_SET_CALLBACK_VIDEOSTREAM, 0, 0);
  SendMessage(hWndC, WM_CAP_SET_CALLBACK_ERROR, 0, 0);
  SendMessage(hWndC, WM_CAP_SET_CALLBACK_STATUSA, 0, 0);
  SendMessage(hWndC, WM_CAP_DRIVER_CONNECT, 0, 0);
  SendMessage(hWndC, WM_CAP_SET_SCALE, 1, 0);
  SendMessage(hWndC, WM_CAP_SET_PREVIEWRATE, 66, 0);
  SendMessage(hWndC, WM_CAP_SET_OVERLAY, 1, 0);
  SendMessage(hWndC, WM_CAP_SET_PREVIEW, 1, 0);
  end;
end;
procedure TCameraActiveFormX.Capture_Photo;
begin
    SendMessage(hWndC, WM_CAP_SAVEDIB, 0, longint(PCHAR(ReadSysDirectory)));

end;

function TCameraActiveFormX.BitmapToString(img:TJPegImage):string ;
var
  ms:TMemoryStream;
  ss:TStringStream;
  s:string;
begin
    ms := TMemoryStream.Create;
    img.SaveToStream(ms);
    ss := TStringStream.Create('');
    ms.Position:=0;
    EncodeStream(ms,ss);//将内存流编码为base64字符流
    s:=ss.DataString;
    ms.Free;
    ss.Free;
    result:=s; 
end;


function TCameraActiveFormX.Get_Acquire_PhotoArrayTo64Str: WideString;
var
  AImage : TImage;
  ABitmap: TJPegImage;
  PhotoImageBytes:WideString;
begin
  pic_bmptojpg(ReadSysDirectory);
  AImage := TImage.Create(nil);
  try
    AImage.Picture.LoadFromFile(ReadSysDirectory);
    ABitmap := TJPEGImage.Create;
    try
      ABitmap.Assign(AImage.Picture.Graphic);
      //下面可以处理这个Bitmap了
      PhotoImageBytes:=BitmapToString(ABitmap);
    finally
      ABitmap.Free;
    end;
  finally
    AImage.Free;
  end;
  result:=PhotoImageBytes;
end;

procedure TCameraActiveFormX.pic_bmptojpg(FileName:string);
Var
Bitmap: TBitmap;
JPeg: TJPegImage;
Begin
  Bitmap := Nil;
  JPeg := Nil;
  Try
    Bitmap := TBitmap.Create;
    Bitmap.LoadFromFile(FileName);
    JPeg := TJPegImage.Create;
    JPeg.Assign(Bitmap);
    JPeg.SaveToFile(ReadSysDirectory);
Finally
    FreeAndNil(Bitmap);
    FreeAndNil(JPeg);
End;
End;

procedure TCameraActiveFormX.Set_Acquire_PhotoArrayTo64Str(
  const Value: WideString);
begin

end;

procedure TCameraActiveFormX.Close_Camera;
begin
      if hWndC <> 0 then begin
      SendMessage(hWndC, WM_CAP_DRIVER_DISCONNECT, 0, 0);
    end;
end;

procedure TCameraActiveFormX.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if hWndC <> 0 then begin
  SendMessage(hWndC, WM_CAP_DRIVER_DISCONNECT, 0, 0);
end;
end;

initialization
  TActiveFormFactory.Create(
    ComServer,
    TActiveFormControl,
    TCameraActiveFormX,
    Class_CameraActiveFormX,
    1,
    '',
    OLEMISC_SIMPLEFRAME or OLEMISC_ACTSLIKELABEL,
    tmApartment);
end.
