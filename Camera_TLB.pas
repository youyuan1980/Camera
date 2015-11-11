unit Camera_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : 1.2
// File generated on 2013-10-29 09:24:23 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\Program Files\Borland\Delphi7\Projects\Camera\Camera.tlb (1)
// LIBID: {7768FDC4-EF1A-4AFD-97C7-9E67C0E8E22E}
// LCID: 0
// Helpfile: 
// HelpString: Camera Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleCtrls, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  CameraMajorVersion = 1;
  CameraMinorVersion = 0;

  LIBID_Camera: TGUID = '{7768FDC4-EF1A-4AFD-97C7-9E67C0E8E22E}';

  IID_ICameraActiveFormX: TGUID = '{6E367F9F-BF71-4702-B89E-1ADD9F5A39FD}';
  DIID_ICameraActiveFormXEvents: TGUID = '{B77C91EC-EF6F-4F13-A769-BC64FB55A2BB}';
  CLASS_CameraActiveFormX: TGUID = '{B80E027B-4903-4242-8CD4-D39308A9ED1C}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum TxActiveFormBorderStyle
type
  TxActiveFormBorderStyle = TOleEnum;
const
  afbNone = $00000000;
  afbSingle = $00000001;
  afbSunken = $00000002;
  afbRaised = $00000003;

// Constants for enum TxPrintScale
type
  TxPrintScale = TOleEnum;
const
  poNone = $00000000;
  poProportional = $00000001;
  poPrintToFit = $00000002;

// Constants for enum TxMouseButton
type
  TxMouseButton = TOleEnum;
const
  mbLeft = $00000000;
  mbRight = $00000001;
  mbMiddle = $00000002;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  ICameraActiveFormX = interface;
  ICameraActiveFormXDisp = dispinterface;
  ICameraActiveFormXEvents = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  CameraActiveFormX = ICameraActiveFormX;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  PPUserType1 = ^IFontDisp; {*}


// *********************************************************************//
// Interface: ICameraActiveFormX
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6E367F9F-BF71-4702-B89E-1ADD9F5A39FD}
// *********************************************************************//
  ICameraActiveFormX = interface(IDispatch)
    ['{6E367F9F-BF71-4702-B89E-1ADD9F5A39FD}']
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    function Get_AutoScroll: WordBool; safecall;
    procedure Set_AutoScroll(Value: WordBool); safecall;
    function Get_AutoSize: WordBool; safecall;
    procedure Set_AutoSize(Value: WordBool); safecall;
    function Get_AxBorderStyle: TxActiveFormBorderStyle; safecall;
    procedure Set_AxBorderStyle(Value: TxActiveFormBorderStyle); safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const Value: WideString); safecall;
    function Get_Color: OLE_COLOR; safecall;
    procedure Set_Color(Value: OLE_COLOR); safecall;
    function Get_Font: IFontDisp; safecall;
    procedure Set_Font(const Value: IFontDisp); safecall;
    procedure _Set_Font(var Value: IFontDisp); safecall;
    function Get_KeyPreview: WordBool; safecall;
    procedure Set_KeyPreview(Value: WordBool); safecall;
    function Get_PixelsPerInch: Integer; safecall;
    procedure Set_PixelsPerInch(Value: Integer); safecall;
    function Get_PrintScale: TxPrintScale; safecall;
    procedure Set_PrintScale(Value: TxPrintScale); safecall;
    function Get_Scaled: WordBool; safecall;
    procedure Set_Scaled(Value: WordBool); safecall;
    function Get_Active: WordBool; safecall;
    function Get_DropTarget: WordBool; safecall;
    procedure Set_DropTarget(Value: WordBool); safecall;
    function Get_HelpFile: WideString; safecall;
    procedure Set_HelpFile(const Value: WideString); safecall;
    function Get_ScreenSnap: WordBool; safecall;
    procedure Set_ScreenSnap(Value: WordBool); safecall;
    function Get_SnapBuffer: Integer; safecall;
    procedure Set_SnapBuffer(Value: Integer); safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    function Get_AlignDisabled: WordBool; safecall;
    function Get_VisibleDockClientCount: Integer; safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure Capture_Photo; safecall;
    function Get_Acquire_PhotoArrayTo64Str: WideString; safecall;
    procedure Set_Acquire_PhotoArrayTo64Str(const Value: WideString); safecall;
    procedure Close_Camera; safecall;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property AutoScroll: WordBool read Get_AutoScroll write Set_AutoScroll;
    property AutoSize: WordBool read Get_AutoSize write Set_AutoSize;
    property AxBorderStyle: TxActiveFormBorderStyle read Get_AxBorderStyle write Set_AxBorderStyle;
    property Caption: WideString read Get_Caption write Set_Caption;
    property Color: OLE_COLOR read Get_Color write Set_Color;
    property Font: IFontDisp read Get_Font write Set_Font;
    property KeyPreview: WordBool read Get_KeyPreview write Set_KeyPreview;
    property PixelsPerInch: Integer read Get_PixelsPerInch write Set_PixelsPerInch;
    property PrintScale: TxPrintScale read Get_PrintScale write Set_PrintScale;
    property Scaled: WordBool read Get_Scaled write Set_Scaled;
    property Active: WordBool read Get_Active;
    property DropTarget: WordBool read Get_DropTarget write Set_DropTarget;
    property HelpFile: WideString read Get_HelpFile write Set_HelpFile;
    property ScreenSnap: WordBool read Get_ScreenSnap write Set_ScreenSnap;
    property SnapBuffer: Integer read Get_SnapBuffer write Set_SnapBuffer;
    property DoubleBuffered: WordBool read Get_DoubleBuffered write Set_DoubleBuffered;
    property AlignDisabled: WordBool read Get_AlignDisabled;
    property VisibleDockClientCount: Integer read Get_VisibleDockClientCount;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property Acquire_PhotoArrayTo64Str: WideString read Get_Acquire_PhotoArrayTo64Str write Set_Acquire_PhotoArrayTo64Str;
  end;

// *********************************************************************//
// DispIntf:  ICameraActiveFormXDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6E367F9F-BF71-4702-B89E-1ADD9F5A39FD}
// *********************************************************************//
  ICameraActiveFormXDisp = dispinterface
    ['{6E367F9F-BF71-4702-B89E-1ADD9F5A39FD}']
    property Visible: WordBool dispid 201;
    property AutoScroll: WordBool dispid 202;
    property AutoSize: WordBool dispid 203;
    property AxBorderStyle: TxActiveFormBorderStyle dispid 204;
    property Caption: WideString dispid -518;
    property Color: OLE_COLOR dispid -501;
    property Font: IFontDisp dispid -512;
    property KeyPreview: WordBool dispid 205;
    property PixelsPerInch: Integer dispid 206;
    property PrintScale: TxPrintScale dispid 207;
    property Scaled: WordBool dispid 208;
    property Active: WordBool readonly dispid 209;
    property DropTarget: WordBool dispid 210;
    property HelpFile: WideString dispid 211;
    property ScreenSnap: WordBool dispid 212;
    property SnapBuffer: Integer dispid 213;
    property DoubleBuffered: WordBool dispid 214;
    property AlignDisabled: WordBool readonly dispid 215;
    property VisibleDockClientCount: Integer readonly dispid 216;
    property Enabled: WordBool dispid -514;
    procedure Capture_Photo; dispid 217;
    property Acquire_PhotoArrayTo64Str: WideString dispid 218;
    procedure Close_Camera; dispid 219;
  end;

// *********************************************************************//
// DispIntf:  ICameraActiveFormXEvents
// Flags:     (4096) Dispatchable
// GUID:      {B77C91EC-EF6F-4F13-A769-BC64FB55A2BB}
// *********************************************************************//
  ICameraActiveFormXEvents = dispinterface
    ['{B77C91EC-EF6F-4F13-A769-BC64FB55A2BB}']
    procedure OnActivate; dispid 201;
    procedure OnClick; dispid 202;
    procedure OnCreate; dispid 203;
    procedure OnDblClick; dispid 204;
    procedure OnDestroy; dispid 205;
    procedure OnDeactivate; dispid 206;
    procedure OnKeyPress(var Key: Smallint); dispid 207;
    procedure OnPaint; dispid 208;
  end;


// *********************************************************************//
// OLE Control Proxy class declaration
// Control Name     : TCameraActiveFormX
// Help String      : CameraActiveFormX Control
// Default Interface: ICameraActiveFormX
// Def. Intf. DISP? : No
// Event   Interface: ICameraActiveFormXEvents
// TypeFlags        : (34) CanCreate Control
// *********************************************************************//
  TCameraActiveFormXOnKeyPress = procedure(ASender: TObject; var Key: Smallint) of object;

  TCameraActiveFormX = class(TOleControl)
  private
    FOnActivate: TNotifyEvent;
    FOnClick: TNotifyEvent;
    FOnCreate: TNotifyEvent;
    FOnDblClick: TNotifyEvent;
    FOnDestroy: TNotifyEvent;
    FOnDeactivate: TNotifyEvent;
    FOnKeyPress: TCameraActiveFormXOnKeyPress;
    FOnPaint: TNotifyEvent;
    FIntf: ICameraActiveFormX;
    function  GetControlInterface: ICameraActiveFormX;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    procedure Capture_Photo;
    procedure Close_Camera;
    property  ControlInterface: ICameraActiveFormX read GetControlInterface;
    property  DefaultInterface: ICameraActiveFormX read GetControlInterface;
    property Visible: WordBool index 201 read GetWordBoolProp write SetWordBoolProp;
    property Active: WordBool index 209 read GetWordBoolProp;
    property DropTarget: WordBool index 210 read GetWordBoolProp write SetWordBoolProp;
    property HelpFile: WideString index 211 read GetWideStringProp write SetWideStringProp;
    property ScreenSnap: WordBool index 212 read GetWordBoolProp write SetWordBoolProp;
    property SnapBuffer: Integer index 213 read GetIntegerProp write SetIntegerProp;
    property DoubleBuffered: WordBool index 214 read GetWordBoolProp write SetWordBoolProp;
    property AlignDisabled: WordBool index 215 read GetWordBoolProp;
    property VisibleDockClientCount: Integer index 216 read GetIntegerProp;
    property Enabled: WordBool index -514 read GetWordBoolProp write SetWordBoolProp;
  published
    property Anchors;
    property  ParentColor;
    property  ParentFont;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property AutoScroll: WordBool index 202 read GetWordBoolProp write SetWordBoolProp stored False;
    property AutoSize: WordBool index 203 read GetWordBoolProp write SetWordBoolProp stored False;
    property AxBorderStyle: TOleEnum index 204 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property Caption: WideString index -518 read GetWideStringProp write SetWideStringProp stored False;
    property Color: TColor index -501 read GetTColorProp write SetTColorProp stored False;
    property Font: TFont index -512 read GetTFontProp write SetTFontProp stored False;
    property KeyPreview: WordBool index 205 read GetWordBoolProp write SetWordBoolProp stored False;
    property PixelsPerInch: Integer index 206 read GetIntegerProp write SetIntegerProp stored False;
    property PrintScale: TOleEnum index 207 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property Scaled: WordBool index 208 read GetWordBoolProp write SetWordBoolProp stored False;
    property Acquire_PhotoArrayTo64Str: WideString index 218 read GetWideStringProp write SetWideStringProp stored False;
    property OnActivate: TNotifyEvent read FOnActivate write FOnActivate;
    property OnClick: TNotifyEvent read FOnClick write FOnClick;
    property OnCreate: TNotifyEvent read FOnCreate write FOnCreate;
    property OnDblClick: TNotifyEvent read FOnDblClick write FOnDblClick;
    property OnDestroy: TNotifyEvent read FOnDestroy write FOnDestroy;
    property OnDeactivate: TNotifyEvent read FOnDeactivate write FOnDeactivate;
    property OnKeyPress: TCameraActiveFormXOnKeyPress read FOnKeyPress write FOnKeyPress;
    property OnPaint: TNotifyEvent read FOnPaint write FOnPaint;
  end;

procedure Register;

resourcestring
  dtlServerPage = 'Servers';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

procedure TCameraActiveFormX.InitControlData;
const
  CEventDispIDs: array [0..7] of DWORD = (
    $000000C9, $000000CA, $000000CB, $000000CC, $000000CD, $000000CE,
    $000000CF, $000000D0);
  CTFontIDs: array [0..0] of DWORD = (
    $FFFFFE00);
  CControlData: TControlData2 = (
    ClassID: '{B80E027B-4903-4242-8CD4-D39308A9ED1C}';
    EventIID: '{B77C91EC-EF6F-4F13-A769-BC64FB55A2BB}';
    EventCount: 8;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$00000000*);
    Flags: $0000001D;
    Version: 401;
    FontCount: 1;
    FontIDs: @CTFontIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnActivate) - Cardinal(Self);
end;

procedure TCameraActiveFormX.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as ICameraActiveFormX;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TCameraActiveFormX.GetControlInterface: ICameraActiveFormX;
begin
  CreateControl;
  Result := FIntf;
end;

procedure TCameraActiveFormX.Capture_Photo;
begin
  DefaultInterface.Capture_Photo;
end;

procedure TCameraActiveFormX.Close_Camera;
begin
  DefaultInterface.Close_Camera;
end;

procedure Register;
begin
  RegisterComponents(dtlOcxPage, [TCameraActiveFormX]);
end;

end.
