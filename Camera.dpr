library Camera;

uses
  ComServ,
  Camera_TLB in 'Camera_TLB.pas',
  CameraActiveFormImpl1 in 'CameraActiveFormImpl1.pas' {CameraActiveFormX: TActiveForm} {CameraActiveFormX: CoClass};

{$E ocx}

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
