unit checkMem;
interface
implementation

uses  sysUtils, Windows;
var HPs : THeapStatus;
var HPe : THeapStatus;
var lost: integer;
initialization
   HPs := getHeapStatus;
finalization
   HPe := getHeapStatus;
   Lost:= HPe.TotalAllocated - HPs.TotalAllocated;
   if lost >  0
   then begin
        MessageBox(0,PCHar('Memory leak detected! Lost: ' + IntToStr(lost)),'',MB_ICONHAND);
        end;
end.

