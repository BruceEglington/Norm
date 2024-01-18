unit normtern;
interface
uses
  normvarb;

procedure QzAbOr;
procedure QzNeKp;
procedure OrAbAn;


implementation

procedure QzAbOr;
{
Qz-Ab-Or diag.
}
var
  F               : double;
begin
   TQZ:=0.0;
   WTQZ:=0.0;
   TTAB:=0.0;
   WTAB:=0.0;
   TOR:=0.0;
   WTOR:=0.0;
   if (LC+NE+KP)<=0.0 then
   begin
      if (WAB+WOR+WQZ)>0.0 then
      begin
         F:=100.0/(WPL+WOR+WQZ);
         WTQZ:=WQZ*F;
         WTAB:=WAB*F;
         WTOR:=WOR*F;
         F:=100.0/(WTAB+WTOR+WTQZ);
         WTQZ:=WTQZ*F;
         WTAB:=WTAB*F;
         WTOR:=WTOR*F;
         F:=100.0/(NPL+ORT+QZ);
         TQZ:=QZ*F;
         TTAB:=AB*F;
         TOR:=ORT*F;
         F:=100.0/(TTAB+TOR+TQZ);
         TQZ:=TQZ*F;
         TTAB:=TTAB*F;
         TOR:=TOR*F;
      end;{if}
   end;{if}
end;{proc QzAbOr}


procedure QzNeKp;
{
Qz-Ne-Kp diag.
}
var
  F               : double;
begin
   TTQZ:=0.0;
   WTTQZ:=0.0;
   TKP:=0.0;
   WTKP:=0.0;
   TNE:=0.0;
   WTNE:=0.0;
   if DIX>0.0 then
   begin
      F:=100.0/(QZ+0.6*(AB+ORT)+0.5*LC+0.3333333*(NE+KP)+(0.333333*NE+0.2*AB)+(0.2*ORT+0.25*LC+0.333333*KP));
      TTQZ:=(QZ+0.6*(AB+ORT)+0.5*LC+0.3333333*(NE+KP))*F;
      TNE:=(0.333333*NE+0.2*AB)*F;
      TKP:=(0.2*ORT+0.25*LC+0.333333*KP)*F;
      WF:=100.0/WDIX;
      WTKP:=(WKP+WOR*0.56826+WLC*0.72469)*WF;
      WTNE:=(WNE+WAB*0.54173)*WF;
      WTTQZ:=100.0-WTKP-WTNE;
   end;{if}
end;{proc QzNeKp}


procedure OrAbAn;
{
Or-Ab-An diag.
}
var
  F               : double;
begin
   WTFAN:=0.0;
   WTFAB:=0.0;
   WTFOR:=0.0;
   TFAN:=0.0;
   TFOR:=0.0;
   TFAB:=0.0;
   if (NPL+ORT)>0.0 then
   begin
      F:=100.0/(WPL+WOR+WQZ);
      WTFAN:=WAN*F;
      WTFAB:=WAB*F;
      WTFOR:=WOR*F;
      F:=100.0/(WTFAN+WTFAB+WTFOR);
      WTFAN:=WTFAN*F;
      WTFAB:=WTFAB*F;
      WTFOR:=WTFOR*F;
      F:=100.0/(NPL+ORT+QZ);
      TFAN:=AN*F;
      TFAB:=AB*F;
      TFOR:=ORT*F;
      F:=100.0/(TFAN+TFAB+TFOR);
      TFAN:=TFAN*F;
      TFAB:=TFAB*F;
      TFOR:=TFOR*F;
   end;{if}
end;{proc OrAbAn}


end.
