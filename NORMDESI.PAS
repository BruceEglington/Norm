unit normdesi;
interface
uses
  normvarb;

procedure DesilHy;
procedure DesilSp;
procedure DesilAb;
procedure DesilOr;
procedure DesilWo;
procedure DesilDi;
procedure DesilLc;
procedure DesilAct;
procedure DesilOl;

implementation

procedure DesilHy;
{
Desilicify Hy to Ol
}
begin
   if (-4.0*SI)<HY
   then begin
      OL:=-3.0*SI;
      HY:=HY+4.0*SI;
      SI:=0.0;
   end
   else begin
      OL:=0.75*HY;
      SI:=SI+0.25*HY;
      HY:=0.0;
   end;
end;{proc DesilHy}


procedure DesilSp;
{
Desilicify Sp to Pv
}
begin
   if (-3.0*SI)<SP
   then begin
      PF:=-2.0*SI;
      SP:=SP+3.0*SI;
      SI:=0.0;
   end
   else begin
      PF:=0.66667*SP;
      SI:=SI+0.3333*SP;
      SP:=0.0;
   end;
end;{proc DesilSp}


procedure DesilAb;
{
Desilicify Ab to Ne
}
begin
   if (-2.5*SI)<AB
   then begin
      NE:=-1.5*SI;
      AB:=AB+2.5*SI;
      SI:=0.0;
   end
   else begin
      NE:=0.6*AB;
      SI:=SI+0.4*AB;
      AB:=0.0;
   end;
end;{proc DesilAb}


procedure DesilOr;
{
Desilicify Or to Lc
}
begin
   if (-5.0*SI)<ORT
   then begin
      LC:=-4.0*SI;
      ORT:=ORT+5.0*SI;
      SI:=0.0;
   end
   else begin
      LC:=0.8*ORT;
      SI:=SI+0.2*ORT;
      ORT:=0.0;
   end;
end;{proc DesilOr}


procedure DesilWo;
{
Desilicify Wo to Ca-orthosilicate
}
begin
   if (-4.0*SI)<WO
   then begin
      CS:=-3.0*SI;
      WO:=WO+4.0*SI;
      SI:=0.0;
   end
   else begin
      CS:=0.75*WO;
      SI:=SI+0.25*WO;
      WO:=0.0;
   end;
end;{proc DesilWo}


procedure DesilDi;
{
Desilicify Di to Ol and Ca-orthosilicate
}
begin
   if (-4.0*SI)<DI
   then begin
      CS:=CS-1.5*SI;
      OL:=OL-1.5*SI;
      DI:=DI+4.0*SI;
      SI:=0.0;
   end
   else begin
      CS:=CS+0.375*DI;
      OL:=OL+0.375*DI;
      SI:=SI+0.25*DI;
      DI:=0.0;
   end;
end;{proc DesilDi}


procedure DesilLc;
{
Desilicify Lc to Kl
}
begin
   if (-4.0*SI)<LC
   then begin
      KP:=-3.0*SI;
      LC:=LC+4.0*SI;
      SI:=0.0;
   end
   else begin
      KP:=0.75*LC;
      SI:=SI+0.25*LC;
      LC:=0.0;
   end;
end;{proc DesilLc}


procedure DesilAct;
{
Desilicify actinolite
}
begin
   if (3.0*AB)>ACT then
   begin
      if ACT>(-3.75*SI) then
      begin
         ED:=-4.0*SI;
         AB:=AB+1.25*SI;
         ACT:=ACT+3.75*SI;
         SI:=0.0;
      end
      else begin
         ED:=1.067*ACT;
         AB:=AB-0.333*ACT;
         SI:=SI+0.267*ACT;
         ACT:=0.0;
      end
   end
   else begin
      if AB>(-1.25*SI) then
      begin
         ED:=-4.0*SI;
         AB:=AB+1.25*SI;
         ACT:=ACT+3.75*SI;
         SI:=0.0;
      end
      else begin
         ED:=3.2*AB;
         ACT:=ACT-3.0*AB;
         SI:=SI+0.8*AB;
         AB:=0.0;
      end;
   end;
end;{proc DesilAct}


procedure DesilOl;
{
Desilicify olivine to spinel
}
begin
   if (0.75*CO)>OL then
   begin
      if OL>(-3.0*SI) then
      begin
         SPIN:=-6.0*SI;
         CO:=CO+4.0*SI;
         OL:=OL+3.0*SI;
         SI:=0.0;
      end
      else begin
         SPIN:=2.0*OL;
         CO:=CO-1.333*OL;
         SI:=SI+0.333*OL;
         OL:=0.0;
      end
   end
   else begin
      if CO>(-4.0*SI) then
      begin
         SPIN:=-6.0*SI;
         CO:=CO+4.0*SI;
         OL:=OL+3.0*SI;
         SI:=0.0;
      end
      else begin
         SPIN:=1.5*CO;
         OL:=OL-0.75*CO;
         SI:=SI+0.25*CO;
         CO:=0.0;
      end;
   end;
end;{proc DesilOl}


end.

