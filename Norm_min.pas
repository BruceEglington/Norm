unit norm_min;
interface
uses
  normvarb;

procedure Apatite;
procedure Halite;
procedure Thenardite;
procedure Pyrite;
procedure Chromite;
procedure Ilmenite;
procedure Fluorite;
procedure Calcite;
procedure Zircon;
procedure Orthoclase;
procedure Albite;
procedure AnCor;
procedure Acmite;
procedure SphRut;
procedure MtHm;
procedure WollHyp;
procedure Diopside;
procedure Riebeckite;
procedure Biotite;
procedure ActHyWo;
procedure Hornblende;
procedure MakeQz;
procedure F2MgRatio;
procedure SetSiZero;

implementation

procedure Apatite;
begin
   if (FU > 0.0) then begin
      if (1.66667*PO)<CA
      then begin
         CalciumInNonsilicates := CalciumInNonsilicates + 1.66667*PO;
         AP:=3.0*PO;
         CA:=CA-1.66667*PO;
         FU:=FU-0.111111*AP;
         PO:=0.0;
      end
      else begin
         CalciumInNonsilicates := CalciumInNonsilicates + CA;
         AP:=1.8*CA;
         PO:=PO-0.6*CA;
         CA:=0.0;
         FU:=FU-0.111111*AP;
      end;
   end
   else begin
      if (1.66667*PO)<CA
      then begin
         AP:=2.66667*PO;
         CA:=CA-1.66667*PO;
         PO:=0.0;
      end
      else begin
         AP:=1.6*CA;
         PO:=PO-0.6*CA;
         CA:=0.0;
      end;
   end;
end;{proc Apatite}


procedure Halite;
begin
   if CL<NA
   then begin
      HL:=CL*2.0;
      NA:=NA-CL;
      CL:=0.0;
   end
   else begin
      HL:=NA*2.0;
      CL:=CL-NA;
      NA:=0.0;
   end;
end;{proc Halite}

procedure Thenardite;
begin
   if (2.0*SO)<NA
   then begin
      TH:=3.0*SO;
      NA:=NA-2.0*SO;
      SO:=0.0;
   end
   else begin
      TH:=1.5*NA;
      SO:=SO-0.5*NA;
      NA:=0.0;
   end;
end;{proc Thenardite}

procedure Pyrite;
begin
   if SU<(2.0*F2)
   then begin
      PY:=1.5*SU;
      F2:=F2-0.5*SU;
      SU:=0.0;
   end
   else begin
      PY:=F2*3.0;
      SU:=SU-2.0*F2;
      F2:=0.0;
   end;
end;{proc Pyrite}


procedure Chromite;
begin
   if NCR<(2.0*F2)
   then begin
      CM:=1.5*NCR;
      F2:=F2-0.5*NCR;
      NCR:=0.0;
   end
   else begin
      CM:=3.0*F2;
      NCR:=NCR-2.0*F2;
      F2:=0.0;
   end;
end;{proc Chromite}

procedure Ilmenite;
begin
   if TI<F2
   then begin
      IL:=2.0*TI;
      F2:=F2-TI;
      TI:=0.0;
   end
   else begin
      IL:=2.0*F2;
      TI:=TI-F2;
      F2:=0.0;
   end;
end;{proc Ilmenite}

procedure Fluorite;
begin
   if FU>0.0 then begin
      if FU<(2.0*CA)
      then begin
         FL:=1.5*FU;
         CA:=CA-0.5*FU;
         CalciumInNonsilicates := CalciumInNonsilicates + 0.5*FU;
         FU:=0.0;
      end
      else begin
         FL:=3.0*CA;
         CalciumInNonsilicates := CalciumInNonsilicates + CA;
         FU:=FU-2.0*CA;
         CA:=0.0;
      end;
   end;{if}
end;{proc Fluorite}


procedure Calcite;
begin
   if CD<CA
   then begin
      CC:=2.0*CD;
      CA:=CA-CD;
      CalciumInNonsilicates := CalciumInNonsilicates + CD;
      CD:=0.0;
   end
   else begin
      CC:=2.0*CA;
      CD:=CD-CA;
      CalciumInNonsilicates := CalciumInNonsilicates + CA;
      CA:=0.0;
   end;
end;{proc Calcite}

procedure Zircon;
begin
   if ZR<SI
   then begin
      ZN:=2.0*ZR;
      SI:=SI-ZR;
      ZR:=0.0;
   end
   else begin
      ZN:=2.0*SI;
      ZR:=ZR-SI;
      SI:=0.0;
   end;
end;{proc Zircon}

procedure Orthoclase;
{
Orthoclase, Potassium metasilicate
}
begin
   if KO>AL
   then begin
      ORT:=5.0*AL;
      KO:=KO-AL;
      KS:=1.5*KO;
      SI:=SI-0.5*KO-3.0*AL;
      AL:=0.0;
   end
   else begin
      ORT:=5.0*KO;
      SI:=SI-3.0*KO;
      AL:=AL-KO;
      KO:=0.0;
   end;
end;{proc Orthoclase}


procedure Albite;
begin
   if NA>AL
   then begin
      AB:=5.0*AL;
      SI:=SI-3.0*AL;
      NA:=NA-AL;
      AL:=0.0;
   end
   else begin
      AB:=5.0*NA;
      SI:=SI-3.0*NA;
      AL:=AL-NA;
      NA:=0.0;
   end;
end;{proc Albite}


procedure AnCor;
{
Anorthite, Corundum
}
begin
   if AL>(2.0*CA)
   then begin
      AN:=5.0*CA;
      CO:=AL-2.0*CA;
      SI:=SI-2.0*CA;
      CA:=0.0;
      AL:=0.0;
   end
   else begin
      AN:=2.5*AL;
      CA:=CA-0.5*AL;
      SI:=SI-AL;
      AL:=0.0;
   end;
end;{proc AnCor}


procedure Acmite;
{
Acmite, Sodium metasilicate
}
begin
   if NA>F3
   then begin
      AC:=4.0*F3;
      NA:=NA-F3;
      NS:=1.5*NA;
      SI:=SI-0.5*NA-2.0*F3;
      NA:=0.0;
      F3:=0.0;
   end
   else begin
      AC:=4.0*NA;
      SI:=SI-2.0*NA;
      F3:=F3-NA;
      NA:=0.0;
   end;
end;{proc Acmite}

procedure SphRut;
{
Sphene, Rutile
}
begin
   if TI>CA
   then begin
      SP:=3.0*CA;
      SI:=SI-CA;
      RU:=TI-CA;
      CalciumInNonsilicates := CalciumInNonsilicates + CA;
      CA:=0.0;
      TI:=0.0;
   end
   else begin
      SP:=3.0*TI;
      SI:=SI-TI;
      CA:=CA-TI;
      CalciumInNonsilicates := CalciumInNonsilicates + TI;
      TI:=0.0;
   end;
end;{proc SphRut}


procedure MtHm;
{
Magnetite, Hematite
}
begin
   if F3>(2.0*F2)
   then begin
      MT:=3.0*F2;
      HM:=F3-2.0*F2;
      F3:=0.0;
      F2:=0.0;
   end
   else begin
      MT:=1.5*F3;
      F2:=F2-0.5*F3;
      F3:=0.0;
   end;
end;{proc MtHm}


procedure WollHyp;
{
Wollastonite, Hypersthene
}
begin
   WO:=2.0*CA;
   HY:=2.0*FEMG;
   SI:=SI-CA-FEMG;
   CA:=0.0;
end;{proc WollHyp}


procedure Diopside;
begin
   if HY<WO
   then begin
      DI:=2.0*HY;
      WO:=WO-HY;
      HY:=0.0;
   end
   else begin
      DI:=2.0*WO;
      HY:=HY-WO;
      WO:=0.0;
   end;
end;{proc Diopside}


procedure Riebeckite;
begin
   if NA>F3 then
   begin
      if F2>(1.5*F3) then
      begin
         RI:=F3*7.5;
         SI:=SI-4.0*F3;
         NA:=NA-F3;
         F2:=F2-1.5*F3;
         F3:=0.0;
         NS:=1.5*NA;
         SI:=SI-0.5*NA;
         NA:=0.0;
      end
      else begin
         RI:=5.0*F2;
         SI:=SI-2.666667*F2;
         F3:=F3-0.666667*F2;
         NA:=NA-0.666667*F2;
         F2:=0.0;
         NS:=1.5*NA;
         SI:=SI-0.5*NA;
         NA:=0.0;
      end
   end
   else begin
      if F2>(1.5*NA) then
      begin
         RI:=7.5*NA;
         F3:=F3-NA;
         F2:=F2-1.5*NA;
         SI:=SI-4.0*NA;
         NA:=0.0;
      end
      else begin
         RI:=5.0*F2;
         SI:=SI-2.666667*F2;
         F3:=F3-0.666667*F2;
         NA:=NA-0.666667*F2;
         F2:=0.0;
         NS:=1.5*NA;
         SI:=SI-0.5*NA;
         NA:=0.0;
      end
   end;
end;{proc Riebeckite}


procedure Biotite;
begin
   if FEMG>(0.6*ORT) then
   begin
      BI:=1.6*ORT;
      FEMG:=FEMG-0.6*ORT;
      ORT:=0.0;
   end
   else begin
      BI:=2.666667*FEMG;
      ORT:=ORT-1.666667*FEMG;
      FEMG:=0.0;
   end;
end;{proc Biotite}


procedure ActHyWo;
{
act, hy, wo
}
begin
   if FEMG<(2.5*CA) then
   begin
      ACT:=3.0*FEMG;
      CA:=CA-0.4*FEMG;
      SI:=SI-1.6*FEMG;
      FEMG:=0.0;
      WO:=2.0*CA;
      SI:=SI-CA;
      CA:=0.0;
   end
   else begin
      ACT:=7.5*CA;
      FEMG:=FEMG-2.5*CA;
      SI:=SI-4.0*CA;
      CA:=0.0;
      HY:=2.0*FEMG;
      SI:=SI-FEMG;
      FEMG:=0.0;
   end;
end;{proc ActHyWo}


procedure Hornblende;
{
Compute hornblende
}
begin
   HO:=ACT+ED+RI;
end;{proc Hornblende}


procedure MakeQz;
{
Make Qz
}
begin
   if (SI > 0.0) then QZ:=SI;
   if SI>=0.0 then SI:=0.0;
end;{proc MakeQz}


procedure F2MgRatio;
{
Compute F2/(F2+Mg) for mesonorm
}
begin
   FEMG:=F2+MG;
   if FEMG>0.0 then begin
     PMG:=MG/FEMG;
     PF2:=1.0-PMG;
   end
   else begin
     PMG:=0.0;
     PF2:=0.0;
   end;
   F2:=0.0;
   MG:=0.0;
end;{proc F2MgRatioM}

procedure SetSiZero;
begin
   SI:=0.0;
end;{proc SetSiZero}


end.