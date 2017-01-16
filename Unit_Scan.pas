unit Unit_Scan;

interface
uses
  SysUtils, Classes, StrUtils,Forms;

type
  TBCdata = class         //структура дл€ старого кодировани€
  public
    tp:    byte;
    enp:   string;
    fio:   string;
    fam:   string;
    im:    string;
    ot:    string;
    sex:   byte;
    dr:    tdatetime;
    de:    tdatetime;
    ogrn:  string;
    okato: string;
    SGN:   string;
  end;

  PolStruct = class       //структура дл€ нового кодировани€
  public
    PolNumInt: int64;
    PolNumStr: string;
    Fam:       string;
    Im:        string;
    Otch:      string;
    Sex:       string;
    Dr:        tdatetime;
    DateEnd:   tdatetime;
    OGRN:      string;
    OKATO:     string;
    ECP:       string;
  end;




procedure readdata(s: string; t: TBCdata);
procedure DecodeKode02(KOD: string; PS: PolStruct);

implementation

uses Unit_Main;

function EncodeBase64(const Value: string): string;
const a = ' .-''0123456789јЅ¬√ƒ≈®∆«»… ЋћЌќѕ–—“”‘’÷„Ўў№ЏџЁёя                |';
var
  c: Byte;
  n: Integer;
  Count: Integer;
  DOut: array[0..3] of Byte;
begin
  Result := '';
  Count := 1;
  while Count <= Length(Value) do begin
    c := Ord(Value[Count]);
    Inc(Count);
    DOut[0] := (c and $FC) shr 2;
    DOut[1] := (c and $03) shl 4;
    if Count <= Length(Value) then begin
      c := Ord(Value[Count]);
      Inc(Count);
      DOut[1] := DOut[1] + (c and $F0) shr 4;
      DOut[2] := (c and $0F) shl 2;
      if Count <= Length(Value) then begin
        c := Ord(Value[Count]);
        Inc(Count);
        DOut[2] := DOut[2] + (c and $C0) shr 6;
        DOut[3] := (c and $3F);
      end
      else begin
        DOut[3] := $40;
      end;
    end
    else begin
      DOut[2] := $40;
      DOut[3] := $40;
    end;
    for n := 0 to 3 do
      Result := Result + a[DOut[n] + 1];
  end;
end;

function decodefio(s: string): string;
begin
  result := ReverseString(encodeBase64(ReverseString(s)));
end;

function htoi64(s: string): int64;
begin
  result := pint64(ReverseString(s))^;
end;

function htoi(s: string): integer;
var
  c: string;
  l: integer;
begin
  c := ReverseString(s);
  l := length(c);
  if l = 1 then result := pbyte(c)^
  else if l = 2 then result := pword(c)^
  else if l = 3 then result := pinteger(c)^
  else if l = 4 then result := pinteger(c)^;
end;

procedure readdata(s: string; t: TBCdata);
var i: integer;
  c: string;
begin
  t.tp := htoi(copy(s, 1, 1));
  t.enp := inttostr(htoi64(copy(s, 2, 8)));
  c := decodefio(copy(s, 10, 42));
  t.fio := c;
  t.sex := htoi(copy(s, 52, 1));
  t.dr := strtodate('1.1.1900') + htoi(copy(s, 53, 2));
  t.de := 0;
  if htoi(copy(s, 55, 2)) > 0 then t.de := strtodate('1.1.1900') + htoi(copy(s, 55, 2));
  t.ogrn := inttostr(htoi64(copy(s, 57, 6)));
  t.okato := inttostr(htoi(copy(s, 63, 3)));
  t.SGN := copy(s, 66, 65);
  i := pos('|', c);
  if i > 0 then begin
    t.im := copy(c, 1, i - 1);
    c := copy(c, i + 1, length(c) - i);
    i := pos('|', c);
    if i > 0 then begin
      t.fam := copy(c, 1, i - 1);
      t.ot := copy(c, i + 1, length(c) - i);
    end;
  end;
end;

//прислали дополнительно 05.08.2011
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function HexToInt(HexStr: string): Int64;
var RetVar: Int64;
    i: byte;
begin
  HexStr:=UpperCase(HexStr);
  if HexStr[length(HexStr)]='H' then
     Delete(HexStr,length(HexStr),1);
  RetVar:=0;
  for i:= 1 to length(HexStr) do
    begin
      RetVar:=RetVar shl 4;
      if HexStr[i] in ['0'..'9']
        then RetVar:=RetVar+(byte(HexStr[i])-48)
        else
          if HexStr[i] in ['A'..'F']
            then RetVar:=RetVar+(byte(HexStr[i])-55)
            else
              begin
                Retvar:=0;
                break;
              end;
    end;
  Result:=RetVar;
end;

function HexToStr(HexStr: string): string;
var l: integer;
begin
  l:=Length(HexStr) div 2;
  SetLength(Result,l);
  HextoBin(PChar(HexStr),PChar(Result),l);
end;

//‘‘ изменил кодировку штрих-кода
//------------------------------------------------------------------------------
procedure DecodeKode02(KOD: string; PS: PolStruct);
var
    FIO: string;
    i: integer;
begin
  try
    PS.PolNumInt:=HexToInt(copy(KOD,3,16));
    PS.PolNumStr:=IntToStr(HexToInt(copy(KOD,3,16)));
    //---получаем ‘»ќ
    FIO:=EncodeBase64(copy(HexToStr(KOD),10,42));
    PS.Fam:='ERROR';
    PS.Im:='ERROR';
    PS.Otch:='ERROR';
    i:=Pos('|',FIO);
    if i>0 then
      begin
        PS.Fam:=Trim(copy(FIO,1,i-1));
        FIO:=copy(FIO,i+1,Length(FIO)-i);
        i:=Pos('|',FIO);
        if i>0 then
          begin
            PS.Im:=Trim(copy(FIO,1,i-1));
            PS.Otch:=Trim(copy(FIO,i+1,Length(FIO)-i));
          end;
      end;

    //--------------
    PS.Sex:=IntToStr(HexToInt(copy(KOD,121,2)));
    PS.DR:=HexToInt(copy(KOD,123,4))+2;       //надо +2 сделать, иначе на 2 дн€ меньше дату дает

    PS.DateEnd:=HexToInt(copy(KOD,127,4))+2;  //надо +2 сделать, иначе на 2 дн€ меньше дату дает
    if(PS.DateEnd=2) then
      PS.DateEnd:=0;  //без срока действи€

    PS.OGRN:= '1027739051460';   //не расшифровали ???
    PS.OKATO:='22000';
    PS.ECP:='0';

  except
    PS.PolNumInt:=0;
    PS.PolNumStr:='0';
    PS.Fam:='0';
    PS.Im:='0';
    PS.Otch:='0';
    PS.Sex:='0';
    PS.DR:=0;
    PS.DateEnd:=0;
    PS.OGRN:='0';
    PS.OKATO:='0';
    PS.ECP:='0';
  end;
end; //‘‘ изменил кодировку штрих-кода
//------------------------------------------------------------------------------


end.

