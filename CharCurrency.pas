unit CharCurrency;

//unit Numinwrd;

interface
function sMoneyInWords_1(Nin: currency): string; export;
function szMoneyInWords_1(Nin: currency): PChar; export;

function sMoneyInWords_2(Nin: currency): string; export;
function szMoneyInWords_2(Nin: currency): PChar; export;

{ Денежная сумма Nin в рублях и копейках прописью
1997, в.2.1, by О.В.Болдырев}

implementation
uses SysUtils, Dialogs, Math;

type

  tri = string[4];
  mood = 1..2;
  gender = (m, f);
  uns = array[0..9] of string[7];
  tns = array[0..9] of string[13];
  decs = array[0..9] of string[12];
  huns = array[0..9] of string[10];
  nums = array[0..4] of string[8];
  //money = array[1..2] of string[5];
  money = array[1..1] of string[5];
  endings = array[gender, mood, 1..3] of tri; {окончания числительных и денег}

const

  units: uns = ('', 'один ', 'два ', 'три ', 'четыре ', 'пять ',
    'шесть ', 'семь ', 'восемь ', 'девять ');
  unitsf: uns = ('', 'одна ', 'две ', 'три ', 'четыре ', 'пять ',
    'шесть ', 'семь ', 'восемь ', 'девять ');
  teens: tns = ('десять ', 'одиннадцать ', 'двенадцать ', 'тринадцать ',
    'четырнадцать ', 'пятнадцать ', 'шестнадцать ',
    'семнадцать ', 'восемнадцать ', 'девятнадцать ');
  decades: decs = ('', 'десять ', 'двадцать ', 'тридцать ', 'сорок ',
    'пятьдесят ', 'шестьдесят ', 'семьдесят ', 'восемьдесят ',
    'девяносто ');
  hundreds: huns = ('', 'сто ', 'двести ', 'триста ', 'четыреста ',
    'пятьсот ', 'шестьсот ', 'семьсот ', 'восемьсот ',
    'девятьсот ');
  numericals: nums = ('', 'тысяч', 'миллион', 'миллиард', 'триллион');
  //RusMon: money = ('рубл', 'копе');
  RusMon: money = ('шту');
  //ends: endings = ((('', 'а', 'ов'), ('ь', 'я', 'ей')), (('а', 'и', ''),  ('йка', 'йки', 'ек')));
  ends: endings = ((('', 'а', 'ов'), ('ка', 'ки', 'к')), (('а', 'и', ''),  ('йка', 'йки', 'ек')));
threadvar

  str: string;

function EndingIndex(Arg: integer): integer;
begin

  if ((Arg div 10) mod 10) <> 1 then
    case (Arg mod 10) of
      1: Result := 1;
      2..4: Result := 2;
    else
      Result := 3;
    end
  else
    Result := 3;
end;

function sMoneyInWords_1(Nin: currency): string;
  { Число Nin прописью, как функция }
var
  //  str: string;

  g: gender; //род
  Nr: comp; {целая часть числа}
  Fr: integer; {дробная часть числа}
  i, iTri, Order: longint; {триада}

  procedure Triad;
  var
    iTri2: integer;
    un, de, ce: byte; //единицы, десятки, сотни

    function GetDigit: byte;
    begin
      Result := iTri2 mod 10;
      iTri2 := iTri2 div 10;
    end;

  begin
    iTri := trunc(Nr / IntPower(1000, i));
    Nr := Nr - int(iTri * IntPower(1000, i));
    iTri2 := iTri;
    if iTri > 0 then
    begin
      un := GetDigit;
      de := GetDigit;
      ce := GetDigit;
      if i = 1 then
        g := f
      else
        g := m; {женского рода только тысяча}

      str := TrimRight(str) + ' ' + Hundreds[ce];
      if de = 1 then
        str := TrimRight(str) + ' ' + Teens[un]
      else
      begin
        str := TrimRight(str) + ' ' + Decades[de];
        case g of
          m: str := TrimRight(str) + ' ' + Units[un];
          f: str := TrimRight(str) + ' ' + UnitsF[un];
        end;
      end;

      if length(numericals[i]) > 1 then
      begin
        str := TrimRight(str) + ' ' + numericals[i];
        str := TrimRight(str) + ends[g, 1, EndingIndex(iTri)];
      end;
    end; //triad is 0 ?

    if i = 0 then
      Exit;
    Dec(i);
    Triad;
  end;

begin

  str := '';
  Nr := int(Nin);
  Fr := round(Nin * 100 + 0.00000001) mod 100;
  if Nr > 0 then
    Order := trunc(Log10(Nr) / 3)
  else
  begin
    str := 'ноль';
    Order := 0
  end;
  if Order > High(numericals) then
    raise Exception.Create('Слишком большое число для суммы прописью');
  i := Order;
  Triad;

  //str := Format('%s %s%s %.2d %s%s', [Trim(str), RusMon[1], ends[m, 2, EndingIndex(iTri)], Fr, RusMon[2], ends[f, 2, EndingIndex(Fr)]]);
  str := Format('(%s) %s%s', [Trim(str), RusMon[1], ends[m, 2, EndingIndex(iTri)] ]);

  //str[1] := (ANSIUpperCase(copy(str, 1, 1)))[1];
  str[2] := (ANSIUpperCase(copy(str, 2, 1)))[1];

  str[Length(str) + 1] := #0;
  Result := str;
end;

function sMoneyInWords_2(Nin: currency): string;
  { Число Nin прописью, как функция }
var
  //  str: string;

  g: gender; //род
  Nr: comp; {целая часть числа}
  Fr: integer; {дробная часть числа}
  i, iTri, Order: longint; {триада}

  procedure Triad;
  var
    iTri2: integer;
    un, de, ce: byte; //единицы, десятки, сотни

    function GetDigit: byte;
    begin
      Result := iTri2 mod 10;
      iTri2 := iTri2 div 10;
    end;

  begin
    iTri := trunc(Nr / IntPower(1000, i));
    Nr := Nr - int(iTri * IntPower(1000, i));
    iTri2 := iTri;
    if iTri > 0 then
    begin
      un := GetDigit;
      de := GetDigit;
      ce := GetDigit;
      if i = 1 then
        g := f
      else
        g := m; {женского рода только тысяча}

      str := TrimRight(str) + ' ' + Hundreds[ce];
      if de = 1 then
        str := TrimRight(str) + ' ' + Teens[un]
      else
      begin
        str := TrimRight(str) + ' ' + Decades[de];
        case g of
          m: str := TrimRight(str) + ' ' + Units[un];
          f: str := TrimRight(str) + ' ' + UnitsF[un];
        end;
      end;

      if length(numericals[i]) > 1 then
      begin
        str := TrimRight(str) + ' ' + numericals[i];
        str := TrimRight(str) + ends[g, 1, EndingIndex(iTri)];
      end;
    end; //triad is 0 ?

    if i = 0 then
      Exit;
    Dec(i);
    Triad;
  end;

begin

  str := '';
  Nr := int(Nin);
  Fr := round(Nin * 100 + 0.00000001) mod 100;
  if Nr > 0 then
    Order := trunc(Log10(Nr) / 3)
  else
  begin
    str := 'ноль';
    Order := 0
  end;
  if Order > High(numericals) then
    raise Exception.Create('Слишком большое число для суммы прописью');
  i := Order;
  Triad;

  //str := Format('%s %s%s %.2d %s%s', [Trim(str), RusMon[1], ends[m, 2, EndingIndex(iTri)], Fr, RusMon[2], ends[f, 2, EndingIndex(Fr)]]);
  str := Format('(%s)', [Trim(str)]);

  //str[1] := (ANSIUpperCase(copy(str, 1, 1)))[1];
  str[2] := (ANSIUpperCase(copy(str, 2, 1)))[1];

  str[Length(str) + 1] := #0;
  Result := str;
end;

function szMoneyInWords_1(Nin: currency): PChar;
begin

  sMoneyInWords_1(Nin);
  Result := @(str[1]);
end;

function szMoneyInWords_2(Nin: currency): PChar;
begin

  sMoneyInWords_2(Nin);
  Result := @(str[1]);
end;

end.

