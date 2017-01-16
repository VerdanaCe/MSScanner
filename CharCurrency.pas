unit CharCurrency;

//unit Numinwrd;

interface
function sMoneyInWords_1(Nin: currency): string; export;
function szMoneyInWords_1(Nin: currency): PChar; export;

function sMoneyInWords_2(Nin: currency): string; export;
function szMoneyInWords_2(Nin: currency): PChar; export;

{ �������� ����� Nin � ������ � �������� ��������
1997, �.2.1, by �.�.��������}

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
  endings = array[gender, mood, 1..3] of tri; {��������� ������������ � �����}

const

  units: uns = ('', '���� ', '��� ', '��� ', '������ ', '���� ',
    '����� ', '���� ', '������ ', '������ ');
  unitsf: uns = ('', '���� ', '��� ', '��� ', '������ ', '���� ',
    '����� ', '���� ', '������ ', '������ ');
  teens: tns = ('������ ', '����������� ', '���������� ', '���������� ',
    '������������ ', '���������� ', '����������� ',
    '���������� ', '������������ ', '������������ ');
  decades: decs = ('', '������ ', '�������� ', '�������� ', '����� ',
    '��������� ', '���������� ', '��������� ', '����������� ',
    '��������� ');
  hundreds: huns = ('', '��� ', '������ ', '������ ', '��������� ',
    '������� ', '�������� ', '������� ', '��������� ',
    '��������� ');
  numericals: nums = ('', '�����', '�������', '��������', '��������');
  //RusMon: money = ('����', '����');
  RusMon: money = ('���');
  //ends: endings = ((('', '�', '��'), ('�', '�', '��')), (('�', '�', ''),  ('���', '���', '��')));
  ends: endings = ((('', '�', '��'), ('��', '��', '�')), (('�', '�', ''),  ('���', '���', '��')));
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
  { ����� Nin ��������, ��� ������� }
var
  //  str: string;

  g: gender; //���
  Nr: comp; {����� ����� �����}
  Fr: integer; {������� ����� �����}
  i, iTri, Order: longint; {������}

  procedure Triad;
  var
    iTri2: integer;
    un, de, ce: byte; //�������, �������, �����

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
        g := m; {�������� ���� ������ ������}

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
    str := '����';
    Order := 0
  end;
  if Order > High(numericals) then
    raise Exception.Create('������� ������� ����� ��� ����� ��������');
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
  { ����� Nin ��������, ��� ������� }
var
  //  str: string;

  g: gender; //���
  Nr: comp; {����� ����� �����}
  Fr: integer; {������� ����� �����}
  i, iTri, Order: longint; {������}

  procedure Triad;
  var
    iTri2: integer;
    un, de, ce: byte; //�������, �������, �����

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
        g := m; {�������� ���� ������ ������}

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
    str := '����';
    Order := 0
  end;
  if Order > High(numericals) then
    raise Exception.Create('������� ������� ����� ��� ����� ��������');
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

