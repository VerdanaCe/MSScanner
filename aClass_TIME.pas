unit aClass_TIME;

interface

uses
  SysUtils;


//1. �������� TimeMetka
function GetTimeMetka(aPresent: TDateTime; aIdVid: integer): string;


implementation

//1. �������� TimeMetka
{
  1: ��������, 09:17:56
  2: ��������, 21 �������� 2013 ����, �������
  3: ��������, 21.09.2013 09:17:56
  4: ��������, 2013_09_21_09_17_56
  5: ��������, 21 �������� 2013 ����
  6: ��������, 21.09.2013
}
//------------------------------------------------------------------------------
function GetTimeMetka(aPresent: TDateTime; aIdVid: integer): string;
var
  Year, Month, Day, Hour, Min, Sec, MSec : Word;    // ���, �����, �����, ����, ������, �������, �����������

  Res: string;
begin

  {������������ FormatDateTime:
    ���:          yy    00-99
                  yyyy  0000-9999
    �����:        m     1-12
                  mm    01-12
    ����:         d     1-31
                  dd    01-31
    ���:          h     0-23
                  hh    00-23
    ������:       n     0-59
                  nn    00-59
    �������:      s     0-59
                  ss    00-59
    �����������:  z     0-999
                  zzz   000-999
  }

  DecodeDate(aPresent, Year, Month, Day);
  DecodeTime(aPresent, Hour, Min, Sec, MSec);

  case aIdVid of
  1:  //��������, 09:17:56
    Res:= FormatDateTime('hh:nn:ss',aPresent);
  //2:  //��������, 21 �������� 2013 ����, �������
    //Res:=IntToStr(Day) + ' ' +  stMonth[Month] + ' ' + IntToStr(Year) + ' ����, ' + stDay[DayOfWeek(aPresent)];
  3:  //��������, 21.09.2013 09:17:56
    Res:= FormatDateTime('dd.mm.yyyy hh:nn:ss',aPresent);
  4:
    Res:= FormatDateTime('yyyy_mm_dd_hh_nn_ss',aPresent);
  //5:
    //Res:=IntToStr(Day) + ' ' +  stMonth[Month] + ' ' + IntToStr(Year) + ' ����';
  6:
    Res:= FormatDateTime('dd.mm.yyyy',aPresent);
  else
    Res:='0000_00_00_00_00_00'; //��������� �� ���������
  end;

  result:=Res;
end; //1. �������� TimeMetka
//------------------------------------------------------------------------------

end.