unit aClass_TIME;

interface

uses
  SysUtils;


//1. получить TimeMetka
function GetTimeMetka(aPresent: TDateTime; aIdVid: integer): string;


implementation

//1. получить TimeMetka
{
  1: например, 09:17:56
  2: например, 21 сент€бр€ 2013 года, суббота
  3: например, 21.09.2013 09:17:56
  4: например, 2013_09_21_09_17_56
  5: например, 21 сент€бр€ 2013 года
  6: например, 21.09.2013
}
//------------------------------------------------------------------------------
function GetTimeMetka(aPresent: TDateTime; aIdVid: integer): string;
var
  Year, Month, Day, Hour, Min, Sec, MSec : Word;    // год, мес€ц, число, часы, минуты, секунды, милисекунды

  Res: string;
begin

  {спецификаци€ FormatDateTime:
    год:          yy    00-99
                  yyyy  0000-9999
    мес€ц:        m     1-12
                  mm    01-12
    день:         d     1-31
                  dd    01-31
    час:          h     0-23
                  hh    00-23
    минуты:       n     0-59
                  nn    00-59
    секунды:      s     0-59
                  ss    00-59
    милисекунды:  z     0-999
                  zzz   000-999
  }

  DecodeDate(aPresent, Year, Month, Day);
  DecodeTime(aPresent, Hour, Min, Sec, MSec);

  case aIdVid of
  1:  //например, 09:17:56
    Res:= FormatDateTime('hh:nn:ss',aPresent);
  //2:  //например, 21 сент€бр€ 2013 года, суббота
    //Res:=IntToStr(Day) + ' ' +  stMonth[Month] + ' ' + IntToStr(Year) + ' года, ' + stDay[DayOfWeek(aPresent)];
  3:  //например, 21.09.2013 09:17:56
    Res:= FormatDateTime('dd.mm.yyyy hh:nn:ss',aPresent);
  4:
    Res:= FormatDateTime('yyyy_mm_dd_hh_nn_ss',aPresent);
  //5:
    //Res:=IntToStr(Day) + ' ' +  stMonth[Month] + ' ' + IntToStr(Year) + ' года';
  6:
    Res:= FormatDateTime('dd.mm.yyyy',aPresent);
  else
    Res:='0000_00_00_00_00_00'; //“аймћетка по умолчанию
  end;

  result:=Res;
end; //1. получить TimeMetka
//------------------------------------------------------------------------------

end.