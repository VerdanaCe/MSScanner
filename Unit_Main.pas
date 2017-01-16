unit Unit_Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DBGridEhGrouping, ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh,
  StdCtrls, Mask,ComObj, DBCtrlsEh, EhLibVCL, GridsEh, DBAxisGridsEh, DBGridEh,
  LMDControl, LMDCustomControl, LMDCustomPanel, LMDCustomBevelPanel,
  LMDBaseEdit, LMDCustomEdit, LMDEdit, CPort, LMDCustomParentPanel,
  LMDBackPanel, LMDCustomPanelFill, LMDButtonPanel, LMDCustomToolBar,
  LMDToolBar, DBVertGridsEh, LMDThemedComboBox, LMDCustomComboBox,
  LMDCustomColorComboBox, LMDColorComboBox, LMDBaseControl,
  LMDBaseGraphicControl, LMDBaseLabel, LMDCustomLabel, LMDLabel,
  LMDCustomMemo, LMDMemo, CPortCtl, DB, ADODB, LMDCustomButton,
  LMDDockButton, sSkinProvider, sSkinManager, AdvEdit, AdvEdBtn, ExtCtrls,
  sPanel, LMDCustomStatusBar, LMDStatusBar, AdvGlowButton, AdvGlassButton,
  GradientLabel, sEdit, FolderDialog, sMemo, sLabel, sDBText, Grids,
  DBGrids, acDBGrid, sComboBoxes, Buttons, sBitBtn, ComCtrls, sStatusBar,
  sMaskEdit, sCustomComboEdit, sComboEdit, AdvDBLookupComboBox, sButton,
  sBevel, frxClass, frxExportXLS,nExcel;

type
  TScanPolis = class(TForm)
    ComPort: TComPort;
    count_rd_data: TLMDLabel;
    ComLed1: TComLed;
    ADOConnect: TADOConnection;
    Q1: TADOQuery;
    Q6: TADOQuery;
    Q11: TADOQuery;
    DataSource2: TDataSource;
    DataSource1: TDataSource;
    Q6person_id: TIntegerField;
    Q6filial_id: TIntegerField;
    Q6last_name: TStringField;
    Q6first_name: TStringField;
    Q6second_name: TStringField;
    Q6date_birth: TDateTimeField;
    sSkinManager1: TsSkinManager;
    sSkinProvider1: TsSkinProvider;
    Q3: TADOQuery;
    Q3filial_name: TWideStringField;
    pr_connect_label: TGradientLabel;
    lb_info: TGradientLabel;
    lb_count: TGradientLabel;
    e_number_table: TsEdit;
    PersonGrid: TsDBGrid;
    PolisGrid: TsDBGrid;
    MemoInfo: TsMemo;
    MemoData: TsMemo;
    SQLMemo: TsMemo;
    part_number: TsComboEdit;
    StatusBar: TsStatusBar;
    sPanel1: TsPanel;
    btnConnScan: TsBitBtn;
    btnSetupScan: TsBitBtn;
    btnOpenPart: TsBitBtn;
    D_Info: TsComboBoxEx;
    TimePanel: TsLabelFX;
    sPanel2: TsPanel;
    fam_label: TsDBText;
    im_label: TsDBText;
    ot_label: TsDBText;
    filial_id_edit: TsDBText;
    enp_label: TsDBText;
    dr_label: TsDBText;
    enp_stop: TsDBText;
    info_label_edit: TsDBText;
    filial_name_label: TsDBText;
    svid_date_stop_edit: TsDBText;
    svid_date_start_edit: TsDBText;
    svid_label: TsDBText;
    sStickyLabel1: TsStickyLabel;
    sStickyLabel2: TsStickyLabel;
    sStickyLabel3: TsStickyLabel;
    sStickyLabel4: TsStickyLabel;
    sStickyLabel5: TsStickyLabel;
    sStickyLabel6: TsStickyLabel;
    DataSource3: TDataSource;
    sStickyLabel7: TsStickyLabel;
    sStickyLabel8: TsStickyLabel;
    sStickyLabel9: TsStickyLabel;
    sStickyLabel10: TsStickyLabel;
    DataSource4: TDataSource;
    Q7: TADOQuery;
    statENP: TsButton;
    OD1: TOpenDialog;
    loadBlank: TsButton;
    btnClosePart: TsBitBtn;
    ENP_DUBL: TADOStoredProc;
    PartGrid1: TDBGridEh;
    Q20: TADOQuery;
    sLabel1: TsLabel;
    sBevel1: TsBevel;
    sBevel3: TsBevel;
    sBevel4: TsBevel;
    frxXLSExport1: TfrxXLSExport;
    PartGrid2: TDBGridEh;
    btnScanKolvo1: TsButton;
    DataSource5: TDataSource;
    QQ19: TADOQuery;

    procedure ClearInfoCaption();    //очистка Label на Форме
    procedure ClearInfoQ();          //очистка запроса Q6 на форме
    procedure ClearInfoCountRD();    //очистить информацию кол-ва считанных данных из порта на форме
    procedure ClearInfoNumberPart();
    procedure ComPortAfterOpen(Sender: TObject);
    procedure D_InfoChange(Sender: TObject);
    procedure ComPortRxChar(Sender: TObject; Count: Integer);
    procedure ComPortAfterClose(Sender: TObject);
    procedure part_numberClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnConnScanClick(Sender: TObject);
    procedure btnSetupScanClick(Sender: TObject);
    procedure btnOpenPartClick(Sender: TObject);
    procedure part_numberButtonClick(Sender: TObject);
    procedure statENPClick(Sender: TObject);
    procedure loadBlankClick(Sender: TObject);
    procedure btnClosePartClick(Sender: TObject);
    procedure btnScanKolvo1Click(Sender: TObject);

   
   


  private
   DLINA: integer;
   InitialDir: string;          //каталог с Excel файлами из которых закачиваем данные в т.ENP

  public
    kolvo: integer;
    part_count: integer;
    part_active:string;
  end;

var
  ScanPolis: TScanPolis;
  CatalogExe:             string;
  CatalogShablon:         string;
  CatalogResultOMSPart:   string;
  CatalogResultOMSAkt:    string;
  CatalogResultTXOAkt:    string;
  CatalogResultSMS_XLS:   string;
  CatalogResultSMS_CSV:   string;
  CatalogResultSMS_XLS_C: string;
  CatalogResultDelo_XTO:  string;
  CatalogResultDelo_OMS:  string;
  CatalogResultDolg:      string;
  Catalog_SMS:            string;
  Catalog_VipNet_SMS:     string;
  Catalog_PoiskTel:       string;
  Catalog_Statistic_SMS:  string;
  TIME_TM:                string;  //например, 2013_09_26

  InitialDir_Contact:     string;  //каталог с Excel-CSV файлами из которых закачиваем данный о контактах
  InitialDir_SMS:         string;  //каталог с Excel-CSV файлом sent_messages.csv с данными о доставке
  InitialDir_SMS_Megafon: string;  //каталог с Excel файлом messages.xlsx с данными о доставке
  Separator1:             string;  //разделитель ; для считывания данных из CSV-файла экспорта контактов
  Separator2:             string;  //разделитель , для считывания данных из TXT-файла экспорта контактов
 const
  stCIFRA ='0123456789'; //набор цифр

implementation

uses Unit_Scan;
const
  stDay : array[1..7] of string[11] =
    ('воскресенье','понедельник','вторник',
     'среда','четверг','пятница','суббота');

  stMonth : array[1..12] of string[8] =
    ('января','февраля','марта',
     'апреля','мая','июня','июля',
     'августа','сентября','октября',
     'ноября','декабря');

  stMonth1 : array[1..12] of string[8] =
    ('январь','февраль','март',
     'апрель','май','июнь','июль',
     'август','сентябрь','октябрь',
     'ноябрь','декабрь');
{$R *.dfm}
procedure makedir(value:string);
var i,x:integer;
    cur_dir:string;
    RootDir:String;
begin
  RootDir:=value[1]+value[2]+value[3];
  SetCurrentDirectory(pchar(RootDir));
  x:=1;
  cur_dir:='';
  if (value[1]='\') then x:=2;
  for i:=x to Length(value) do
   begin
    if not (value[i]='\')then
      cur_dir:=cur_dir+value[i];
    if (value[i]='\')or (i=length(value)) then
     begin
      if not DirectoryExists(cur_dir) then
       CreateDirectory(pchar(cur_dir),0);
      SetCurrentDirectory(pchar(cur_dir));
      cur_dir:='';
     end;
   end;
end; //создать каталог
//1.1. работа с преобразованияем Hex - Str - Hex
//------------------------------------------------------------------------------
function HexToStr(s: string): string;
var
  i: integer;
  l: integer;
  ss: string;
  tmp: string;
begin
  ss:= '';
  l:= trunc(length(s)/2);
  for i:=0 to l-1 do
  begin
    tmp:= copy(s,1,2);
    delete(s,1,2);
    ss:= ss+ char(strtoint('$'+tmp));
  end;
  result:= ss;
end;

function StrToHex(s: string): string;
var
  i: integer;
  l: integer;
  ss: string;
  tmp: string[1];
begin
  ss:= '';
  l:= length(s);
  for i:=0 to l-1 do
  begin
    tmp:= copy(s,1,1);
    delete(s,1,1);
    ss:= ss+ inttohex(ord(tmp[1]),2);
  end;
  result:= ss;
end;
//1.1. работа с преобразованияем Hex - Str - Hex
//------------------------------------------------------------------------------
//2. ФОРМА
//******************************************************************************

//2.1.1. очистить информацию Label на форме
//------------------------------------------------------------------------------
procedure TScanPolis.ClearInfoCaption();
begin
  filial_id_edit.Caption:='';
  filial_name_label.Caption:='';
  enp_label.Caption:='';
  fam_label.Caption:='';
  im_label.Caption:='';
  dr_label.Caption:='';
  enp_label.Caption:='';
  ot_label.Caption:='';
  svid_label.Caption:= '';
  svid_date_start_edit.Caption:= '';
  svid_date_stop_edit.Caption:= '';

  info_label_edit.Caption:='';


end; //2.1.1. очистить информацию Label на форме
//------------------------------------------------------------------------------

//2.1.2. очистить информацию Q6 на форме
//------------------------------------------------------------------------------
procedure TScanPolis.ClearInfoQ();
begin
  SQLMemo.Clear;
 Q1.Close;
 Q3.Close;
 Q6.Close;
end; //2.1.2. очистить информацию Q6 на форме
//------------------------------------------------------------------------------

//2.1.3. очистить информацию кол-ва считанных данных из порта
//------------------------------------------------------------------------------
procedure TScanPolis.ClearInfoCountRD();
begin
  count_rd_data.Caption:='_';
end; //2.1.3. очистить информацию кол-ва считанных данных из порта
//------------------------------------------------------------------------------

//2.1.4. очистить информацию № партии
//------------------------------------------------------------------------------
procedure TScanPolis.ClearInfoNumberPart();
begin
  part_number.Text:='';
end; //2.1.4. очистить информацию № партии
//------------------------------------------------------------------------------

procedure TScanPolis.btnConnScanClick(Sender: TObject);
begin
  if ComPort.Connected then
    ComPort.Close
  else
    begin
    ComPort.Open;
    end;
end;
procedure TScanPolis.btnSetupScanClick(Sender: TObject);
begin
 ComPort.ShowSetupDialog;
end;


procedure TScanPolis.ComPortAfterOpen(Sender: TObject);
begin
  btnConnScan.Caption := 'Отключить';
end;

procedure TScanPolis.ComPortAfterClose(Sender: TObject);
begin
 if btnConnScan <> nil then
    btnConnScan.Caption := 'Подключить';
end;

procedure TScanPolis.D_InfoChange(Sender: TObject);
begin
 DLINA:=StrToInt(D_Info.Text);
end;

procedure TScanPolis.ComPortRxChar(Sender: TObject; Count: Integer);
var
  Str: String;

  S: String;
  t: TBCdata;       //структура старого кодирования
  t02: PolStruct;   //структура нового кодирования
  err_id: integer;
  enp: string;
  pr_id: integer;
  fil_id: integer;
  scan_date: string;
  scan_date_time: string; //[19]
  part: string;           //[16]
begin
  ComPort.ReadStr(S, Count);                      //считываем порт как стринг
    kolvo:=kolvo+Count;                           //кол-во считанных сканером данных
    count_rd_data.Caption:=Format('%d',[kolvo]);  //кол-во считанных данных на экран
  MemoData.Text := MemoData.Text + StrToHex(S);   //данные из com-порта
  Application.ProcessMessages;
  DLINA:=132;

  if Length(part_number.Text)<>0 then
    begin
      if kolvo=DLINA then
        begin //конец блока данных
          try
          kolvo:=0;
          MemoInfo.Clear;

            t := TBCdata.create;

            //определим вид кодировки
            str:=copy(MemoData.Text,1,2);
            if str='01' then
              begin                          //старый вид кодировки
                s := hextostr(MemoData.Text);
                readdata(s, t);
              end
            else
              if str='02' then               //новый вид кодировки
                begin
                  //преобразуем структуры
                  t02:=PolStruct.Create;
                    DecodeKode02(MemoData.Text, t02);
                    t.enp:=t02.PolNumStr;
                    t.fam:=t02.Fam;
                    t.im:=t02.Im;
                    t.ot:=t02.Otch;
                    t.sex:=StrToInt(t02.Sex);
                    t.dr:=t02.Dr;
                    t.de:=t02.DateEnd;
                    t.ogrn:=t02.OGRN;
                    t.okato:=t02.OKATO;
                    t.SGN:=t02.ECP;
                  t02.Free;
                end
              else
                begin                        //неизвестный вид кодировки
                  t.tp:=   0;
                  t.enp:=  '_';
                  t.fio:=  '_';
                  t.fam:=  '_';
                  t.im:=   '_';
                  t.ot:=   '_';
                  t.sex:=  0;
                  t.dr:=   0;
                  t.de:=strtodate('1.1.1900');
                  t.ogrn:= '_';
                  t.okato:='_';
                  t.SGN:=  '_';
                end;
            ClearInfoQ();
            MemoInfo.clear;
            MemoInfo.lines.add('ЕНП   - ' + Trim(t.enp));
            MemoInfo.lines.add('');
            MemoInfo.lines.add('ФАМ   - ' + Trim(t.fam));
            MemoInfo.lines.add('ИМ    - ' + Trim(t.im));
            MemoInfo.lines.add('ОТ    - ' + Trim(t.ot));
            MemoInfo.lines.add('ПОЛ   - ' + inttostr(t.sex));

            if(datetostr(t.dr)='30.12.1899') then
              MemoInfo.lines.add('ДР    - ')
            else
              MemoInfo.lines.add('ДР    - ' + datetostr(t.dr));

            if(datetostr(t.de)='30.12.1899') then
              MemoInfo.lines.add('СРОК ДЕЙСТВИЯ - ')
            else
              MemoInfo.lines.add('СРОК ДЕЙСТВИЯ - ' + datetostr(t.de));

            MemoInfo.lines.add('ОГРН  - ' + Trim(t.ogrn));
            MemoInfo.lines.add('ОКАТО - ' + Trim(t.okato));
           enp:=Trim(t.enp); SQLMemo.Lines.Add(enp);
           if Length(enp)<16 then begin enp:='0'+ Trim(t.enp); SQLMemo.Lines.Add(enp); end;
            //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
           scan_date_time:=DateTimeToStr(Now);

             Q6.Close;
             Q6.SQL.Clear;  //ищем в БД по ФИО и ДР.
              str:='select person_id,filial_id,last_name,first_name,second_name,date_birth ';     Q6.SQL.Add(str);
              str:='from PERSONA where ';                                               Q6.SQL.Add(str);
              str:=Format('last_name=''%s'' and ',[Trim(t.fam)]);                          Q6.SQL.Add(str);
              str:=Format('first_name=''%s'' and ',[Trim(t.im)]);                          Q6.SQL.Add(str);
              str:=Format('second_name=''%s'' and ',[Trim(t.ot)]);                         Q6.SQL.Add(str);
              str:=Format('date_birth=''%s'' ',[datetostr(t.dr)]);                                 Q6.SQL.Add(str);
             Q6.Open;
              if  Q6.RecordCount <> 0 then
               begin
                   if  Q6.RecordCount = 1 then
                   begin
                    pr_id:= Q6.FieldByName('person_id').AsInteger;
                    Q1.Close;
                    Q1.SQL.Clear;
                     str:='select top 1 PERSON_ALL_ID, POLIS_ID, POLIS_START_DATE, POLIS_STOP_DATE, POLIS_NUMBER, '; Q1.SQL.Add(str); SQLMemo.Lines.Add(str);
                     str:='ENP, VPOLIS, COUNT_POL  from POLIS_ALL where VPOLIS=2 and ';               Q1.SQL.Add(str);SQLMemo.Lines.Add(str);
                     str:=Format('PERSON_ALL_ID =%d',[pr_id]);                                      Q1.SQL.Add(str);SQLMemo.Lines.Add(str);
                     str:='order by POLIS_START_DATE desc';                                                      Q1.SQL.Add(str);SQLMemo.Lines.Add(str);
                    Q1.Open;
                     fil_id:= Q6.FieldByName('filial_id').AsInteger;
                     part:=part_number.Text;
                    enp_label.Caption:= enp;

                    if(datetostr(t.de)='30.12.1899') then
                      enp_stop.Caption:='без срока действия'
                    else
                      enp_stop.Caption:='действует до '+datetostr(t.de);
                    Q3.Close;
                    Q3.SQL.Clear;
                      str := Format('select filial_name from filial where filial_id=%d',[fil_id]); Q3.SQL.Add(str);
                    Q3.Open;
                 //   filial_name_label.Text:=Q3.FieldByName('filial_name').AsString;
                    
                    if ((Length(part)=0) or (part=part_number.Text)) then
                      begin
                        scan_date:= DateTostr(Date);

                        Q11.Close;
                        Q11.SQL.Clear;
                         str := Format('update blank set PERSON_ID=%d,',[pr_id]);            Q11.SQL.Add(str);  SQLMemo.Lines.Add(str);
                         str := Format('FILIAL_ID=%d,',[fil_id]);                            Q11.SQL.Add(str);  SQLMemo.Lines.Add(str);
                         str := Format('last_name=''%s'',',[Trim(t.fam)]);                       Q11.SQL.Add(str);  SQLMemo.Lines.Add(str);
                         str := Format('first_name=''%s'',',[Trim(t.im)]);                       Q11.SQL.Add(str); SQLMemo.Lines.Add(str);
                         str := Format('second_name=''%s'',',[Trim(t.ot)]);                      Q11.SQL.Add(str);  SQLMemo.Lines.Add(str);

                         if(datetostr(t.dr)='30.12.1899') or (Length(datetostr(t.dr))=0) then
                          str:= 'date_birth=null,'
                         else
                          str:=Format('date_birth=''%s'',',[datetostr(t.dr)]);
                         Q11.SQL.Add(str);  SQLMemo.Lines.Add(str);

                        // if(datetostr(t.de)='30.12.1899') or (Length(datetostr(t.de))=0) then
                          //str:='enp_stop_date=null,'
                          // else
                        // str:=Format('enp_stop_date=''%s'',',[datetostr(t.de)]);
                        // Q11.SQL.Add(str);

                         str:=Format('svid=''%s'',',[svid_label.Caption]);                    Q11.SQL.Add(str);    SQLMemo.Lines.Add(str);
                         str:=Format('svid_start_date=''%s'',',[svid_date_start_edit.Caption]);     Q11.SQL.Add(str);   SQLMemo.Lines.Add(str);
                         if Length(svid_date_stop_edit.Caption)=0 then
                          begin str:='svid_stop_date=null,'; Q11.SQL.Add(str);   SQLMemo.Lines.Add(str); end
                          else begin str:=Format('svid_stop_date=''%s'',',[svid_date_stop_edit.Caption]);       Q11.SQL.Add(str);   SQLMemo.Lines.Add(str); end;

                         str:=Format('scan_date=''%s'',',[scan_date]);                           Q11.SQL.Add(str);  SQLMemo.Lines.Add(str);
                         str:=Format('part=''%s'',',[Trim(part_number.Text)]);        Q11.SQL.Add(str);    SQLMemo.Lines.Add(str);
                         str:='MARK=1 where ';                                        Q11.SQL.Add(str);  SQLMemo.Lines.Add(str);
                         str:=Format('NAME_FILE=''%s'' and ',[Trim(part_number.Text)]);        Q11.SQL.Add(str);    SQLMemo.Lines.Add(str);
                         str:=Format('enp=''%s'';',[enp]);                                 Q11.SQL.Add(str); SQLMemo.Lines.Add(str);

                        Q11.ExecSQL;

                        end;
                   end
                 else
                   begin
                    ENP_DUBL.Active:=false;
                    ENP_DUBL.Parameters[1].Value:=pr_id;
                    ENP_DUBL.Parameters[2].Value:=Trim(t.enp);
                    if ENP_DUBL.Parameters[0].Value <0 then
                    begin
                      err_id:=ENP_DUBL.Parameters[0].Value;
                     StatusBar.Panels[1].Text:='';
                    end
                   end;


                   if  Q6.RecordCount > 1 then
                    begin
                     StatusBar.Panels[1].Text:='Дубль';
                    end;
               end
              else
               begin
               //не найден в списке
                  ClearInfoCaption();   //очистка данных на форме
                  ClearInfoQ();
                 ClearInfoCountRD();
                 MemoData.Clear;  //цифровые данные com-порта удаляем
                info_label_edit.Caption:='НЕ НАЙДЕН В СПИСКЕ !!! Попробуйте повторить ввод данных!';
                enp_stop.Caption:='';
                 Application.ProcessMessages;
                kolvo:=0;
                end;
            t.free;
            MemoData.Clear;
          except
            MemoData.Clear;
            MemoInfo.Clear;
            ClearInfoCaption();   //очистка данных на форме
            ClearInfoQ();
            ClearInfoCountRD();
            info_label_edit.Caption:='ОШИБКА! Попробуйте повторить ввод данных!';
            enp_stop.Caption:='';
            Application.ProcessMessages;
            kolvo:=0;
               end;


         end; //kol=dlina
        end
  else
    begin //if Length(part_number.Text)=0
      MemoData.Clear;
      MemoInfo.Clear;
      ClearInfoCaption();   //очистка данных на форме
      ClearInfoQ();
      ClearInfoCountRD();
      info_label_edit.Caption:='ОШИБКА! Не указан № партии полисов! Укажите № партии.';
      enp_stop.Caption:='';
      Application.ProcessMessages;
      kolvo:=0;
    end;  //if Length(part_number.Text)=0
end;



procedure TScanPolis.part_numberClick(Sender: TObject);
begin
  part_number.Text:='637_29_11_2016';
end;

procedure TScanPolis.FormCreate(Sender: TObject);
var
str: string;

 Present: TDateTime;        // текущая дата и время
 Year, Month, Day : Word;   // год, месяц и число, как отдельные числа
 pr_conn: boolean;
 //IniFileName : TIniFile;   //файл настроек
// dlina_index: integer;     //длина пакета (номер строки)
begin
 Present:= Now;            // получить текущую дату
 DecodeDate(Present, Year, Month, Day);
 TimePanel.Caption := IntToStr(Day)+ ' ' +  stMonth[Month] + ' '+ IntToStr(Year)+  ' года, '+ stDay[DayOfWeek(Present)];

  TIME_TM:=FormatDateTime('yyyy_mm_dd',Present);   //например, 2013_09_26

  //определим каталог *.exe
  GetDir(0,CatalogExe);
  CatalogExe:=             CatalogExe+'\';
  CatalogShablon:=         CatalogExe+'\Шаблоны\';
  CatalogResultOMSPart:=   CatalogExe+'\Результат\01_ОМС Партия\';
  CatalogResultOMSAkt:=    CatalogExe+'\Результат\02_ОМС Акт ПРМ-ПРД\';
  CatalogResultTXOAkt:=    CatalogExe+'\Результат\03_ХТО Акт ПРМ-ПРД\';
  CatalogResultSMS_XLS:=   CatalogExe+'\Результат\04_Для SMS\XLS\';
  CatalogResultSMS_CSV:=   CatalogExe+'\Результат\04_Для SMS\CSV\';
  CatalogResultSMS_XLS_C:= 'C:\';
  CatalogResultDelo_XTO:=  CatalogExe+'\Результат\06_Как дела ХТО\';
  CatalogResultDelo_OMS:=  CatalogExe+'\Результат\07_Как дела ОМС\';
  CatalogResultDolg:=      CatalogExe+'\Результат\08_Долг ИТ\';
  Catalog_SMS:=            CatalogExe+'\Результат\09_SMS отчет\';
  Catalog_PoiskTel:=       CatalogExe+'\Результат\10_Поиск по телефону\';
  Catalog_Statistic_SMS:=  CatalogExe+'\Результат\11_Статистика SMS\';
  //предварительная проверка подключения к БД
  //-----------------------------------------

 try
    pr_conn:= ADOConnect.Connected;
    if(pr_conn=false) then
      begin
        ADOConnect.Connected := false;
        ADOConnect.Connected := true;
        pr_connect_label.Caption:='Подключено к БД';
      end;
  except
    pr_connect_label.Caption:='Ошибка подключения к БД:';
    exit;
  end;
 kolvo:=0;
 ClearInfoCaption();
 enp_stop.Caption:='';
 end;



procedure TScanPolis.btnOpenPartClick(Sender: TObject);

begin
 StatusBar.Panels[3].Text:='ЗАГРУЗИТЕ ФАЙЛ ИЗ ТФОМС С НОМЕРАМИ БЛАНКОВ';
 loadBlank.Visible:=true;
 btnClosePart.Visible:=true;
end;


procedure TScanPolis.part_numberButtonClick(Sender: TObject);
begin
 part_number.Text:='637_29_11_2016';
end;




procedure TScanPolis.statENPClick(Sender: TObject);
var
  str: string;
begin
    //str:='select e.part, count(*) as kolvo ';       Q7.SQL.Add(str);
    //str:='from enp e group by 1 order by 1 desc;';  Q7.SQL.Add(str);
   StatusBar.Panels[1].Text:='выполняется SQL-запрос ... ждите ...';
   StatusBar.Panels[2].Text:='_';
  //lb_info.Caption:='_';
  //lb_count.Caption:='_';
  Application.ProcessMessages;
    Q7.Close;
    Q7.SQL.Clear;
      str:='select top 5 NAME_FILE, COUNT(*) as kolvo ';                                          Q7.SQL.Add(str);
      str:='from BLANK where NAME_FILE LIKE ''%_2016%''  group by NAME_FILE order by 1 desc';          Q7.SQL.Add(str);
    Q7.Open;
  StatusBar.Panels[1].Text:= '_';
  StatusBar.Panels[2].Text:='_';
  Application.ProcessMessages;
end; //статистика обработанных ЕНП


procedure TScanPolis.loadBlankClick(Sender: TObject);
var
str: string;
 cnt: integer;
 cnt_all: integer;
 WordFileName: string;
 imf, name_file:    string;
 w,row,table:  Variant;  //Word
 number_table: integer;
 number_line:  integer;

 kolvo_columns: integer; //кол-во столбцов в таблице
 kolvo_rows: integer;    //кол-во строк в таблице

 i,j,k: integer;

 enp:          string;
 number_blank: string;

 cnt_error_dubl: integer;

 catalog_doc: string;     //каталог с файлами *.doc
 mask: string;            //маска для поиска *.doc
 SR : TSearchRec;         //поисковая переменная
 FileList :  TStrings;    //список файлов *.doc
 begin
  StatusBar.Panels[1].Text:= '_';
  StatusBar.Panels[2].Text:='_';
  Application.ProcessMessages;

  try
    number_table:=StrToInt(e_number_table.Text); //номер таблицы
  except
    number_table:=2; //номер таблицы по умолчанию при возникновении ошибки
  end;
  number_line:=1;  //номер стартовой строки в таблице

  if OD1.Execute then
    begin
      catalog_doc:= ExtractFilePath(OD1.FileName); 
      mask:= catalog_doc + '*.doc';
      FindFirst(mask,faAnyFile,SR);

      FileList := TStringList.Create();
      repeat
        str:= catalog_doc+ ExtractFileName(SR.Name);
        FileList.Add(str);  //массив стрингов (взяли все файлы по маске *.doc в выбранном каталоге)
      until FindNext(SR) <> 0;
      FindClose(SR);

      for k:=0 to FileList.Count-1 do    //пробегаемся по всем файлам в выбранном каталоге
        begin
          WordFileName:=FileList.Strings[k];
          imf:= ExtractFileName(WordFileName);
          name_file:=StringReplace(imf,ExtractFileExt(imf),'',[]);

          StatusBar.Panels[1].Text:= Format('%d: %s',[k+1,name_file]);
          Application.ProcessMessages;

          try
            w:=CreateOleObject('Word.Application');     //word-шаблон акта экспертного контроля
            w.Documents.Add(WordFileName);
            w.Visible:=false;

            //определим таблицу для работы в Word-файле
            //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            row:=w.ActiveDocument.Tables.item(number_table).Rows.Item(number_line);  //начинаем с 1-й строки
            table:=w.ActiveDocument.Tables.item(number_table);

            //определим кол-во столбцов и строк в таблице
            kolvo_columns:=(w.ActiveDocument.Tables.Item(number_table).Columns.Count)/3;
            kolvo_rows:=w.ActiveDocument.Tables.Item(number_table).Rows.Count;

            str:=Format('%d %d',[kolvo_columns,kolvo_rows]);

            cnt_all:=0;
            cnt:=0;
            cnt_error_dubl:=0; //кол-во дублирующихся записей
            for i := 0 to (kolvo_columns-1) do
              begin //пробежим по столбцам
                for j := 1 to kolvo_rows do
                  begin //пробежим по строкам
                    inc(cnt_all);
                    StatusBar.Panels[2].Text:= Format('%d',[cnt_all]);
                    Application.ProcessMessages;

                    enp:= Trim(table.Cell(j,(i*3+1)).range.Text);
                    number_blank:= Trim(table.Cell(j,(i*3+2)).range.Text);

                    if Length(enp)=16 then
                      begin //запишем ЕНП и номер бланка
                        try
                          Q3.Close;
                          Q3.SQL.Clear;
                            str:= 'insert into blank (number_blank,enp,name_file) ';                      Q3.SQL.Add(str);
                            str:= Format('values(''%s'',''%s'',''%s''); ',[number_blank,enp,name_file]);  Q3.SQL.Add(str);
                          Q3.ExecSQL;


                          inc(cnt); //счетчик записанных бланков
                        except
                          inc(cnt_error_dubl); //имеем дублирующиеся записи
                        end;
                      end; //if Length(enp)=16

                  end; //for i := 1 to kolvo_rows
              end; //for i := 1 to kolvo_columns

            w.Quit;        //очистим память, если ошибка при подключении
            w:=Unassigned;
          except
            str:='Ошибка при работе с Word-файлом: '+WordFileName;
            StatusBar.Panels[1].Text:=str;
            Application.ProcessMessages;
            Application.MessageBox(PChar(str),'Ошибка',MB_OK+MB_ICONERROR);

            StatusBar.Panels[1].Text:='Закрываем OLE-Word. Ждите...';
            Application.ProcessMessages;
            w.Quit;        //очистим память, если ошибка при подключении
            w:=Unassigned;

            str:='Ошибка при работе с Word-файлом: '+WordFileName;
            StatusBar.Panels[1].Text:=str;
            Application.ProcessMessages;
          end;

          if cnt_error_dubl>0 then
            begin
              str:=Format(' %s  Выполнено! Кол-во записанных бланков: %d,    Кол-во дублирующих записей: %d',[name_file,cnt,cnt_error_dubl]);
              StatusBar.Panels[0].Text:=str;
            end;

        end; //for k:=0 to k.Count-1

      FileList.Destroy;

      if cnt_error_dubl=0 then
        begin
          str:=Format('%s ЗАГРУЖЕНО! Кол-во записанных бланков: %d',[name_file,cnt]);
          Application.MessageBox(PChar(str),'Информация',MB_OK+MB_ICONINFORMATION);

        end;

    end; //if OD1.Execute
    lb_count.Caption:=IntToStr(cnt);
    lb_info.Caption:=Format('Партия - %s : ',[name_file]);
    loadBlank.Visible:=false;
    StatusBar.Panels[3].Text:='';
    part_number.Text:=name_file;
//----------------------------------------------------------------------------


end; //Закачать номера бланков в т.BLANK
//------------------------------------------------------------------------------

procedure TScanPolis.btnClosePartClick(Sender: TObject);
var
  str: string;
  Excel,WB : Variant;
  WorkBook:  IXLSWorkBook;
  WorkSheet: IXLSWorksheet;

  date_now: string;      //дата построения отчета

  sm: integer;   //смещение
  cnt: integer;  //счетчик

  kolvo_all: integer;    //общее кол-во полисов в партии

  s1,s2: string;
  ShExcel: string;
  FileExcel: string;

  pr_connect: boolean;
begin
   //предварительная проверка подключения к БД
  //-----------------------------------------
  try
    pr_connect:= ADOConnect.Connected;
    if(pr_connect=false) then
      begin
        ADOConnect.Connected := false;
        ADOConnect.Connected := true;
        pr_connect_label.Caption:='Подключено к БД';
      end;
  except
     pr_connect_label.Caption:='Ошибка подключения к БД: ';
    exit;
  end;
 if length(part_number.Text)<>0 then
    begin
      lb_info.Caption:= 'выполняется SQL-запрос ...';
      lb_count.Caption:='_';
      Application.ProcessMessages;

      date_now:= DateTimeToStr(Now); //время, когда формировался отчет

      Q11.Close;
      Q11.SQL.Clear;
        str:=Format('select * from blank b where b.part=''%s'' and b.oms_report_date is null;',[part_number.Text]); Q11.SQL.Add(str);
      Q11.Open;
      if Q11.RecordCount <> 0 then
        begin  //отчет ОМС строит 1-й раз, необходимо проставить дату отчета в поле OMS_REPORT_DATE
          str := Format('Вы уверены, что хотите "ЗАКРЫТЬ" партию %s  ???',[part_number.Text]);
          if  Application.MessageBox(PChar(str),'Проверка действий',MB_OKCANCEL+MB_ICONINFORMATION) = id_OK then
            begin
              lb_info.Caption:='Проставление даты отчета отдела ОМС...';
              lb_count.Caption:='_';     //исходное состояние
              Application.ProcessMessages;
                Q11.Close;
                Q11.SQL.Clear;
                  str:=Format('update blank set oms_report_date=''%s'' where part=''%s'';',[date_now,part_number.Text]); Q11.SQL.Add(str);
                Q11.ExecSQL;
              lb_info.Caption:='Проставление даты отчета отдела ОМС - OK!';
              lb_count.Caption:='_';     //исходное состояние
              Application.ProcessMessages;

              lb_info.Caption:='Подготовка данных для отдела ХТО...';
              lb_count.Caption:='_';     //исходное состояние
              Application.ProcessMessages;
               // Q11.Close;
               // Q11.SQL.Clear;
                 // str:=Format('select * from it_p_update_history_part(''%s'',''%s'');',[part_number.Text,date_now]); Q11.SQL.Add(str);
               // Q11.Open;
              lb_info.Caption:='Подготовка данных для отдела ХТО - ОК!';
              lb_count.Caption:='_';     //исходное состояние
              Application.ProcessMessages;
            end  //Проверка действий
          else
            begin
              exit;  //не стали закрывать партию полисов
            end;
        end; //if Q11.RecordCount <> 0 then

      lb_info.Caption:= 'выполняется SQL-запрос ...';
      lb_count.Caption:='_';
      Application.ProcessMessages;
       // Q11.Close;
       // Q11.SQL.Clear;
         // str:=Format('select * from it_p_akt_prm_prd(''%s'');',[part_number.Text]); Q11.SQL.Add(str);
       // Q11.Open;
      lb_info.Caption:='_';
      lb_count.Caption:='_';     //исходное состояние
      Application.ProcessMessages;
      if Q11.RecordCount <> 0 then
        begin //строим Акт приема-передачи
          ShExcel:=  CatalogShablon+'Акт приема-передачи_отд_ОМС.xlt';
          FileExcel:=CatalogResultOMSAkt+Format('Акт приема-передачи_%s.xls',[part_number.Text]);

          WorkBook:=TXLSWorkBook.Create;
          try
            if WorkBook.Open(ShExcel)=1 then
              begin //шаблон-Excel - открыт
                WorkSheet:=WorkBook.Sheets[1];
                WorkSheet.Name:=Format('%s',[part_number.Text]);

                WorkSheet.Cells[1,2].Value:=  'сформирован:  '+date_now;  //дата формирования Акта приема-передачи
                WorkSheet.Cells[6,3].Value:=  part_number.Text;           //номер партии
                WorkSheet.Cells[2,4].Value:=  copy(part_number.Text,1,3); //номер партии сокращенный по просьбе Васяниной Т.А.

                sm:=11;       //смещение в Excel-файле
                cnt:=0;       //счетчик
                kolvo_all:=0; //общее кол-во полисов в партии
                while not Q11.Eof do
                  begin
                    inc(cnt);
                    lb_count.Caption:=Format('%d',[cnt]);
                    Application.ProcessMessages;

                    WorkSheet.Cells[(sm+cnt),1].Value:= Q11.FieldByName('filial_id').AsInteger;
                    WorkSheet.Cells[(sm+cnt),2].Value:= Q11.FieldByName('filial_name').AsString;
                    WorkSheet.Cells[(sm+cnt),3].Value:= Q11.FieldByName('kolvo').AsInteger;

                    kolvo_all:=kolvo_all+ Q11.FieldByName('kolvo').AsInteger;

                    s1:=Format('A%d',[sm+cnt+1]); s2:=Format('C%d',[sm+cnt+1]); //вставим строку
                    Workbook.Sheets[1].Range[s1, s2].EntireRow.Insert(xlShiftDown);

                    Q11.Next;
                  end;

                WorkSheet.Cells[9,2].Value:=  Format('акцепт отд.ОМС %s',[Q11.FieldByName('oms_report_date').AsString]); //дата закрытия партии в отделе ОМС
                WorkSheet.Cells[7,3].Value:=  kolvo_all;           //общее кол-во полисов в партии

               // str:= szMoneyInWords_2(kolvo_all);
               // WorkSheet.Cells[8,1].Value:= str;

                //удалим 2-ве вспомогательные последние строки
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                s1:=Format('A%d',[sm+cnt+1]);
                Workbook.Sheets[1].Range[s1, s1].EntireRow.Delete(xlShiftUp);
                s1:=Format('A%d',[sm+cnt+1]);
                Workbook.Sheets[1].Range[s1, s1].EntireRow.Delete(xlShiftUp);

                WorkSheet.Cells[(sm+cnt+1),3].Value:=  kolvo_all;   //общее кол-во полисов в партии

                //сохранить
                str:= ExtractFilePath(FileExcel);
                if DirectoryExists(ExtractFilePath(str))=false then //если папки нет, тогда ее создадим
                  MakeDir(str);    //создаем папку для результата

                WorkBook.SaveAs(FileExcel); //сохраняем результат

                //откроем файл
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                try
                  Excel:=CreateOleObject('Excel.Application');
                  WB:=Excel.WorkBooks.Add(FileExcel);
                  //Excel.WindowState := -4140; //Excel свернуть на панель задач
                  Excel.Visible:=true;
                finally
                  Excel:=null;
                  WB:=null;
                end;
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

              end; //if WorkBook.Open(ShExcel)=1
          except
            str:='Ошибка при формировании файла '+FileExcel;
            Application.MessageBox(PChar(str),'Ошибка',MB_OK+MB_ICONSTOP);
          end;
          WorkBook.Close;
        end  //if Q11.RecordCount <> 0 then - есть ошибки при Проверке 1
      else
        begin
          str:='Нет данных в указанной партии: '+part_number.Text;
          Application.MessageBox(PChar(str),'Информация',MB_OK+MB_ICONINFORMATION);
        end;

    end  //if length(part_number.Text)<>0
  else
    begin
      str:='Номер партии не указан!';
      Application.MessageBox(PChar(str),'Информация',MB_OK+MB_ICONINFORMATION);
    end;

end; //"Закрыть"  партию полисов в отделе ОМС  }
 //begin
// btnOpenPart.Visible:=false;
 // end;



procedure TScanPolis.btnScanKolvo1Click(Sender: TObject);
var
  str: string;
begin
   QQ19.Close;
    QQ19.SQL.Clear;
      str:='select top 1 NAME_FILE, COUNT(*) as kolvo ';                                          QQ19.SQL.Add(str);
      str:='from BLANK where PART LIKE ''%_2016%''  group by NAME_FILE order by 1 desc';          QQ19.SQL.Add(str);
    QQ19.Open;
end;


end.
