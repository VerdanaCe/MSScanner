unit commonClasses;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ToolWin, ComCtrls, Grids, DBGrids, DB, IBCustomDataSet, IBQuery,
  IBDatabase, RXCtrls, StdCtrls, ExtCtrls,IniFiles, ztvZipTV, ztvUnARJ, ztvUnRar,
  ztvRegister, ztvbase, ztvUnZip,IBSQL,ComObj, FileCtrl, ShlObj, DateUtils;

type

  {------- информация о выполнении функции -------}
  TFuncInfo = class
  public  //информация о выполнении функции
    state:            boolean;
    info:             string;
  end;

  {коннектор}
  TConnector = class(TObject)
  private
    FDatabase: TIBDatabase;
    FTransaction: TIBTransaction;
    function GetDatabase(): TIBDatabase;
    function GetTransaction(): TIBTransaction;
    procedure SetDatabase(aDatabase: TIBDatabase);
    procedure SetTransaction(aTransaction: TIBTransaction);
  public
    constructor Create(aDatabase: TIBDataBase; aTransaction: TIBTransaction);
  published
    property Database: TIBDatabase read GetDatabase write SetDatabase;
    property Transaction: TIBTransaction read GetTransaction write SetTransaction;
  end;

  {информационный элемент}
  TPacMedInfo = class
  public  //информация - одна строчка-результат объединненного запроса Q11
    idlpu:  string;
    nlpu:   string;
    fio_vr: string;
    rstamp: string;
    idpom:  char;
    npr:    integer;
    npr_s:  integer;
    pol:    char;
    fam:    string;
    im:     string;
    otch:   string;
    dr:     string;
    d_beg:  string;
    d_end:  string;
    ncard:  string;
    mkb:    string;
    ishod:  integer;
    celpos: integer;
    tarif:  integer;  //необходимо поделить / 100
    summa:  integer;  //необходимо поделить / 100
    spe_vr: integer;
    idprof: integer;
    raion:  string;
    ulica:  string;
    dom:    string;
    korp:   string;
    kvart:  string;
    pr:     char;
    //constructor Create();
    //destructor Destroy();
  end;

  TQ_otv = class
  public
    List: TStringList;
    lines: string;
  end;

  TextEdit = class
  public
    state: boolean;
    name: string;
  end;

  TVolumeInfo = class
  public  //информация о периоде оплаченных объемов, загруженных в БД
    state:                boolean;
    name:                 string;
    min_volume_god:       integer;
    max_volume_god:       integer;
    min_volume_mes:       integer;
    max_volume_mes:       integer;
    str_min_volume_mes:   string;
    str_max_volume_mes:   string;
    min_volume_bd_godmes: integer;
    max_volume_bd_godmes: integer;
  end;

  TChInfo = class
  public  //информация о cчет-фактуре
    nom_ch:           string;
    date_ch:          string;
    kolvo:            currency;
    summ_ch:          currency;
  end;



  TRstampInfo = class
  public  //информация о реестре
    rstamp:           string;
    mes:              integer;
    god:              integer;
    dat_zagr:         string;
    nom_ch:           string;
    date_ch:          string;
    kolvo:            currency;
    summ_ch:          currency;
    note1_date:       string;
    info:             string;
  end;


  procedure MakeDir(value:string);                 //создание каталога
  function  Coalesce(aValue: variant): Double;     //функция Coalesce возвращает 0, если параметр null, иначе - сам параметр
  function SelectDir(Caption: string; HandleForm: HWND): string;      //выбор каталога
  //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

implementation

{-------- коннектор --------}
constructor TConnector.Create(aDatabase: TIBDataBase; aTransaction: TIBTransaction);
begin
  FDatabase := aDatabase;
  FTransaction := aTransaction;
end;

function TConnector.GetDatabase(): TIBDataBase;
begin
  result := FDatabase;
end;

function TConnector.GetTransaction(): TIBTransaction;
begin
  result := FTransaction;
end;

procedure TConnector.SetDatabase(aDatabase: TIBDataBase);
begin
  FDatabase := aDatabase;
end;

procedure TConnector.SetTransaction(aTransaction: TIBTransaction);
begin
  FTransaction := aTransaction;
end; {-------- коннектор --------}


{---------создать каталог-----------}
procedure MakeDir(value:string);
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
end;
{---------создать каталог-----------}


{функция Coalesce возвращает 0, если параметр null, иначе - сам параметр}
function Coalesce(aValue: variant): Double;
var
  d: double;
begin
  if aValue = null then
    d := 0
  else
  begin
    try
      d := aValue
    except
      d := 0;
    end;
  end;

  d:=StrToCurr(FloatToStrF(d,ffFixed,10,2));
  result := d;
end;
{функция Coalesce возвращает 0, если параметр null, иначе - сам параметр}

{----Удалить каталог со всем содержимым----}
//------------------------------------------------------------------------------
function DeleteDir(Dir  : string)  : boolean;
Var
 Found  : integer;
 SearchRec : TSearchRec;
begin
  result:=false;
  if IOResult<>0 then ;
  ChDir(Dir);
  if IOResult<>0 then begin
   ShowMessage('Не могу войти в каталог: '+Dir); exit;
  end;
  Found := FindFirst('*.*', faAnyFile, SearchRec);
  while Found = 0 do
  begin
   if (SearchRec.Name<>'.')and(SearchRec.Name<>'..') then
    if (SearchRec.Attr and faDirectory)<>0 then begin
     if not DeleteDir(SearchRec.Name) then exit;
    end else
     if not DeleteFile(SearchRec.Name) then begin
      ShowMessage('Не могу удалить файл: '+SearchRec.Name); exit;
     end;
    Found := FindNext(SearchRec);
  end;
  FindClose(SearchRec);
  ChDir('..'); RmDir(Dir);
  result:=IOResult=0;
end; {----Удалить каталог со всем содержимым----}
//------------------------------------------------------------------------------

//выбор каталога
//------------------------------------------------------------------------------
function SelectDir(Caption: string; HandleForm: HWND): string;  //передаем handle Формы
var
  options : TSelectDirOpts;   //выбор каталога
  chosenDirectory : string;
  TitleName : string;
  lpItemID : PItemIDList;
  BrowseInfo : TBrowseInfo;
  DisplayName : array[0..MAX_PATH] of char;
  TempPath : array[0..MAX_PATH] of char;
  catalog: string;
begin
  TempPath:='';
  catalog:=TempPath;

  try
    FillChar(BrowseInfo, sizeof(TBrowseInfo), #0);
    BrowseInfo.hwndOwner := HandleForm;
    BrowseInfo.pszDisplayName := @DisplayName;
    TitleName := Caption;
    BrowseInfo.lpszTitle := PChar(TitleName);
    BrowseInfo.ulFlags := BIF_RETURNONLYFSDIRS;
    lpItemID := SHBrowseForFolder(BrowseInfo);

    if lpItemId <> nil then
      begin
        SHGetPathFromIDList(lpItemID, TempPath);
        GlobalFreePtr(lpItemID);
        catalog:=TempPath;
        //TempPath;  -выбранный каталог
      end;
  except
    TempPath:='';
    catalog:=TempPath;
    
    try
    //при ошибке попробуем выбрать другой способ выбора каталога
    chosenDirectory := 'C:\';  // Установка начального каталога
    if SelectDirectory(chosenDirectory, options, 0) then
      catalog:= chosenDirectory;
    except
      ;
    end;
  end;

  if Length(catalog)<>0 then
    if catalog[Length(catalog)] <> '\' then
      catalog := catalog + '\';

  result:= catalog; //выбранный каталог
end; //выбор каталога
//------------------------------------------------------------------------------

end.
