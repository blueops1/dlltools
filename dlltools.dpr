library dlltools;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  SysUtils,
  Classes,
  Variants,
  Grids,
  ComObj,
  Printers,
  Windows,
  Messages,
  Graphics,
  Dialogs,
  Controls,
  forms;

Function AboutMe(Yhw:string) :string;stdcall;
  begin
    if Yhw='Yyh' then
      result:='Yhw'
    else
      result:='Hyf'
  end;

Function ExportStrGridToExcel(Args: array of const): Boolean;stdcall;
var iCount, jCount: Integer;
	XLApp: Variant;
	Sheet: Variant;
	I: Integer;
 begin
	Result := False;
	if not VarIsEmpty(XLApp) then
	begin
		XLApp.DisplayAlerts := False;
		XLApp.Quit;
		VarClear(XLApp);
	end;
	try
		XLApp := CreateOleObject('Excel.Application');
	except
		Exit;
	end;
	XLApp.WorkBooks.Add;
	XLApp.SheetsInNewWorkbook := High(Args) + 1;
	for I := Low(Args) to High(Args) do
	begin
		with TStringGrid(Args[I].VObject) do
		begin
			XLApp.WorkBooks[1].WorkSheets[I+1].Name := Name;
			Sheet := XLApp.Workbooks[1].WorkSheets[Name];
			for jCount := 0 to RowCount - 1 do
			begin
				for iCount := 0 to ColCount - 1 do
				begin
					Sheet.Cells[jCount + 1, iCount + 1] := Cells[iCount, jCount];
				end;
			end;
		end;
	end;
	XlApp.Visible := True;
 end;

 Function SplitString(Source, Deli: string ): String;stdcall;
 var
    EndOfCurrentString: byte;
    StringList:TStringList;
 begin
    StringList:=TStringList.Create;
    while Pos(Deli, Source)>0 do
    begin
      EndOfCurrentString := Pos(Deli, Source);
      StringList.add(Copy(Source, 1, EndOfCurrentString - 1));
      Source := Copy(Source, EndOfCurrentString + length(Deli), length(Source) - EndOfCurrentString);
    end;
    Result := StringList[0];
    StringList.Add(source);
end;

function CharHeight: Word;stdcall;
var
    Metrics: TTextMetric;
begin
    GetTextMetrics(Printer.Canvas.Handle, Metrics);
    Result := Metrics.tmHeight;
end;

//取得字符的平均宽度
function AvgCharWidth: Word;stdcall;
var
    Metrics: TTextMetric;
begin
    GetTextMetrics(Printer.Canvas.Handle, Metrics);
    Result := Metrics.tmAveCharWidth;
end;

//取得纸张的物理尺寸---单位：点
function GetPhicalPaper: TPoint;stdcall;
var
    PageSize : TPoint;
begin
    //PageSize.X/; 纸张物理宽度-单位:点
    //PageSize.Y/; 纸张物理高度-单位:点
    Escape(Printer.Handle, GETPHYSPAGESIZE, 0,nil,@PageSize);
    Result := PageSize;
end;

//取得纸张的逻辑宽度--可打印区域,取得纸张的逻辑尺寸
function PaperLogicSize: TPoint;stdcall;
var
    APoint: TPoint;
begin
    APoint.X := Printer.PageWidth;
    APoint.Y := Printer.PageHeight;
    Result := APoint;
end;

//纸张水平对垂直方向的纵横比例
function HVLogincRatio: Extended;stdcall;
var
    AP: TPoint;
begin
    Ap := PaperLogicSize;
    Result := Ap.y/Ap.X;
end;

//取得纸张的横向偏移量-单位：点
function GetOffSetX: Integer;stdcall;
begin
    Result := GetDeviceCaps(Printer.Handle, PhysicalOffSetX);
end;

//取得纸张的纵向偏移量-单位：点
function GetOffSetY: Integer;stdcall;
begin
    Result := GetDeviceCaps(Printer.Handle, PhysicalOffSetY);
end;

//毫米单位转换为英寸单位
function MmToInch(Length: Extended): Extended;stdcall;
begin
    Result := Length/25.4;
end;

//英寸单位转换为毫米单位
function InchToMm(Length: Extended): Extended; stdcall;
begin
    Result := Length*25.4;
end;

//取得水平方向每英寸打印机的点数
function HPointsPerInch: Integer;stdcall;
begin
    Result := GetDeviceCaps(Printer.Handle, LOGPIXELSX);
end;

//取得纵向方向每英寸打印机的光栅数
function VPointsPerInch: Integer;stdcall;
begin
    Result := GetDeviceCaps(Printer.Handle, LOGPIXELSY)
end;

//横向点单位转换为毫米单位
function XPointToMm(Pos: Integer): Extended;stdcall;
begin
    Result := Pos*25.4/HPointsPerInch;
end;

//纵向点单位转换为毫米单位
function YPointToMm(Pos: Integer): Extended;stdcall;
begin
    Result := Pos*25.4/VPointsPerInch;
end;

//设置纸张高度-单位：mm
procedure SetPaperHeight(Value:integer);stdcall;
var
    Device : array[0..255] of char;
    Driver : array[0..255] of char;
    Port : array[0..255] of char;
    hDMode : THandle;
    PDMode : PDEVMODE;
begin
    //自定义纸张最小高度127mm
    if Value < 127 then Value := 127;
    //自定义纸张最大高度432mm
    if Value > 432 then Value := 432;
    Printer.PrinterIndex := Printer.PrinterIndex;
    Printer.GetPrinter(Device, Driver, Port, hDMode);
    if hDMode <> 0 then
    begin
        pDMode := GlobalLock(hDMode);
        if pDMode <> nil then
        begin
            pDMode^.dmFields := pDMode^.dmFields or DM_PAPERSIZE or DM_PAPERLENGTH;
                        pDMode^.dmPaperSize := DMPAPER_USER;
            pDMode^.dmPaperLength := Value * 10;
            pDMode^.dmFields := pDMode^.dmFields or DMBIN_MANUAL;
            pDMode^.dmDefaultSource := DMBIN_MANUAL;
            GlobalUnlock(hDMode);
        end;
    end;
    Printer.PrinterIndex := Printer.PrinterIndex;
end;

//设置纸张宽度：单位--mm
Procedure SetPaperWidth(Value:integer); stdcall;
var
    Device : array[0..255] of char;
    Driver : array[0..255] of char;
    Port : array[0..255] of char;
    hDMode : THandle;
    PDMode : PDEVMODE;
begin
    //自定义纸张最小宽度76mm
    if Value < 76 then Value := 76;
    //自定义纸张最大宽度216mm
    if Value > 216 then Value := 216;
    Printer.PrinterIndex := Printer.PrinterIndex;
    Printer.GetPrinter(Device, Driver, Port, hDMode);
    if hDMode <> 0 then
    begin
        pDMode := GlobalLock(hDMode);
        if pDMode <> nil then
        begin
            pDMode^.dmFields := pDMode^.dmFields or DM_PAPERSIZE or
            DM_PAPERWIDTH;
            pDMode^.dmPaperSize := DMPAPER_USER;
            pDMode^.dmPaperWidth := Value * 10;  //将毫米单位转换为0.1mm单位
            pDMode^.dmFields := pDMode^.dmFields or DMBIN_MANUAL;
            pDMode^.dmDefaultSource := DMBIN_MANUAL;
            GlobalUnlock(hDMode);
        end;
    end;
    Printer.PrinterIndex := Printer.PrinterIndex;
end;

//在 (Xmm, Ymm)处按指定配置文件信息和字体输出字符串
procedure PrintText(Txt: string;item:string;configfilename:string);stdcall;
var
//    OrX, OrY: Extended;
    Px, Py,x,y: Integer;
//    AP: TPoint;
    Fn: TStrings;
    FileName: string;
    OffSetX, OffSetY: Integer;
begin
    //打开配置文件，读出横向和纵向偏移量
    try
        Fn := TStringList.Create;
        FileName := ExtractFilePath(Application.ExeName) + ConfigFileName;
        if FileExists(FileName) then
        begin
            Fn.LoadFromFile(FileName);
            OffSetX := StrToInt(Fn.Values['右移']); //横向偏移量
            OffSetY := StrToInt(Fn.Values['下移']); //纵向偏移量
                X := strtoint(Fn.Values[item+'X']) + OffSetX;
    Y := strtoint(Fn.Values[item+'y']) + OffSetY;
    Px := Round(Round(X * HPointsPerInch * 10000/25.4) / 10000);
    Py := Round(Round(Y * VPointsPerInch * 10000/25.4) / 10000);
    Py := Py - GetOffSetY; //因为是绝对坐标, 因此, 不用换算成相对于Y轴坐标
    Px := Px + 2 * AvgCharWidth;
    Printer.Canvas.Font.Name := Fn.Values[item+'font'];
    Printer.Canvas.Font.Size := strtoint(Fn.Values[item+'FontSize']);
        end
        else
        begin
        //如果没有配置文件，则生成
        Fn.Values['右移'] := '0';
        Fn.Values['下移'] := '0';
        Fn.SaveToFile(FileName);
        end;
    finally
        Fn.Free;
    end;

    //Printer.Canvas.Font.Color := clGreen;
    Printer.Canvas.TextOut(Px, Py, Txt);
end;

{
Function DXZH(f : String) : String;
var dx,d2,zs,xs,s1,s2,h,jg:string;
   i,ws,l,w,j,lx:integer;
begin
  f := Trim(f);
  if copy(f,1,1)='0' then begin
    Delete(f,1,1);end
  else ;
  dx:='零壹贰叁肆伍陆柒捌玖';
  d2:='拾佰仟万亿';
  i := AnsiPos('.',f);   //小数点位置
  If i = 0 Then
     zs := f     //整数
  Else begin
     zs:=copy(f,1,i - 1);  //整数部分
     xs:=copy(f,i + 1,200);
  End;
  ws:= 0; l := 0;
  For i :=Length(zs) downTo 1 do begin
    ws := ws + 1; h := '';
    w:=strtoint(copy(zs, i, 1));
    if (w=0) and (i=1) then jg:='零';
    If w > 0 Then
       Case ws of
         2..5:h:=copy(d2,(ws-1)*2-1,2);
         6..8:begin
           h:=copy(d2,(ws-5)*2-1,2);
           If AnsiPos('万',jg)=0 Then h:=h+'万';
           end;
         10..13:h :=copy(d2,(ws-9)*2-1, 2);
       End;

    jg:=copy(dx,(w+1)*2-1,2) + h + jg;
    If ws=9 Then jg :=copy(jg,1,2)+'亿'+copy(jg,3,200);
  end;
  j:=AnsiPos('零零',jg);
  While j>0 do begin
    jg :=copy(jg, 1, j - 1)+copy(jg,j+2,200);
    j :=AnsiPos('零零',jg);
  end;
  If (Length(jg)>1)And(copy(jg,length(jg)-1,2)='零')Then jg :=copy(jg,1,Length(jg)-2);
  j := AnsiPos('零亿',jg);
  If j > 0 Then jg:=copy(jg,1, j - 1)+copy(jg, j + 2,200);
  //转换小数部分
  If (Length(jg)>1) then    //定义元
     jg :=jg+'元'
  else
     jg:=jg;
  lx := Length(xs);
  If lx=0Then begin          //如果小数为零
    jg :=jg + '整' ;
  End;
  If lx=1Then begin         //如果小数为一位
    s1:=copy(dx, strtoint(copy(xs,1,1))*2 + 1, 2);
    if s1<>'零' then
      jg := jg+s1+'角'+'整' ;
    if s1='零' then
      jg := jg+'整' ;
  End;
  If lx>=2Then begin        //小数为两位
    s1:=copy(dx, strtoint(copy(xs,1,1))*2 + 1, 2);
    s2:=copy(dx, strtoint(copy(xs,2,1))*2 + 1, 2) ;
    if (s1='零')and (s2='零') then
       jg := jg +'整' ;
    if (s1<>'零')and (s2<>'零') then
       jg := jg +s1+'角'+s2+'分' ;
    if (s1<>'零')and (s2='零') then
       jg := jg +s1+'角'+'整' ;
    if (s1='零')and (s2<>'零') then
       jg := jg +s1+s2+'分' ;
  End;
  DXZH:=jg;
End;
}
function SmallTOBig(small:real):string;stdcall;
var
    SmallMonth,BigMonth:string;
    wei1,qianwei1:string[2];
    qianwei,dianweizhi,qian:integer;
begin
    {------- 修改参数令值更精确 -------}
    qianwei:=-2;{小数点后的位置，需要的话也可以改动-2值}
    Smallmonth:=formatfloat('0.00',small);{转换成货币形式，需要的话小数点后加多几个零}
    {---------------------------------}
    dianweizhi :=pos('.',Smallmonth);{小数点的位置}
    for qian:=length(Smallmonth) downto 1 do{循环小写货币的每一位，从小写的右边位置到左边}
        begin
            if qian<>dianweizhi then{如果读到的不是小数点就继续}
                begin
                    case strtoint(copy(Smallmonth,qian,1)) of{位置上的数转换成大写}
                        1:wei1:='壹';
                        2:wei1:='贰';
                        3:wei1:='叁';
                        4:wei1:='肆';
                        5:wei1:='伍';
                        6:wei1:='陆';
                        7:wei1:='柒';
                        8:wei1:='捌';
                        9:wei1:='玖';
                        0:wei1:='零';
                    end;
                    case qianwei of{判断大写位置，可以继续增大到real类型的最大值}
                        -3:qianwei1:='厘';
                        -2:qianwei1:='分';
                        -1:qianwei1:='角';
                        0 :qianwei1:='元';
                        1 :qianwei1:='拾';
                        2 :qianwei1:='佰';
                        3 :qianwei1:='仟';
                        4 :qianwei1:='万';
                        5 :qianwei1:='拾';
                        6 :qianwei1:='佰';
                        7 :qianwei1:='仟';
                        8 :qianwei1:='亿';
                        9 :qianwei1:='拾';
                        10:qianwei1:='佰';
                        11:qianwei1:='仟';
                    end;
                    inc(qianwei);
                    BigMonth :=wei1+qianwei1+BigMonth;{组合成大写金额}
                end;
        end;
    SmallTOBig:=BigMonth;
end;


{$R *.res}

exports
AboutMe,
//取得分割字符的第一个字符
SplitString,
//stringgrid导出到execl;
ExportStrGridToExcel,
//取得字符的高度
CharHeight,
//取得字符的平均宽度
AvgCharWidth,
//取得纸张的物理尺寸---单位：点
GetPhicalPaper,
//取得纸张的逻辑宽度--可打印区域,取得纸张的逻辑尺寸
PaperLogicSize,
//纸张水平对垂直方向的纵横比例
HVLogincRatio,
//取得纸张的横向偏移量-单位：点
GetOffSetX,
//取得纸张的纵向偏移量-单位：点
GetOffSetY,
//毫米单位转换为英寸单位
MmToInch,
//英寸单位转换为毫米单位
InchToMm,
//取得水平方向每英寸打印机的点数
HPointsPerInch,
//取得纵向方向每英寸打印机的光栅数
VPointsPerInch,
//横向点单位转换为毫米单位
XPointToMm,
//纵向点单位转换为毫米单位
YPointToMm,
//设置纸张高度-单位：mm
SetPaperHeight,
//设置纸张宽度：单位--mm
SetPaperWidth,
//在 (Xmm, Ymm)处按指定配置文件信息和字体输出字符串
PrintText,
//小写金额转换为大写金额
SmallTOBig;

begin

end.
