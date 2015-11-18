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
  ComObj;

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


{$R *.res}

exports
AboutMe,
SplitString,
ExportStrGridToExcel;

begin

end.
