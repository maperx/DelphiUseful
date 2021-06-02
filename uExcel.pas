// --- Copyright © 2020 Matias A. P.
// --- All rights reserved
// --- maperx@gmail.com

unit uExcel;

interface

uses Winapi.Windows, System.SysUtils, Variants, ComObj, System.Classes;

type
  TExcel = class(TObject)
  private
    FExcel, FActiveSheet: OleVariant;
    FFileName: string;
    FIsNew: Boolean;
  public
    constructor Create(const FileName: string = '');
    destructor Destroy; override;
    procedure SetActiveSheet(const Idx: Integer);
    procedure Save;
    procedure SaveAs(const FileName: string);
    function GetCellData(const Row, Col: Integer): OleVariant; overload;
    function GetCellData(const Cell: string): OleVariant; overload;
    function GetCellFloat(const Row, Col: Integer): Double; overload;
    function GetCellFloat(const Cell: string): Double; overload;
    function GetCellInteger(const Row, Col: Integer): Integer; overload;
    function GetCellInteger(const Cell: string): Integer; overload;
    function GetCellDate(const Row, Col: Integer): TDate; overload;
    function GetCellDate(const Cell: string): TDate; overload;

    procedure SetCellData(const Row, Col: Integer; Data: OleVariant); overload;
    procedure SetCellData(const Cell: string; Data: OleVariant); overload;
    procedure SetColsStrings(const Row, Col: Integer; var Strs: TStringList);
    procedure SetRowsStrings(const Row, Col: Integer; var Strs: TStringList);

    procedure SetRangeBold(const Cell1, Cell2: string; Bold: Boolean = True);

    procedure NextRow;
    function GetActiveCell: string;
  published
    property FullPath: string read FFileName;
  end;

implementation

const
  Worksheet = -4167;

constructor TExcel.Create(const FileName: string);
begin
  FExcel := CreateOleObject('Excel.Application');
  if VarIsNull(FExcel) then
    Exit;
  FExcel.Visible := False;
  FFileName := FileName;
  FIsNew := not FileExists(FFileName);
  if FIsNew then begin
    FExcel.Workbooks.Add(Worksheet);
  end
  else begin
    FExcel.Workbooks.Open(FFileName);
  end;
  SetActiveSheet(1);
end;

destructor TExcel.Destroy;
begin
  try
    if not VarIsEmpty(FExcel) then begin
      FExcel.DisplayAlerts := False;
      FExcel.Quit;
      FExcel := Unassigned;
      FActiveSheet := Unassigned;
    end;
  except

  end;
  inherited;
end;

function TExcel.GetActiveCell: string;
begin
  Result := FExcel.ActiveCell.Address;
end;

function TExcel.GetCellData(const Cell: string): OleVariant;
begin
  REsult := FActiveSheet.Range[Cell].Value;
end;

function TExcel.GetCellDate(const Row, Col: Integer): TDate;
begin
  REsult := StrToDateDef(FActiveSheet.Cells[Row, Col].Value, -1);
end;

function TExcel.GetCellDate(const Cell: string): TDate;
begin
  REsult := StrToDateDef(FActiveSheet.Range[Cell].Value, -1);
end;

function TExcel.GetCellFloat(const Row, Col: Integer): Double;
begin
  REsult := StrToFloatDef(FActiveSheet.Cells[Row, Col].Value, -1);
end;

function TExcel.GetCellFloat(const Cell: string): Double;
begin
  REsult := StrToFloatDef(FActiveSheet.Range[Cell].Value, -1);
end;

function TExcel.GetCellInteger(const Cell: string): Integer;
begin
  REsult := -1;
  try
    REsult := StrToIntDef(FActiveSheet.Range[Cell].Value, -1);
  except

  end;
end;

function TExcel.GetCellInteger(const Row, Col: Integer): Integer;
begin
  REsult := StrToIntDef(FActiveSheet.Cells[Row, Col].Value, -1);
end;

procedure TExcel.NextRow;
begin
  FExcel.ActiveCell.Offset(1, 0).Select;
end;

function TExcel.GetCellData(const Row, Col: Integer): OleVariant;
begin
  REsult := FActiveSheet.Cells[Row, Col].Value;
end;

procedure TExcel.Save;
begin
  if FFileName = EmptyStr then
    Exit;
  try
    if FIsNew then
      FExcel.Workbooks[1].SaveAs(FFileName)
    else
      FExcel.Workbooks[1].Save;
  except
    on E: Exception do
      raise Exception.Create(E.Message);
  end;
end;

procedure TExcel.SaveAs(const FileName: string);
begin
  if FileName = EmptyStr then
    Exit;
  try
    FExcel.Workbooks[1].SaveAs(FileName);
    FFileName := FileName;
  except
    on E: Exception do
      raise Exception.Create(E.Message);
  end;
end;

procedure TExcel.SetActiveSheet(const Idx: Integer);
begin
  FActiveSheet := FExcel.Workbooks[1].WorkSheets[Idx];
end;

procedure TExcel.SetCellData(const Cell: string; Data: OleVariant);
begin
  FActiveSheet.Range[Cell].Value := Data;
end;

procedure TExcel.SetColsStrings(const Row, Col: Integer; var Strs: TStringList);
var
  i: Integer;
begin
  for i := 0 to Strs.Count - 1 do
    FActiveSheet.Cells[Row, Col + i].Value := Strs[i];
end;

procedure TExcel.SetRowsStrings(const Row, Col: Integer; var Strs: TStringList);
var
  i: Integer;
begin
  for i := 0 to Strs.Count - 1 do
    FActiveSheet.Cells[Row + i, Col].Value := Strs[i];
end;

procedure TExcel.SetRangeBold(const Cell1, Cell2: string; Bold: Boolean);
begin
  FActiveSheet.Range[Cell1, Cell2].Font.Bold := Bold;
end;

procedure TExcel.SetCellData(const Row, Col: Integer; Data: OleVariant);
begin
  FActiveSheet.Cells[Row, Col].Value := Data;
end;

end.
