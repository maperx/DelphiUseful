// --- Copyright © 2020 Matias A. P.
// --- All rights reserved
// --- maperx@gmail.com

unit uPDF;

interface

uses mORMotReport, SynGDIPlus, Vcl.Dialogs, Vcl.DBGrids, Vcl.Graphics, System.Classes, Data.DB, System.SysUtils;

type
  TTextAlign = (taLeft, taRight, taCenter, taJustified);

  TPDF = class(TObject)
  private
    FontCode128Path, FontEanPath: string;
    FontCode128Ok, FontEanOk: Boolean;
    function GetFont: Vcl.Graphics.TFont;
  public
    FGDI: TGDIPages;
    constructor Create();
    destructor Destroy; override;
    procedure Preview;
    procedure NewLines(const Count: Byte = 1);
    procedure NewPage;
    procedure Text(const Str: string; WithNewLine: Boolean = True; Align: TTextAlign = taLeft; Bold: Boolean = False; Underline: Boolean = False; Italic: Boolean = False);
    procedure TextAt(const Str: string; XPos: Integer; Align: TTextAlign = taLeft; Bold: Boolean = False; Underline: Boolean = False; Italic: Boolean = False);
    procedure AddHeaderText(const Str: string; Align: TTextAlign = taLeft);
    procedure AddFooterText(const Str: string; Align: TTextAlign = taLeft);
    procedure AddDataSet(var DataSet: TDataSet; BottomLine: Boolean = True; AddHeaders: Boolean = True; const RecIdx: Integer = -1; RecCount: Integer = 0; FirstField: Integer = 0;
      FieldsCount: Integer = 0);
    procedure AddDataSetGrp(const GrpKeyFd, GrpCaptionFd: string; var DataSet: TDataSet; AddHeaders: Boolean = True; FirstField: Integer = 0; FieldsCount: Integer = 0;
      NewPageGrp: Boolean = False);
    procedure AddGrid(var Grid: TDBGrid; WithBottomGrayLine: Boolean = True; AddHeaders: Boolean = True);
    procedure DrawBarCode(const BarCode: string; FontSize: Integer = 24; WithNewLine: Boolean = True; XPos: Integer = 0; Align: TTextAlign = taLeft);
    procedure DrawLine(const doble: Boolean = False);

    property Font: Vcl.Graphics.TFont read GetFont;
  end;

implementation

uses Winapi.Windows, Winapi.Messages;

{ TPDF }

constructor TPDF.Create;
begin
  FGDI := TGDIPages.Create(nil);
  FGDI.BeginDoc;
  FGDI.Orientation := TPrinterOrientation.poPortrait;
  FGDI.Font.Name := 'Consolas';
  FGDI.Font.Size := 10;
  FGDI.SaveLayout;
  FontCode128Ok := False;
  FontCode128Path := IncludeTrailingBackslash(GetCurrentDir) + 'code128.ttf';
  if FileExists(FontCode128Path) then begin
    FontCode128Ok := (AddFontResource(pchar(FontCode128Path)) > 0);
    SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0);
  end;

  FontEanOk := False;
  FontEanPath := IncludeTrailingBackslash(GetCurrentDir) + 'BarcodeFont.ttf';
  if FileExists(FontEanPath) then begin
    FontEanOk := (AddFontResource(pchar(FontEanPath)) > 0);
    SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0);
  end;
end;

destructor TPDF.Destroy;
begin
  FGDI.Free;
  if FontCode128Ok then begin
    RemoveFontResource(pchar(FontCode128Path));
    SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0);
  end;
  if FontEanOk then begin
    RemoveFontResource(pchar(FontEanPath));
    SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0);
  end;
  inherited;
end;

procedure TPDF.DrawBarCode(const BarCode: string; FontSize: Integer; WithNewLine: Boolean; XPos: Integer; Align: TTextAlign);
begin
  if (not FontCode128Ok) and (not FontEanOk) then
    Exit;
  FGDI.SaveLayout;
  FGDI.TextAlign := mORMotReport.TTextAlign(Align);
  if BarCode.Length > 13 then
    FGDI.Font.Name := 'Code 128'
  else
    FGDI.Font.Name := 'barcode font';
  FGDI.Font.Size := 24;
  if XPos > 0 then begin
    FGDI.DrawTextAt(BarCode, XPos);
    if WithNewLine then
      FGDI.NewLine;
  end
  else
    FGDI.DrawText(BarCode, WithNewLine);
  FGDI.RestoreSavedLayout;
end;

procedure TPDF.DrawLine(const doble: Boolean);
begin
  FGDI.DrawLine(doble);
end;

procedure TPDF.NewLines(const Count: Byte);
var
  i: Integer;
begin
  for i := 1 to Count do
    FGDI.NewLine;
end;

procedure TPDF.NewPage;
begin
  FGDI.NewPage();
end;

procedure TPDF.Preview;
begin
  FGDI.EndDoc;
  FGDI.ShowPreviewForm([rNextPage, rPreviousPage, rPrint, rExportPDF, rClose]);
  FGDI.BeginDoc;
end;

procedure TPDF.Text(const Str: string; WithNewLine: Boolean; Align: TTextAlign; Bold: Boolean; Underline: Boolean; Italic: Boolean);
// var   ta: TTextAlign;
begin
  // ta := TTextAlign(FGDI.TextAlign);
  FGDI.Font.Style := [];
  if Bold then
    FGDI.Font.Style := FGDI.Font.Style + [fsBold];
  if Underline then
    FGDI.Font.Style := FGDI.Font.Style + [fsUnderline];
  if Italic then
    FGDI.Font.Style := FGDI.Font.Style + [fsItalic];
  FGDI.TextAlign := mORMotReport.TTextAlign(Align);
  FGDI.DrawText(Str, WithNewLine);
  FGDI.Font.Style := [];
  FGDI.TextAlign := mORMotReport.TTextAlign(mORMotReport.TTextAlign.taLeft);
end;

procedure TPDF.TextAt(const Str: string; XPos: Integer; Align: TTextAlign; Bold, Underline, Italic: Boolean);
begin
  FGDI.Font.Style := [];
  if Bold then
    FGDI.Font.Style := FGDI.Font.Style + [fsBold];
  if Underline then
    FGDI.Font.Style := FGDI.Font.Style + [fsUnderline];
  if Italic then
    FGDI.Font.Style := FGDI.Font.Style + [fsItalic];
  FGDI.TextAlign := mORMotReport.TTextAlign(Align);
  FGDI.DrawTextAt(Str, Xpos);
  FGDI.Font.Style := [];
  FGDI.TextAlign := mORMotReport.TTextAlign(mORMotReport.TTextAlign.taLeft);
end;

procedure TPDF.AddHeaderText(const Str: string; Align: TTextAlign);
var
  ta: TTextAlign;
begin
  ta := TTextAlign(FGDI.TextAlign);
  FGDI.TextAlign := mORMotReport.TTextAlign(Align);
  FGDI.AddTextToHeader(Str);
  FGDI.TextAlign := mORMotReport.TTextAlign(ta);
end;

procedure TPDF.AddFooterText(const Str: string; Align: TTextAlign);
var
  ta: TTextAlign;
begin
  ta := TTextAlign(FGDI.TextAlign);
  FGDI.TextAlign := mORMotReport.TTextAlign(Align);
  FGDI.AddTextToFooter(Str);
  FGDI.TextAlign := mORMotReport.TTextAlign(ta);
end;

procedure TPDF.AddDataSet(var DataSet: TDataSet; BottomLine: Boolean = True; AddHeaders: Boolean = True; const RecIdx: Integer = -1; RecCount: Integer = 0; FirstField: Integer = 0;
  FieldsCount: Integer = 0);
var
  i, sumaAnchos, c: Integer;
  tits: array of string;
  porcAnchosCols: array of Integer;
  bm: TBookmark;
begin
  sumaAnchos := 0;
  if FieldsCount > 0 then
    c := FieldsCount
  else
    c := DataSet.Fields.Count - FirstField;
  SetLength(tits, c);
  SetLength(porcAnchosCols, c);
  if FieldsCount > 0 then
    c := FieldsCount - 1
  else
    c := DataSet.Fields.Count - 1;
  for i := FirstField to c do
    Inc(sumaAnchos, DataSet.Fields[i].DisplayWidth);
  for i := Low(tits) to High(tits) do begin
    tits[i] := DataSet.Fields[i + FirstField].DisplayLabel;
    porcAnchosCols[i] := DataSet.Fields[i + FirstField].DisplayWidth * 100 div sumaAnchos;
  end;
  FGDI.AddColumns(porcAnchosCols);
  if AddHeaders then
    FGDI.AddColumnHeaders(tits, BottomLine, True);
  for i := Low(tits) to High(tits) do
    case Ord(DataSet.Fields[i + FirstField].Alignment) of
      0: // taLeftJustify
        FGDI.SetColumnAlign(i, TColAlign.caLeft);
      1: // taRightJustify
        FGDI.SetColumnAlign(i, TColAlign.caRight);
      2: // taCenter
        FGDI.SetColumnAlign(i, TColAlign.caCenter);
    end;
  bm := DataSet.GetBookmark;
  if RecCount = 0 then begin
    DataSet.First;
    DataSet.DisableControls;
  end;
  try
    if (RecIdx >= 0) and (RecCount > 0) then begin
      DataSet.RecNo := RecIdx;
      for c := RecIdx to RecIdx + RecCount - 1 do begin
        for i := Low(tits) to High(tits) do
          tits[i] := DataSet.Fields[i + FirstField].Text;
        FGDI.DrawTextAcrossCols(tits);
        DataSet.Next;
        if DataSet.Eof then
          Break;
      end;
    end
    else begin
      while not DataSet.Eof do begin
        for i := Low(tits) to High(tits) do
          tits[i] := DataSet.Fields[i + FirstField].Text;
        FGDI.DrawTextAcrossCols(tits);
        DataSet.Next;
      end;
    end;
  finally
    DataSet.GotoBookmark(bm);
    if RecCount = 0 then
      DataSet.EnableControls;
  end;

end;

procedure TPDF.AddDataSetGrp(const GrpKeyFd, GrpCaptionFd: string; var DataSet: TDataSet; AddHeaders: Boolean; FirstField: Integer; FieldsCount: Integer; NewPageGrp: Boolean);
var
  i, c, r: Integer;
  id: Int64;

  procedure DrawGrp;
  begin
    FGDI.Font.Style := [fsBold];
    FGDI.DrawText(Format('%s (%d)', [DataSet.FieldByName(GrpCaptionFd).AsString, id]));
    // Text(DataSet.FieldByName(GrpCaptionFd).AsString, False);
    // Text(IntToStr(id), True, taRight);
    FGDI.DrawLine();
    FGDI.Font.Style := [];
  end;

begin
  if DataSet.IsEmpty then
    Exit;
  i := 1;
  c := 1;
  DataSet.First;
  DataSet.DisableControls;
  try
    id := DataSet.FieldByName(GrpKeyFd).AsLargeInt;
    DrawGrp;
    DataSet.Next;
    while not DataSet.Eof do begin
      if id <> DataSet.FieldByName(GrpKeyFd).AsLargeInt then begin
        AddDataSet(DataSet, False, AddHeaders, i, c, FirstField, FieldsCount);
        FGDI.NewLine;
        id := DataSet.FieldByName(GrpKeyFd).AsLargeInt;
        if NewPageGrp then
          FGDI.NewPage();
        DrawGrp;
        i := i + c;
        c := 1;
      end
      else
        Inc(c);
      DataSet.Next;
    end;
    AddDataSet(DataSet, False, AddHeaders, i, c, FirstField, FieldsCount);
  finally
    DataSet.EnableControls;
  end;
end;

procedure TPDF.AddGrid(var Grid: TDBGrid; WithBottomGrayLine: Boolean; AddHeaders: Boolean);
var
  linea: array of string;
  porcAnchosCols: array of Integer;
  i, c, sumaAnchos: Integer;
  // bm:TBookmark;
begin
  sumaAnchos := 0;
  c := Grid.Columns.Count;
  SetLength(linea, c);
  SetLength(porcAnchosCols, c);
  for i := 0 to c - 1 do
    Inc(sumaAnchos, Grid.Columns[i].Width);
  for i := 0 to Grid.Columns.Count - 1 do begin
    linea[i] := Grid.Columns[i].Title.Caption;
    porcAnchosCols[i] := Grid.Columns[i].Width * 100 div sumaAnchos;
  end;
  FGDI.AddColumns(porcAnchosCols);
  if AddHeaders then
    FGDI.AddColumnHeaders(linea, WithBottomGrayLine, True);
  for i := 0 to Grid.Columns.Count - 1 do
    case Ord(Grid.Columns[i].Alignment) of
      0: // taLeftJustify
        FGDI.SetColumnAlign(i, TColAlign.caLeft);
      1: // taRightJustify
        FGDI.SetColumnAlign(i, TColAlign.caRight);
      2: // taCenter
        FGDI.SetColumnAlign(i, TColAlign.caCenter);
    end;

  Grid.DataSource.DataSet.First;
  Grid.DataSource.DataSet.DisableControls;
  try
    while not Grid.DataSource.DataSet.Eof do begin
      for i := Low(linea) to High(linea) do
        linea[i] := Grid.Columns[i].Field.Text;
      FGDI.DrawTextAcrossCols(linea);
      Grid.DataSource.DataSet.Next;
    end;
  finally
    Grid.DataSource.DataSet.EnableControls;
  end;
end;

function TPDF.GetFont: TFont;
begin
  Result := FGDI.Font;
end;

end.
