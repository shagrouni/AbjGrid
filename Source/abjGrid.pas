{$IFDEF VER130}
{$DEFINE VER5U}
{$ENDIF}
{$IFDEF VER140}
{$DEFINE VER5U}
{$DEFINE VER6U}
{$ENDIF}
{$IFDEF VER150}
{$DEFINE VER5U}
{$DEFINE VER6U}
{$DEFINE VER7U}
{$ENDIF}
{$IFDEF VER160}
{$DEFINE VER8U}
{$ENDIF}
{$IFDEF VER170}
{$DEFINE VER9U}
{$ENDIF}
{$IFDEF VER180}
{$DEFINE VER10U}
{$ENDIF}

{$IFDEF VER185}
{$DEFINE VER11U}
{$ENDIF}
{$IFDEF VER190}
{$DEFINE VER11U}
{$ENDIF}
{$IFDEF VER200}
{$DEFINE VER12U}
{$ENDIF}

{$IFDEF VER210}
{$DEFINE VER14U}
{$ENDIF}

{$IFDEF VER220}
{$DEFINE VER15U}
{$ENDIF}


unit abjGrid {$IFDEF VER8U} platform {$ENDIF};

{
  By Khaled Shagrouni
  shagrouni@gmail.com

}
interface

uses

  Windows, SysUtils, Messages, Classes, Graphics, Controls, Forms,
  Grids, Clipbrd, ComObj, StdCtrls, DateUtils, StrUtils
{$IFDEF VER8U}, Types, Variants, System.ComponentModel{$ENDIF};

type

  TFieldState = record
    Width: integer;
    Position: integer;
    Caption: string;
  end;

const
  EmptyFieldState: TFieldState = (Width: 0; Position: 0; Caption: '');

type
  TBeforeRowEvent = procedure(Sender: TObject; ARow: Longint;
    var Accept: boolean) of object;
  TAfterRowEvent = procedure(Sender: TObject; ARow: Longint) of object;

  TExportTarget = (etText, etCSV, etExcel, etWord, etHTML, etXML);
  TExportEvent = procedure(Sender: TObject; Target: TExportTarget;
    ACol, ARow: integer; var S: string; var Cancel: boolean) of object;
  TSortEvent = procedure(Sender: TObject; ACol: integer; Ascending: boolean;
    var Cancel: boolean) of object;

  TBeforeEditEvent = procedure(Sender: TObject; ACol, ARow: integer;
    const Value: string) of object;
  TAfterEditEvent = procedure(Sender: TObject; ACol, ARow: integer;
    const Value: string) of object;
  TEditChangeEvent = procedure(Sender: TObject) of object;
  TBeforePasteEvent = procedure(Sender: TObject; ACol, ARow: integer;
    const Value: string; var Cancel: boolean) of object;

  TSaveOption = (soIncludeHeaders, soFixedLenght);
  THackClipboard = class(TClipboard);

  TAbjGrid = class;

  TGridEditor = Class(TMemo)
  private
    procedure CMExit(var Message: TCMExit); message CM_EXIT;
    procedure CNCommand(var Message: TWMCommand); message CN_COMMAND;
  public
    Grid: TAbjGrid;
    Modified: boolean;
    constructor Create(AOwner: TComponent); override;
    procedure Put_Content_In_Cell;
    procedure Show_Edit_In_Cell(ACol, ARow: integer; CaretUnderMosue: boolean;
      const AText: string); overload;
    procedure Show_Edit_In_Cell; overload;
    procedure DoResize;
    procedure DoHide;

    function CharIndexFromPoint(P: TPoint): integer;
  published
    procedure KeyUp(var Key: Word; Shift: TShiftState); override;
  End;

  TAbjGrid = class(TStringGrid)
  private
    FGridOnBuild: boolean;
    // FColMoving: boolean;
    FCurSortCol: integer;
    FFlat: boolean;
    FAlternatingBackColor: TColor;
    FSelectColor: TColor;
    FFSelectColor: TColor;
    FFixedSelectColor: TColor;
    FShadowColor, FShadowColor2, FLightColor, FLightColor2: TColor;
    FColState: array [0 .. 255] of integer;
    FFieldState: array [0 .. 255] of TFieldState;
    FAutoRowNumber: boolean;
    FAutoColWidth: boolean;
    FFieldDelimiter: string;
    // FCellValue: string;
    FCommentDelimiter: string;
    FAutoRowHeight: boolean;
    FAutoRowIncrement: boolean;
    FAllowRowDelete: boolean;
    FAllowRowInsert: boolean;
    FFixedColor: TColor;
    FBeforeRowDelete: TBeforeRowEvent;
    FAfterRowDelete: TAfterRowEvent;
    FBeforeRowInsert: TBeforeRowEvent;
    FAfterRowInsert: TAfterRowEvent;
    FBeforeEditEvent: TBeforeEditEvent;
    FAfterEditEvent: TAfterEditEvent;
    FBeforePasteEvent: TBeforePasteEvent;
    FBackColor: TColor;
    FRowHeaderWidth: integer;
    FRowHeaderFont: TFont;
    FColHeaderFont: TFont;
    FAllowSorting: boolean;
    FSampleData: boolean;
    FGridLineColor: TColor;
    FExport: TExportEvent;
    FColHeaderHeight: integer;
    FgoEditInOptions: boolean;
    FTitle: TStrings;
    FonColumnResize: TNotifyEvent;
    FonRowResize: TNotifyEvent;

    FBeforeSort: TSortEvent;
    FAfterSort: TSortEvent;
    FAllowPaste: boolean;
    FAllowEdit: boolean;
    FEditor: TGridEditor;
    FEditChange: TEditChangeEvent;

    procedure DoOnColumnResize;
    procedure DoOnRowResize;

    procedure SetAutoRowNumber(const Value: boolean);
    procedure SetAutoRowIncrement(const Value: boolean);
    procedure SetFieldDelimiter(const Value: string);
    procedure SetCommentDelimiter(const Value: string);

    procedure SetAllowRowDelete(const Value: boolean);
    procedure SetAllowRowInsert(const Value: boolean);
    procedure SetAllowSorting(const Value: boolean);
    procedure SetSampleData(const Value: boolean);

    procedure SetAutoRowHeight(const Value: boolean);
    procedure SetAutoColWidth(const Value: boolean);
    procedure SetColHeaderHeight(const Value: integer);
    procedure SetRowHeaderWidth(const Value: integer);

    procedure SetAlternatingBackColor(const Value: TColor);
    procedure SetBackColor(const Value: TColor);
    procedure SetFixedColor(const Value: TColor);
    procedure SetGridLineColor(const Value: TColor);
    procedure SetSelectColor(const Value: TColor);

    procedure SetColHeaderFont(const Value: TFont);
    procedure SetRowHeaderFont(const Value: TFont);

    procedure SetFlat(const Value: boolean);
    procedure SetTitle(const Value: TStrings);
    procedure WMCopy(var Message: TMessage); message WM_COPY;
    procedure WMPaste(var Message: TMessage); message WM_PASTE;
{$IFNDEF VER6U}
    procedure ChangeGridOrientation(RightToLeftOrientation: boolean);
{$ENDIF}
    procedure SetAllowPaste(const Value: boolean);
    procedure SetAllowEdit(const Value: boolean);
    procedure SetEditor(const Value: TGridEditor);
  protected

    // procedure ResizeCol(Index: Longint; OldSize, NewSize: Integer); override;
    procedure DrawCell(ACol, ARow: integer; Rect: TRect;
      State: TGridDrawState); override;

    function DrawFixedGridLine(ACol, ARow: integer; Selected: boolean): TRect;
    function DrawGridLine(ACol, ARow: integer; Selected: boolean): TRect;
    { procedure KeyDown(var Key: Word; Shift: TShiftState); override;
      procedure DoEnter; override;
      procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
      procedure MouseUp(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;
      procedure MouseDown(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;
    }

    function GetEditText(ACol, ARow: Longint): string; override;
    function SelectCell(ACol, ARow: Longint): boolean; override;
    procedure ColumnMoved(FromIndex, ToIndex: Longint); override;

    procedure GridDrawCell(ACol, ARow: integer; Rect: TRect;
      State: TGridDrawState);
    procedure DrawHeaderRow(ACol, ARow: integer; Selected: boolean);
    procedure DrawHeaderRowLtoR(ACol, ARow: integer; Selected: boolean);
    procedure DrawHeaderRowRtoL(ACol, ARow: integer; Selected: boolean);
    procedure xDrawHeaderRow(ACol, ARow: integer; Selected: boolean);

    procedure DrawHeaderCol(ACol, ARow: integer; Rect: TRect;
      Selected: boolean);
    procedure DrawHeaderColRToL(ACol, ARow: integer; Selected: boolean);
    procedure DrawHeaderColLToR(ACol, ARow: integer; Selected: boolean);
    procedure xDrawHeaderCol(ACol, ARow: integer; Rect: TRect;
      Selected: boolean);

    function GetCellHeight(ACol, ARow: integer): integer;

    function GetCellText(ACol, ARow: integer): string;
    // procedure DrawCellFocus(ACol, ARow:integer); //moved to puplic
    function IsColInSelect(ACol: integer): boolean;
    function GetFieldDelimiter: string;
    function GetShadeColor(clr: TColor; Value: integer): TColor;
    procedure InitGrid;
    procedure InitColors;
    procedure Invalidate_AbjGrid;
    procedure DrawSort(ACol: integer; SortDec: boolean; AColor: TColor);
    procedure ColWidthsChanged; override;
    procedure RowHeightsChanged; override;

  public
    GridState: TGridState;
    GlobalCellMargin: NativeInt;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure CalcSizingState(X, Y: integer; var State: TGridState;
      var Index: Longint; var SizingPos, SizingOfs: integer;
      var FixedInfo: TGridDrawInfo); override;

    property InplaceEditor;
    procedure SetColsWidths(Arr: array of integer);
    procedure SetColsCaptions(Arr: array of string);
    function ColByCaption(Caption: string): integer;
    function iCol(Caption: string): integer;
    procedure RowDelete(ARow: integer);
    procedure ColDelete(ACol: integer); overload;
    procedure ColDelete(X, Y: integer); overload;
    procedure RowInsert(ARow: integer);
    procedure ColInsert(ACol: integer); overload;
    procedure ColInsert(X, Y: integer); overload;
    procedure RowMove(FromRow, ToRow: integer);
    procedure ColMove(FromCol, ToCol: integer);
    procedure DeleteRows;
    procedure DeleteCols;
    procedure ClearRows;
    procedure ClearCols;
    procedure ClearCol(ACol: integer);
    procedure ClearFixedCols;
    procedure ClearFixedRows;
    procedure Clear;
    procedure ClearAll;
    procedure ExchangeCells(Cell_1, Cell_2: TPoint);
    procedure RowSelect(ARow: integer);
    procedure LoadFromStringList(var L: TStringList; DrawHeader: boolean;
      Index: integer);
    // procedure LoadFromStringList2(var L: TStringList; DrawHeader: boolean; Index: integer);
    procedure LoadFromFile(const FileName: string; HasHeader: boolean);
    procedure SaveToFile(const FileName: string; IncludeHeader: boolean); overload;
    procedure SaveToFile(const FileName: string;
      SaveOption: TSaveOption); overload;
    procedure SaveToFile(const FileName: string; SaveOption: TSaveOption;
      Target: TExportTarget); overload;

{$IFDEF VER5U} procedure SaveToExcel(const FileName, Title, Footer: string; ShowApp: boolean); {$ENDIF}
{$IFDEF UNICODE}
    procedure SaveToExcel(const FileName, Title, Footer: string;
      ShowApp: boolean);
{$ENDIF}
    procedure SaveToHTML(const FileName, Title, Footer: string);
    procedure SaveToXML(const FileName: string; Metadat: boolean);
    procedure SortGrid(ACol: integer); overload;
    procedure SortGrid(ACol: integer; SortDec: boolean); overload;
    function IndexOfCol(ACol: integer; S: string; Start: integer): integer;
    procedure CopyToClipboard;
    procedure PasteFromClipboard;
    procedure MoveToNext;
    procedure InvalidateCell(ACol, ARow: Longint);
    procedure InvalidateRowA(ARow: Longint);

    function RowUnderMouse(X, Y: integer): integer;
    function ColUnderMouse(X, Y: integer): integer;
    function CellUnderMouse: TPoint; overload;
    function CellUnderMouse(var ACol, ARow: NativeInt): TPoint; overload;

    procedure RowAutoHeight(ARow: Longint; Value: boolean);
    procedure SelectAll;
    function GetParentForm: TCustomForm;
    function DetermineCellHeight(S: string; ACol, ARow: integer): integer;
    function Selected(ACol, ARow: integer): boolean;
    function SelectionRect: TRect;
    procedure DrawCellFocus(ACol, ARow: integer);
    function NewColor(clr: TColor; Value: integer): TColor;
    procedure DrawCellText(Rect: TRect; const S: string;
      const Format: NativeInt);
  published
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure KeyUp(var Key: Word; Shift: TShiftState); override;
    procedure DblClick; override;
    procedure DoEnter; override;
    procedure MouseMove(Shift: TShiftState; X, Y: integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState;
      X, Y: integer); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState;
      X, Y: integer); override;

    property AlternatingBackColor: TColor read FAlternatingBackColor
      write SetAlternatingBackColor;
    property AllowRowDelete: boolean read FAllowRowDelete
      write SetAllowRowDelete;
    property AllowRowInsert: boolean read FAllowRowInsert
      write SetAllowRowInsert;
    property AllowSorting: boolean read FAllowSorting write SetAllowSorting;
    property AllowPaste: boolean read FAllowPaste write SetAllowPaste;
    property AllowEdit: boolean read FAllowEdit write SetAllowEdit;
    property AutoRowNumber: boolean read FAutoRowNumber write SetAutoRowNumber;
    property AutoColWidth: boolean read FAutoColWidth write SetAutoColWidth;
    property AutoRowHeight: boolean read FAutoRowHeight write SetAutoRowHeight;
    property AutoRowIncrement: boolean read FAutoRowIncrement
      write SetAutoRowIncrement;
    property BackColor: TColor read FBackColor write SetBackColor;
    property CommentDelimiter: string read FCommentDelimiter
      write SetCommentDelimiter; // new
    property FieldDelimiter: string read FFieldDelimiter
      write SetFieldDelimiter; // new
    property Flat: boolean read FFlat write SetFlat;
    property FixedColor: TColor read FFixedColor write SetFixedColor
      default clBtnFace;
    property SelectColor: TColor read FSelectColor write SetSelectColor;
    property GridLineColor: TColor read FGridLineColor write SetGridLineColor;
    property Title: TStrings read FTitle write SetTitle;
    property BeforeRowDelete: TBeforeRowEvent read FBeforeRowDelete
      write FBeforeRowDelete;
    property AfterRowDelete: TAfterRowEvent read FAfterRowDelete
      write FAfterRowDelete;

    property BeforeEdit: TBeforeEditEvent read FBeforeEditEvent
      write FBeforeEditEvent;
    property AfterEdit: TAfterEditEvent read FAfterEditEvent
      write FAfterEditEvent;
    property BeforePaste: TBeforePasteEvent read FBeforePasteEvent
      write FBeforePasteEvent;

    property EditChange: TEditChangeEvent read FEditChange write FEditChange;

    property BeforeRowInsert: TBeforeRowEvent read FBeforeRowInsert
      write FBeforeRowInsert;
    property AfterRowInsert: TAfterRowEvent read FAfterRowInsert
      write FAfterRowInsert;
    property RowHeaderWidth: integer read FRowHeaderWidth
      write SetRowHeaderWidth;
    property ColHeaderHeight: integer read FColHeaderHeight
      write SetColHeaderHeight;
    property RowHeaderFont: TFont read FRowHeaderFont write SetRowHeaderFont;
    property ColHeaderFont: TFont read FColHeaderFont write SetColHeaderFont;
    property SampleData: boolean read FSampleData write SetSampleData;
    property OnExport: TExportEvent read FExport write FExport;
    property BeforeSort: TSortEvent read FBeforeSort write FBeforeSort;
    property AfterSort: TSortEvent read FAfterSort write FAfterSort;
    property OnColumnResize: TNotifyEvent read FonColumnResize
      write FonColumnResize;
    property OnRowResize: TNotifyEvent read FonRowResize write FonRowResize;
    property Editor: TGridEditor read FEditor write SetEditor;

  end;

procedure Register;
function CustomSort(List: TStringList; Index1, Index2: integer): integer;

function StrValue(S: string): Real;
function ValReal(S: string): Real;
function ValInt(S: string): integer;
function HTMLColor(Color: TColor): string;
function IsDate(d: string): boolean;
function GetDate(S: string; var DT: TDateTime): boolean;
function KVal(S: string): integer;
procedure Clipboard_SetBuffer(Format: Word; var Buffer; Size: integer);
function DrawTextX(Handle: THandle; WS: WideString; ARect: TRect;
  Flags: Longint): integer;

procedure Gradient(Col_D, Col_L: TColor; Bmp: TBitmap);
function BlendColor(Color1, Color2: TColor; A: Byte): TColor;

var
  CellValue: string;
  FSortDec: boolean;

implementation

uses
  uGridSort;

procedure Register;
begin
  RegisterComponents('Abjad Controls', [TAbjGrid]);
end;

function MonthDayYearFormat(d: TDateTime): string; overload;
var
  Year, Month, Day: Word;
begin
  DecodeDate(d, Year, Month, Day);

  result := IntToStr(Month) + '/' + IntToStr(Day) + '/' + IntToStr(Year);

end;

function MonthDayYearFormat(S: string): string; overload;
begin
  result := MonthDayYearFormat(StrToDate(S));
end;


constructor TAbjGrid.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FFlat := True;
  BorderStyle := bsNone;
  DefaultDrawing := false;
  DefaultRowHeight := 17;
  RowHeaderWidth := 35;
  ColHeaderHeight := 17;
  GlobalCellMargin := 0;
  FCommentDelimiter := ';';
  FFieldDelimiter := 'TAB';

  FRowHeaderFont := TFont.Create;
  FColHeaderFont := TFont.Create;

  FRowHeaderFont.Assign(self.Font);
  FColHeaderFont.Assign(self.Font);

  FCurSortCol := -1;

  FFixedColor := clBtnFace;
  FAlternatingBackColor := Color;
  FBackColor := Color;
  FGridLineColor := clBtnShadow;
  InitColors;
  FSelectColor := FFSelectColor;
  FTitle := TStringList.Create;

  Editor := TGridEditor.Create(self);
  Editor.Visible := false;

end;

destructor TAbjGrid.Destroy;
begin
  FRowHeaderFont.Free;
  FColHeaderFont.Free;
  FTitle.Free;

  Editor.Free;

  inherited Destroy;
end;

// -----------
procedure TAbjGrid.InitGrid;
var
  C: Word;
  txt: string;
  i: integer;
begin

  for i := Low(FColState) to high(FColState) do
    FColState[i] := 0;
  for i := Low(FFieldState) to high(FFieldState) do
  begin
    FFieldState[i] := EmptyFieldState;
  end;

  FColState[0] := 0;

  Canvas.Font := Font;
  Canvas.Font.Style := Canvas.Font.Style + [fsBold];
  DefaultRowHeight := Canvas.TextHeight('K|') + 6;
  ColWidths[0] := FRowHeaderWidth;
  RowHeights[0] := FColHeaderHeight;

  if ColCount < 2 then
    ColCount := 2;

  for C := 1 to ColCount - 1 do
  begin
    txt := Cells[C, 0];
    if FAutoColWidth then // New
      if ColWidths[C] <= (Canvas.TextWidth(txt) + 10) then
        ColWidths[C] := Canvas.TextWidth(txt) + 10;
  end;

  col := 1;
  row := 1;
end;

procedure TAbjGrid.InitColors;
begin
  FShadowColor := GetShadeColor(FixedColor, 60);
  FShadowColor2 := GetShadeColor(FixedColor, 99);
  FLightColor := NewColor(FixedColor, 60);
  FLightColor2 := GetShadeColor(FLightColor, 10);
  FFSelectColor := NewColor(clHighlight, 68);
  FFixedSelectColor := NewColor(clHighlight, 74);

end;

procedure TAbjGrid.ColDelete(ACol: integer);
begin

  if (ColCount = 2) and (ACol = 1) then
  begin
    Cols[ACol].Clear;
    exit;
  end;

  if ColCount > 2 then
  begin
    Cols[ACol].Clear;
    DeleteColumn(ACol);

    if ACol < ColCount then
      col := ACol;
  end;

end;

procedure TAbjGrid.ColDelete(X, Y: integer);
var
  ACol, ARow: integer;
begin
  MouseToCell(X, Y, ACol, ARow);
  if ACol <> -1 then
    ColDelete(ACol);

end;

procedure TAbjGrid.SetColsWidths(Arr: array of integer);
var
  i: integer;
begin
  for i := 0 to High(Arr) do
  begin
    if i < ColCount then
      if Arr[i] <> -2 then
        ColWidths[i] := Arr[i];
  end;
end;

procedure TAbjGrid.SetColsCaptions(Arr: array of string);
var
  i: integer;
begin
  for i := 0 to High(Arr) do
  begin
    if i < ColCount then
      if Arr[i] <> '@' then
        // if Cells[i, 0] <> Arr[i] then
        Cells[i, 0] := Arr[i];

  end;

end;

function TAbjGrid.ColByCaption(Caption: string): integer;
var
  i: integer;
begin
  result := 0;
  for i := 0 to ColCount - 1 do
  begin
    if UpperCase(Cells[i, 0]) = UpperCase(Caption) then
    begin
      result := i;
      exit;
    end;
  end;

end;

function TAbjGrid.iCol(Caption: string): integer;
begin
  result := ColByCaption(Caption);
end;

procedure TAbjGrid.RowAutoHeight(ARow: integer; Value: boolean);
var
  H: integer;
  MaxH: integer;
  i: integer;
begin
  if Value then
  begin
    MaxH := DefaultRowHeight;
    for i := FixedCols to ColCount - 1 do
    begin
      H := GetCellHeight(i, ARow);
      if H > MaxH then
        MaxH := H;

    end;
    if MaxH > RowHeights[ARow] then
      RowHeights[ARow] := MaxH;
  end
  else
    RowHeights[ARow] := DefaultRowHeight;

end;

procedure TAbjGrid.RowDelete(ARow: integer);
var
  Accept: boolean;
begin
  if RowCount <= ARow then
    exit;

  Accept := True;
  if assigned(BeforeRowDelete) then
    BeforeRowDelete(self, row, Accept);
  if Accept then
  begin
    if (ARow = FixedRows) and (RowCount = (ARow + 1)) then
    begin
      Rows[ARow].Clear;
      // exit;
    end
    else
    begin
      Rows[ARow].Clear;
      DeleteRow(ARow);
      if ARow < RowCount then
        row := ARow;
    end;

    //
    if assigned(AfterRowDelete) then
      AfterRowDelete(self, row);
  end;

end;

procedure TAbjGrid.DeleteRows;
var
  i: integer;
begin
  for i := RowCount - 1 downto FixedRows do
    RowDelete(i);
end;

procedure TAbjGrid.DblClick;
begin
  if FAllowEdit then
    Editor.Show_Edit_In_Cell;
  inherited;

end;

procedure TAbjGrid.DeleteCols;
var
  i: integer;
begin
  for i := ColCount - 1 downto FixedCols do
    ColDelete(i);
end;

procedure TAbjGrid.ClearCols;
var
  row, col: integer;
begin
  for col := ColCount - 1 downto FixedCols do
  begin
    for row := FixedRows to RowCount - 1 do
      Cells[col, row] := '';
    if Cols[col].Objects[col] <> nil then
      Cols[col].Objects[col] := nil;
  end;

end;

procedure TAbjGrid.ClearCol(ACol: integer);
var
  row: integer;
begin
  for row := RowCount - 1 downto 0 do
  begin
    Cells[ACol, row] := '';
    if assigned(Cols[ACol].Objects[ACol]) then
      Cols[ACol].Objects[ACol] := nil;
  end;

end;

procedure TAbjGrid.ClearRows;
var
  row, col: integer;
begin
  for row := RowCount - 1 downto FixedRows do
  begin
    for col := FixedCols to ColCount - 1 do
      Cells[col, row] := '';
    if assigned(Rows[row].Objects[row]) then
      Rows[row].Objects[row] := nil;
    RowHeights[row] := DefaultRowHeight;
  end;
end;

procedure TAbjGrid.ClearFixedRows;
var
  row: integer;
begin
  for row := FixedRows - 1 downto 0 do
    Rows[row].Clear;
end;

procedure TAbjGrid.ClearFixedCols;
var
  col: integer;
begin
  for col := FixedCols - 1 downto 0 do
    Cols[col].Clear;
end;

procedure TAbjGrid.RowInsert(ARow: integer);
var
  i: integer;
  Accept: boolean;
begin

  Accept := True;
  if assigned(BeforeRowInsert) then
    BeforeRowInsert(self, row, Accept);
  if Accept then
  begin
    // RowInsert(Row);
    // --------------
    if (ARow < FixedRows) then
      ARow := FixedRows;
    RowCount := RowCount + 1;
    Rows[RowCount - 1].Clear;
    for i := RowCount - 1 downto ARow + 1 do
      inherited RowMoved(i, i - 1);
    // ------------------
    if assigned(AfterRowInsert) then
      AfterRowInsert(self, row);
  end;

end;

procedure TAbjGrid.ColInsert(ACol: integer);
var
  i: integer;
begin
  if (ACol < FixedRows) then
    ACol := FixedCols;
  ColCount := ColCount + 1;
  Cols[ColCount - 1].Clear;
  for i := ColCount - 1 downto ACol + 1 do
    inherited ColumnMoved(i, i - 1);

end;

procedure TAbjGrid.ColInsert(X, Y: integer);
var
  ACol, ARow: integer;
begin
  MouseToCell(X, Y, ACol, ARow);
  if ACol <> -1 then
    ColInsert(ACol);
end;

procedure TAbjGrid.ColMove(FromCol, ToCol: integer);
begin
  inherited ColumnMoved(FromCol, ToCol);
end;

procedure TAbjGrid.RowMove(FromRow, ToRow: integer);
var
  ObjFrom, ObjTo: TObject;
begin
  ObjFrom := Rows[FromRow].Objects[FromRow];
  ObjTo := Rows[ToRow].Objects[ToRow];
  inherited RowMoved(FromRow, ToRow);
  Rows[FromRow].Objects[FromRow] := ObjTo;
  Rows[ToRow].Objects[ToRow] := ObjFrom;
end;

procedure TAbjGrid.Clear;
begin
  ClearRows;
  ClearCols;
end;

procedure TAbjGrid.ClearAll;
begin
  DeleteRows;
  DeleteCols;
  ClearFixedCols;
  ClearFixedRows;
end;

procedure TAbjGrid.KeyDown(var Key: Word; Shift: TShiftState);
var
  ACol, ARow: integer;
begin
  if row = RowCount - 1 then
    if FAutoRowIncrement then
      if Key = VK_DOWN then
        RowCount := RowCount + 1;

  inherited KeyDown(Key, Shift);

  if (Key = VK_ESCAPE) and EditorMode then
  begin
    EditorMode := false;
    Cells[col, row] := CellValue;
    CellValue := '';
  end;

  if (ssShift in Shift) and (Key = VK_DELETE) then
    // if (Key = VK_DELETE) then
    for ARow := Selection.Top to Selection.Bottom do
      for ACol := Selection.Left to Selection.Right do
        Cells[ACol, ARow] := '';

  if (ssCtrl in Shift) and (Key = VK_DELETE) then // Ctrl+delete
    if FAllowRowDelete then
      RowDelete(row);

  if FAllowRowInsert then
    if (ssAlt in Shift) and (Key = vk_Insert) then // Alt+Insert
      RowInsert(row);

  if ((ssCtrl in Shift) and (Key = vk_Insert)) or
    ((ssCtrl in Shift) and (Key = ord('c'))) or
    ((ssCtrl in Shift) and (Key = ord('C'))) then
  begin
    CopyToClipboard;
  end;

  if (ssShift in Shift) and (Key = vk_Insert) or
    ((ssCtrl in Shift) and (Key = ord('v'))) or
    ((ssCtrl in Shift) and (Key = ord('V'))) then
  begin
    PasteFromClipboard;
  end;

  if (ssCtrl in Shift) and (ssAlt in Shift) and (ssShift in Shift) then
    if BiDiMode = bdRightToLeft then
      BiDiMode := bdLeftToRight
    else
      BiDiMode := bdRightToLeft;
end;

procedure TAbjGrid.KeyUp(var Key: Word; Shift: TShiftState);
begin

  inherited KeyUp(Key, Shift);
  if FAllowEdit then
  begin
    case Key of
      13:
        begin
          Editor.Show_Edit_In_Cell(col, row, false, '');
        end;
      vk_F2:
        begin
          Editor.Show_Edit_In_Cell(col, row, false, '');
        end;
    end;
  end;

end;

procedure TAbjGrid.MouseMove(Shift: TShiftState; X, Y: integer);

begin
  inherited MouseMove(Shift, X, Y);
end;

procedure TAbjGrid.MouseDown(Button: TMouseButton; Shift: TShiftState;
  X, Y: integer);
var
  ACol, ARow: NativeInt;
begin
  CellUnderMouse(ACol, ARow);
  if (ACol <> col) and (ARow <> row) then
    Editor.DoHide;
  inherited;

end;

procedure TAbjGrid.MouseUp(Button: TMouseButton; Shift: TShiftState;
  X, Y: integer);
var
  ACol, ACol2, ARow: integer;
begin

  if FgoEditInOptions then
    Options := Options + [goEditing]
  else
    Options := Options - [goEditing];

  inherited MouseUp(Button, Shift, X, Y);
  if not FAllowSorting then
    exit;
  if (gsColSizing in [GridState]) or (gsColMoving in [GridState]) then
    exit;
  MouseToCell(X, Y, ACol, ARow);
  if (FixedRows > 0) and (ARow < FixedRows) and (ACol >= FixedCols) then
  begin
    MouseToCell(X + 20, Y, ACol2, ARow);
    if ACol2 <> ACol then
    begin
      if FCurSortCol = ACol then
        FSortDec := not FSortDec
      else
        FSortDec := false;

      FCurSortCol := ACol;
      SortGrid(ACol);
      if CanFocus then
        SetFocus;
    end;

  end;

end;

function TAbjGrid.GetEditText(ACol, ARow: integer): string;
begin
  CellValue := Cells[col, row];
  result := CellValue;

  if assigned(OnGetEditText) then
    OnGetEditText(self, ACol, ARow, result);

end;

/// /=======

procedure TAbjGrid.GridDrawCell(ACol, ARow: integer; Rect: TRect;
  State: TGridDrawState);
var
  // sAlig : integer;
  txt: string;
  // fCol: integer;
  vRect: TRect;
  Lang: integer;
  flag: integer;
  H: integer;
  Done: boolean;
begin

  Lang := 0;
  Done := false;

  vRect := Rect;

  txt := Cells[ACol, ARow];

  if (ACol < FixedCols) { and ((FixedCols < ACol)) } then
  begin
    if FFlat then
    begin
      Inc(vRect.Bottom, 1);
      Dec(vRect.Top, 1);
    end;
    DrawHeaderRow(ACol, ARow, (ARow = row));
    exit;
  end;

  if ARow < FixedRows then
  begin
    DrawHeaderCol(ACol, ARow, vRect, (ACol = col));
    exit;
  end;

  Canvas.Font.Style := [];
  Canvas.Font.Color := clBlack;
  if FAutoRowHeight then
    if not(csDesigning in ComponentState) then
    begin
      H := GetCellHeight(ACol, ARow); // must check all cells in the row
      if H > RowHeights[ARow] then
        RowHeights[ARow] := H;
    end;

  if UseRightToLeftAlignment then
    Lang := Lang + DT_RIGHT + DT_RTLREADING;

  if (gdSelected in State) and not(gdFocused in State) then
  begin
    Canvas.Brush.Color := FSelectColor;
  end
  else
  begin
    if (ARow mod 2) = 1 then
      Canvas.Brush.Color := FBackColor // Color
    else
      Canvas.Brush.Color := FAlternatingBackColor;
  end;


  FGridOnBuild := True;
  Canvas.FillRect(Rect);

  if UseRightToLeftAlignment then
  begin
    Dec(vRect.Right, 5);
    Inc(vRect.Left, 3);
  end
  else
  begin
    Inc(vRect.Left, 5);
    Dec(vRect.Right, 4);
  end;
  Inc(vRect.Top, 2);
  Dec(vRect.Bottom, 1);

  Canvas.Font.Assign(Font);
  if not Done then
  begin
    if RowHeights[ARow] > (Canvas.TextHeight(txt) * 2) then
      flag := DT_WORDBREAK
    else
      flag := 0;
    DrawTextX(Canvas.Handle, { ' '+ } txt, vRect, Lang + flag);
    // ad space for arabic text
  end;

  if (gdFocused in State) then
  begin
    if goDrawFocusSelected in Options then
      DrawCellFocus(ACol, ARow)
    else
      Canvas.DrawFocusRect(Rect);
  end;

  FGridOnBuild := false;

end;

// ==============================================

procedure TAbjGrid.DrawHeaderRowLtoR(ACol, ARow: integer; Selected: boolean);
var
  ARect { , fRect } : TRect;
begin

  ARect := CellRect(ACol, ARow);
  Dec(ARect.Top, 1);

  if (goFixedVertLine in Options) or (goVertLine in Options) then
    if ACol > 0 then
      Dec(ARect.Left, 1);

  if FFlat and Selected then
    Canvas.Brush.Color := FSelectColor // FFixedSelectColor
  else
    Canvas.Brush.Color := FixedColor;

  if (ARow = row + 1) and not(FFlat) then
    Inc(ARect.Top, 1); // do not paint over selected row above
  Canvas.FillRect(ARect);

  ARect := DrawFixedGridLine(ACol, ARow, Selected);

  if (not FFlat) then
    if (Selected) then
    begin
      // if (not (goFixedVertLine in Options)) and (not (goFixedHorzLine in Options)) then
      // Dec(ARect.Bottom,1);

      Canvas.MoveTo(ARect.Right, ARect.Top);
      Canvas.Pen.Color := FShadowColor2;
      Canvas.LineTo(ARect.Right, ARect.Bottom);
      Canvas.LineTo(ARect.Left - 1, ARect.Bottom);

      if (ACol = FixedCols - 1) or (goFixedVertLine in Options) then
      begin
        Canvas.Pen.Color := FShadowColor;
        Canvas.MoveTo(ARect.Right - 1, ARect.Top + 1);
        Canvas.LineTo(ARect.Right - 1, ARect.Bottom - 1);
      end;

      Canvas.Pen.Color := FShadowColor;
      Canvas.MoveTo(ARect.Right - 1, ARect.Bottom - 1);
      Canvas.LineTo(ARect.Left + 1, ARect.Bottom - 1);

      if (ACol = 0) or (goFixedVertLine in Options) then
      begin
        Canvas.Pen.Color := FLightColor;
        Canvas.MoveTo(ARect.Left, ARect.Bottom - 1);
        Canvas.LineTo(ARect.Left, ARect.Top);
      end;

      Canvas.Pen.Color := FLightColor;
      Canvas.MoveTo(ARect.Left, ARect.Top);
      Canvas.LineTo(ARect.Right - 1, ARect.Top);

      Canvas.Pixels[ARect.Right - 1, ARect.Top] := FFixedColor;
    end
    else
    begin

      if (ACol < FixedCols) and (ACol <> FixedCols - 1) then
        Dec(ARect.Right, 1);
      if not(goHorzLine in Options) and not(goFixedHorzLine in Options) then
      begin

        if ARow = RowCount - 1 then
          Dec(ARect.Bottom, 1);
        if ARow = 0 then
          Dec(ARect.Top, 1)
        else
          Inc(ARect.Top, 1);
      end;

      if ARow = 0 then
        Dec(ARect.Top, 1)
      else if (goFixedHorzLine in Options) then
      begin
        Canvas.Pen.Color := FShadowColor2;
        Canvas.MoveTo(ARect.Left, ARect.Top);
        Canvas.LineTo(ARect.Right, ARect.Top);
      end;

      if (ACol = 0) or (goFixedVertLine in Options) then
      begin
        Canvas.Pen.Color := FLightColor2;
        Canvas.MoveTo(ARect.Left, ARect.Bottom - 1);
        Canvas.LineTo(ARect.Left, ARect.Top + 1);
      end;

      if (ARow = 0) or (goFixedHorzLine in Options) then
      begin
        Canvas.Pen.Color := FLightColor2;
        Canvas.MoveTo(ARect.Left, ARect.Top + 1);
        Canvas.LineTo(ARect.Right, ARect.Top + 1);
      end;

      if (ACol = FixedCols - 1) or (goFixedVertLine in Options) then
      begin
        Canvas.Pen.Color := FShadowColor;
        Canvas.MoveTo(ARect.Right, ARect.Top + 1);
        Canvas.LineTo(ARect.Right, ARect.Bottom);
      end;

      if (ARow = RowCount - 1) or (goFixedHorzLine in Options) then
      begin
        Canvas.Pen.Color := FShadowColor;
        Canvas.MoveTo(ARect.Right, ARect.Bottom);
        Canvas.LineTo(ARect.Left - 1, ARect.Bottom);
      end;

      if not(goFixedHorzLine in Options) then
      begin
        if (goFixedVertLine in Options) then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left, ARect.Top - 1);
          Canvas.LineTo(ARect.Left, ARect.Bottom + 1);

          Canvas.Pen.Color := FShadowColor;
          Canvas.MoveTo(ARect.Right, ARect.Top - 1);
          Canvas.LineTo(ARect.Right, ARect.Bottom + 1);

        end;
      end;

      if not(goFixedVertLine in Options) then
      begin
        if (goFixedHorzLine in Options) then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left, ARect.Top + 1);
          Canvas.LineTo(ARect.Right + 1, ARect.Top + 1);

          Canvas.Pen.Color := FShadowColor;
          Canvas.MoveTo(ARect.Left, ARect.Top);
          Canvas.LineTo(ARect.Right + 1, ARect.Top);
          Canvas.MoveTo(ARect.Left, ARect.Bottom);
          Canvas.LineTo(ARect.Right + 1, ARect.Bottom);

        end;
      end;
      if (not(goFixedVertLine in Options)) and (not(goFixedHorzLine in Options))
      then
      begin
        if ACol = 0 then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left, ARect.Top);
          Canvas.LineTo(ARect.Left, ARect.Bottom);
        end;
        if ACol = FixedCols - 1 then
        begin
          Canvas.Pen.Color := FShadowColor;
          Canvas.MoveTo(ARect.Right, ARect.Top);
          Canvas.LineTo(ARect.Right, ARect.Bottom);
        end;
        if ARow = 0 then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left, ARect.Top + 1);
          Canvas.LineTo(ARect.Right + 1, ARect.Top + 1);
        end;
      end;
    end;

end;

procedure TAbjGrid.DrawHeaderRowRtoL(ACol, ARow: integer; Selected: boolean);
var
  ARect { , fRect } : TRect;
begin
  ARect := CellRect(ACol, ARow);
  Dec(ARect.Top, 1);

  if (goFixedVertLine in Options) or (goVertLine in Options) then
    if ACol > 0 then
      // Dec(ARect.Left, 1);   //LeftToRight
      Inc(ARect.Right, 1); // righttoleft

  if FFlat and Selected then
    Canvas.Brush.Color := FSelectColor // FFixedSelectColor
  else
    Canvas.Brush.Color := FixedColor;

  if (ARow = row + 1) and not(FFlat) then
    Inc(ARect.Top, 1); // do not paint over selected row above
  Canvas.FillRect(ARect);

  ARect := DrawFixedGridLine(ACol, ARow, Selected);

  if (not FFlat) then
    if (Selected) then
    begin
      if ACol < FixedCols - 1 then
        Inc(ARect.Left, 1);

      if ACol = 0 then
      begin
        Canvas.Pen.Color := FShadowColor2;
        Canvas.MoveTo(ARect.Right, ARect.Top);
        Canvas.LineTo(ARect.Right, ARect.Bottom);
      end
      else
        Inc(ARect.Right, 1);

      Canvas.Pen.Color := FShadowColor2;
      Canvas.MoveTo(ARect.Right, ARect.Bottom);
      Canvas.LineTo(ARect.Left - 1, ARect.Bottom);

      Canvas.Pen.Color := FShadowColor;
      Canvas.MoveTo(ARect.Right - 1, ARect.Top + 1);
      Canvas.LineTo(ARect.Right - 1, ARect.Bottom - 1);
      Canvas.Pen.Color := FShadowColor;
      Canvas.MoveTo(ARect.Right - 1, ARect.Bottom - 1);
      Canvas.LineTo(ARect.Left + 1, ARect.Bottom - 1);

      Canvas.Pen.Color := FLightColor;
      Canvas.MoveTo(ARect.Left, ARect.Bottom - 1);
      Canvas.LineTo(ARect.Left, ARect.Top);

      Canvas.Pen.Color := FLightColor;
      Canvas.MoveTo(ARect.Left, ARect.Top);
      Canvas.LineTo(ARect.Right - 1, ARect.Top);

      Canvas.Pixels[ARect.Right - 1, ARect.Top] := FFixedColor;
    end // if  (Selected) then
    else
    begin

      // if (ACol < FixedCols) and (ACol <> FixedCols-1) then
      if ACol = 0 then
        Dec(ARect.Right, 1);

      if ACol < FixedCols - 1 then
        Inc(ARect.Left, 1);

      if not(goHorzLine in Options) and not(goFixedHorzLine in Options) then
      begin

        if ARow = RowCount - 1 then
          Dec(ARect.Bottom, 1);
        if ARow = 0 then
          Dec(ARect.Top, 1)
        else
          Inc(ARect.Top, 1);
      end;

      if ARow = 0 then
        Dec(ARect.Top, 1)
      else if (goFixedHorzLine in Options) then
      begin
        Canvas.Pen.Color := FShadowColor2;
        Canvas.MoveTo(ARect.Left, ARect.Top);
        Canvas.LineTo(ARect.Right, ARect.Top);
      end;

      if (ACol = FixedCols - 1) or (goFixedVertLine in Options) then // rtl
      begin
        Canvas.Pen.Color := FLightColor2;
        Canvas.MoveTo(ARect.Left, ARect.Bottom - 1);
        Canvas.LineTo(ARect.Left, ARect.Top + 1);
      end;

      if (ARow = 0) or (goFixedHorzLine in Options) then
      begin
        Canvas.Pen.Color := FLightColor2;
        Canvas.MoveTo(ARect.Left, ARect.Top + 1);
        Canvas.LineTo(ARect.Right, ARect.Top + 1);
      end;

      if (ACol = 0) or (goFixedVertLine in Options) then // rtl
      begin
        Canvas.Pen.Color := FShadowColor;
        Canvas.MoveTo(ARect.Right, ARect.Top + 1);
        Canvas.LineTo(ARect.Right, ARect.Bottom);
      end;

      if (ARow = RowCount - 1) or (goFixedHorzLine in Options) then
      begin
        Canvas.Pen.Color := FShadowColor;
        Canvas.MoveTo(ARect.Right, ARect.Bottom);
        Canvas.LineTo(ARect.Left - 1, ARect.Bottom);
      end;

      if not(goFixedHorzLine in Options) then // same as ltr
      begin
        if (goFixedVertLine in Options) then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left, ARect.Top - 1);
          Canvas.LineTo(ARect.Left, ARect.Bottom + 1);

          Canvas.Pen.Color := FShadowColor;
          Canvas.MoveTo(ARect.Right, ARect.Top - 1);
          Canvas.LineTo(ARect.Right, ARect.Bottom + 1);
        end;
      end;

      if not(goFixedVertLine in Options) then
      begin
        if (goFixedHorzLine in Options) then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left, ARect.Top + 1);
          Canvas.LineTo(ARect.Right + 1, ARect.Top + 1);

          Canvas.Pen.Color := FShadowColor;
          Canvas.MoveTo(ARect.Left, ARect.Top);
          Canvas.LineTo(ARect.Right + 1, ARect.Top);

          Canvas.MoveTo(ARect.Left, ARect.Bottom);
          Canvas.LineTo(ARect.Right + 1, ARect.Bottom);
        end;
      end;

      if (not(goFixedVertLine in Options)) and (not(goFixedHorzLine in Options))
      then
      begin
        if ACol = FixedCols - 1 then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left, ARect.Top);
          Canvas.LineTo(ARect.Left, ARect.Bottom);
        end;
        if ACol = 0 then
        begin
          Canvas.Pen.Color := FShadowColor;
          Canvas.MoveTo(ARect.Right, ARect.Top);
          Canvas.LineTo(ARect.Right, ARect.Bottom);
        end;
        if ARow = 0 then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left, ARect.Top + 1);
          Canvas.LineTo(ARect.Right + 1, ARect.Top + 1);
        end;
      end;

      if not Selected then
      begin
        if ACol = 0 then
        begin
          Canvas.Pen.Color := Parent.Brush.Color;
          Canvas.MoveTo(ARect.Right + 1, ARect.Top);
          Canvas.LineTo(ARect.Right + 1, ARect.Bottom);
        end;
      end;

    end; // if not (Selected) then

end;

procedure TAbjGrid.DrawHeaderRow(ACol, ARow: integer; Selected: boolean);
var
  txt: string;
  ARect, Rect, fRect: TRect;
  txtFlag: integer;
begin

  if UseRightToLeftAlignment then
    DrawHeaderRowRtoL(ACol, ARow, Selected)
  else
    DrawHeaderRowLtoR(ACol, ARow, Selected);

  ARect := CellRect(ACol, ARow);
  Dec(ARect.Top, 1);

  if (goFixedVertLine in Options) or (goVertLine in Options) then
    if ACol > 0 then
      Dec(ARect.Left, 1);

  Rect := ARect;

  txt := Cells[ACol, ARow];
  Inc(Rect.Top, 1);
  // Canvas.Font.Color := font.color;

  if ARow >= FixedRows then
    if FAutoRowNumber and (ACol = 0) then
    begin
      txt := IntToStr(ARow - FixedRows + 1);
    end;
  if txt <> '' then
  begin
    SetBkMode(Canvas.Handle, TRANSPARENT);

    Canvas.Font.Assign(FRowHeaderFont);
    if Selected then
      Canvas.Font.Color := GetShadeColor(NewColor(Canvas.Brush.Color, 30), 80)
    else
      Canvas.Font.Color := GetShadeColor(NewColor(Canvas.Brush.Color, 30), 30);

    fRect := Rect;
    InflateRect(fRect, -2, 0);
    Dec(fRect.Right, 2);
    txtFlag := DT_RIGHT; // DT_CENTER;
    Canvas.Font.Assign(FRowHeaderFont);
    DrawTextX(Canvas.Handle, txt, fRect, txtFlag);
  end;

  Canvas.Font.Assign(Font);
  Canvas.Brush.Color := Color;
  Canvas.Pen.Color := Color;

end;

procedure TAbjGrid.xDrawHeaderRow(ACol, ARow: integer; Selected: boolean);
var
  txt: string;
  Rect, fRect: TRect;
begin

  Rect := CellRect(ACol, ARow);

  txt := Cells[ACol, ARow];

  if FFlat and Selected then
    Canvas.Brush.Color := FFixedSelectColor
  else
    Canvas.Brush.Color := FixedColor;
  Inc(Rect.Bottom, 1);
  Canvas.FillRect(Rect);

  if UseRightToLeftAlignment then
  begin
    Dec(Rect.Left, 1);
    // Dec(Rect.Right, 1);
    Canvas.FillRect(Rect);
    Dec(Rect.Right, 1);
  end;
  Dec(Rect.Bottom, 1);

  if FFlat then
  begin
    if goFixedHorzLine in Options then
    begin
      Inc(Rect.Bottom, 1);
      if ARow <> 0 then
        Dec(Rect.Top, 1);
      if ACol <> 0 then
        Dec(Rect.Left, 1);
      Inc(Rect.Right, 1);
      Canvas.Pen.Color := FShadowColor;
      Canvas.Rectangle(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom);
      Dec(Rect.Bottom, 1);
      if ARow <> 0 then
        Inc(Rect.Top, 1);
      if ACol <> 0 then
        Inc(Rect.Left, 1);
      Dec(Rect.Right, 1);
    end;
  end
  else
  begin
    if Selected then
    begin
      Canvas.MoveTo(Rect.Right, Rect.Top);
      Canvas.Pen.Color := FShadowColor2;
      Canvas.LineTo(Rect.Right, Rect.Bottom);
      Canvas.LineTo(Rect.Left - 1, Rect.Bottom);

      Canvas.MoveTo(Rect.Right - 1, Rect.Top + 1);
      Canvas.Pen.Color := FShadowColor;
      Canvas.LineTo(Rect.Right - 1, Rect.Bottom - 1);
      Canvas.LineTo(Rect.Left + 1, Rect.Bottom - 1);

      Canvas.Pen.Color := FLightColor;

      Canvas.MoveTo(Rect.Left, Rect.Bottom - 1);
      Canvas.LineTo(Rect.Left, Rect.Top);
      Canvas.LineTo(Rect.Right - 1, Rect.Top);

    end
    else
    begin
      if ARow = 0 then
      begin
        Canvas.Pen.Color := Parent.Brush.Color;
        Canvas.MoveTo(Rect.Left, Rect.Top);
        Canvas.LineTo(Rect.Right + 1, Rect.Top);
        Inc(Rect.Top, 1);
      end;

      if ACol = 0 then
      begin
        Canvas.Pen.Color := Parent.Brush.Color;
        if UseRightToLeftAlignment then
        begin
          Canvas.MoveTo(Rect.Right, Rect.Top);
          Canvas.LineTo(Rect.Right, Rect.Bottom + 1);
          Dec(Rect.Right, 1);
        end
        else
        begin
          Canvas.MoveTo(Rect.Left, Rect.Top);
          Canvas.LineTo(Rect.Left, Rect.Bottom + 1);
          Inc(Rect.Left, 1);
        end;
      end;

      if goFixedHorzLine in Options then
      begin

        Canvas.Pen.Color := FShadowColor;

        Canvas.MoveTo(Rect.Right, Rect.Top);
        Canvas.LineTo(Rect.Right, Rect.Bottom);
        Canvas.LineTo(Rect.Left - 1, Rect.Bottom);

        Canvas.Pen.Color := FLightColor;

        Canvas.MoveTo(Rect.Left, Rect.Bottom - 1);
        Canvas.LineTo(Rect.Left, Rect.Top);
        Canvas.LineTo(Rect.Right, Rect.Top);
      end;
      Dec(Rect.Left, 1);
    end;
  end;

  Inc(Rect.Top, 1);
  // Canvas.Font.Color := font.color;

  if ARow >= FixedRows then
    if FAutoRowNumber and (ACol = 0) then
    begin
      txt := IntToStr(ARow - FixedRows + 1);
    end;
  if txt <> '' then
  begin
    SetBkMode(Canvas.Handle, TRANSPARENT);

    Canvas.Font.Assign(FRowHeaderFont);
    if Selected then
      Canvas.Font.Color := GetShadeColor(NewColor(Canvas.Brush.Color, 30), 80)
    else
      Canvas.Font.Color := GetShadeColor(NewColor(Canvas.Brush.Color, 30), 30);

    fRect := Rect;
    InflateRect(fRect, -2, 0);

    Inc(fRect.Top, 1);
    Inc(fRect.Left, 2);

    DrawTextX(Canvas.Handle, txt, fRect, DT_CENTER);

    Dec(fRect.Top, 1);
    Dec(fRect.Left, 2);
    Canvas.Font.Assign(FRowHeaderFont);

    DrawTextX(Canvas.Handle, txt, fRect, DT_CENTER);

  end;

  Canvas.Font.Assign(Font);
  Canvas.Brush.Color := Color;
  Canvas.Pen.Color := Color;

end;

{ TODO : To long.. to split }
procedure TAbjGrid.DrawHeaderCol(ACol, ARow: integer; Rect: TRect;
  Selected: boolean);
var
  fRect, ARect: TRect;
  txtFormat: Longint;
  txt: string;
  // txtShadow: integer;
begin


  // -----------

  if UseRightToLeftAlignment then
    DrawHeaderColRToL(ACol, ARow, Selected)
  else
    DrawHeaderColLToR(ACol, ARow, Selected);

  ARect := CellRect(ACol, ARow);
  Dec(ARect.Top, 1);

  if (goFixedVertLine in Options) or (goVertLine in Options) then
    if ACol > 0 then
      Dec(ARect.Left, 1);

  Rect := ARect;

  txt := Cells[ACol, ARow];
  Inc(Rect.Top, 1);


  // --------

  txt := Cells[ACol, ARow];
  txtFormat := DT_WORDBREAK + DT_CENTER;

  txtFormat := DrawTextBiDiModeFlags(txtFormat);

  if txt <> '' then
  begin
    Dec(Rect.Right, 16); // for sort icom
    SetBkMode(Canvas.Handle, TRANSPARENT);
    Inc(Rect.Top, 1);
    if not FFlat and Selected then
      Dec(Rect.Top, 1);

    Canvas.Font.Assign(FColHeaderFont);
    if Selected then
      Canvas.Font.Color := GetShadeColor(Canvas.Brush.Color, 80)
    else
    begin
      Canvas.Font.Color := GetShadeColor(Canvas.Brush.Color, 30);
    end;

    fRect := Rect;

    Inc(fRect.Top, 1);
    Inc(fRect.Left, 2);

    DrawTextX(Canvas.Handle, txt, fRect, txtFormat);

    // fRect := Rect;
    Dec(fRect.Top, 1);
    // Dec(fRect.Left,1);

    Canvas.Font.Assign(FColHeaderFont);
    DrawTextX(Canvas.Handle, txt, fRect, txtFormat);

  end;
  if FCurSortCol = ACol  then
    DrawSort(FCurSortCol, FSortDec, FFixedColor);
//  else
//    DrawSort(ACol, false, clBlack);

  Canvas.Font.Assign(Font);
  Canvas.Brush.Color := Color;
  Canvas.Pen.Color := Color;

end;
 { TODO : To lomg.. to split }
procedure TAbjGrid.DrawHeaderColLToR(ACol, ARow: integer; Selected: boolean);
var
  ARect: TRect;
begin

  ARect := CellRect(ACol, ARow);

  if FFlat and Selected then
    Canvas.Brush.Color := FSelectColor // FFixedSelectColor
  else
    Canvas.Brush.Color := FixedColor;
  InflateRect(ARect, 1, 1);
  Canvas.FillRect(ARect);
  InflateRect(ARect, -1, -1);

  if (goFixedVertLine in Options) or (goVertLine in Options) then
    if ACol > 0 then
      Dec(ARect.Left, 1);

  if (ARow = row + 1) and not(FFlat) then
    Inc(ARect.Top, 1); // do not paint over selected row above

  if FFlat then
    DrawFixedGridLine(ACol, ARow, Selected);
  if (not FFlat) then
  begin
    if (Selected) then
    begin
      Canvas.Pen.Color := FShadowColor2;
      Canvas.MoveTo(ARect.Right, ARect.Top);
      Canvas.LineTo(ARect.Right, ARect.Bottom);

      Canvas.Pen.Color := FShadowColor;
      Canvas.MoveTo(ARect.Right - 1, ARect.Top + 1);
      Canvas.LineTo(ARect.Right - 1, ARect.Bottom - 1);

      Canvas.Pen.Color := FLightColor;
      Canvas.MoveTo(ARect.Left + 1, ARect.Bottom - 1);
      Canvas.LineTo(ARect.Left + 1, ARect.Top);

      Canvas.Pen.Color := FShadowColor;
      Canvas.MoveTo(ARect.Left, ARect.Top - 1);
      Canvas.LineTo(ARect.Left, ARect.Bottom);

      if (goFixedHorzLine in Options) or (ARow = 0) then
      begin
        Canvas.Pen.Color := FLightColor;
        Canvas.MoveTo(ARect.Left + 1, ARect.Top);
        Canvas.LineTo(ARect.Right - 1, ARect.Top);
      end;

      if (ARow = FixedRows - 1) then
      begin
        Canvas.Pen.Color := FShadowColor2;
        Canvas.MoveTo(ARect.Right, ARect.Bottom);
        Canvas.LineTo(ARect.Left - 1, ARect.Bottom);
      end;
      if (goFixedHorzLine in Options) then
      begin
        if (ARow = FixedRows - 1) then
          Dec(ARect.Bottom, 1);
        Canvas.Pen.Color := FShadowColor;
        Canvas.MoveTo(ARect.Right - 1, ARect.Bottom);
        Canvas.LineTo(ARect.Left + 1, ARect.Bottom);

        Canvas.MoveTo(ARect.Right - 1, ARect.Top - 1);
        Canvas.LineTo(ARect.Left + 1, ARect.Top - 1);

      end;

      Canvas.Pixels[ARect.Right - 1, ARect.Top] := FFixedColor;
    end;

    if (not Selected) then
    begin

      if ARow = 0 then
      begin
        Inc(ARect.Top, 1);
      end;

      if (ACol = FixedCols) and (goFixedVertLine in Options) then
      begin
        Canvas.MoveTo(ARect.Left, ARect.Top);
        Canvas.LineTo(ARect.Left, ARect.Bottom);
      end;

      if (goFixedVertLine in Options) then
      begin
        Canvas.Pen.Color := FLightColor2;
        Canvas.MoveTo(ARect.Left + 1, ARect.Top);
        Canvas.LineTo(ARect.Left + 1, ARect.Bottom);

        Canvas.Pen.Color := FShadowColor;
        Canvas.MoveTo(ARect.Right, ARect.Top);
        Canvas.LineTo(ARect.Right, ARect.Bottom);

        Canvas.MoveTo(ARect.Left, ARect.Top);
        Canvas.LineTo(ARect.Left, ARect.Bottom);
      end;

      if (goFixedHorzLine in Options) then
      begin
        Canvas.Pen.Color := FLightColor2;
        Canvas.MoveTo(ARect.Left + 1, ARect.Top);
        Canvas.LineTo(ARect.Right, ARect.Top);

        Canvas.Pen.Color := FShadowColor;
        Canvas.MoveTo(ARect.Left, ARect.Bottom);
        Canvas.LineTo(ARect.Right, ARect.Bottom);

        Canvas.Pen.Color := FShadowColor;
        Canvas.MoveTo(ARect.Left - 1, ARect.Top - 1);
        Canvas.LineTo(ARect.Right, ARect.Top - 1);
      end;

      if not(goFixedHorzLine in Options) then
      begin
        if (goFixedVertLine in Options) then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left + 1, ARect.Top - 1);
          Canvas.LineTo(ARect.Left + 1, ARect.Bottom + 1);

          Canvas.Pen.Color := FShadowColor;
          Canvas.MoveTo(ARect.Left, ARect.Top - 1);
          Canvas.LineTo(ARect.Left, ARect.Bottom + 1);
        end;
      end;

      if not(goFixedVertLine in Options) then
      begin
        if (goFixedHorzLine in Options) then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left, ARect.Top);
          Canvas.LineTo(ARect.Right + 1, ARect.Top);

          Canvas.Pen.Color := FShadowColor;
          Canvas.MoveTo(ARect.Left, ARect.Bottom);
          Canvas.LineTo(ARect.Right + 1, ARect.Bottom);
        end;
      end;

      if (not(goFixedVertLine in Options)) then
      begin
        if ACol = ColCount - 1 then
        begin
          Canvas.Pen.Color := FShadowColor;
          Canvas.MoveTo(ARect.Right, ARect.Top);
          Canvas.LineTo(ARect.Right, ARect.Bottom);
        end;
      end;

      if not(goFixedHorzLine in Options) then
      begin
        if ARow = 0 then
        begin
          Canvas.Pen.Color := FLightColor2;
          Canvas.MoveTo(ARect.Left, ARect.Top);
          Canvas.LineTo(ARect.Right + 1, ARect.Top);
        end;

        if ARow = FixedRows - 1 then
        begin
          Canvas.Pen.Color := FShadowColor;
          Canvas.MoveTo(ARect.Left, ARect.Bottom);
          Canvas.LineTo(ARect.Right, ARect.Bottom);
        end;
      end;
    end;
  end;
end;

procedure TAbjGrid.DrawHeaderColRToL(ACol, ARow: integer; Selected: boolean);
var
  ARect: TRect;
begin
  ARect := CellRect(ACol, ARow);

  if (ARow = 0) and not(Selected) then
    Inc(ARect.Top, 1);

  if ARow = FixedRows - 1 then
    if not FFlat then
      Dec(ARect.Bottom, 1);

  if ACol = ColCount - 1 then
    if UseRightToLeftAlignment then
      Dec(ARect.Left, 1);


  if FFlat and Selected then
    Canvas.Brush.Color := FSelectColor // FFixedSelectColor
  else
    Canvas.Brush.Color := FixedColor;

  Canvas.FillRect(ARect);

  if FFlat then
    DrawFixedGridLine(ACol, ARow, Selected);

  if (not FFlat) then
  begin
    if (goFixedHorzLine in Options) then
    begin
      Canvas.Pen.Color := FLightColor2;
      Canvas.MoveTo(ARect.Left - 1, ARect.Top);
      Canvas.LineTo(ARect.Right + 1, ARect.Top);
    end
    else
    begin
      Canvas.Pen.Color := FixedColor;
      Canvas.MoveTo(ARect.Left + 1, ARect.Bottom);
      Canvas.LineTo(ARect.Right, ARect.Bottom);
    end;

    if (goFixedVertLine in Options) or Selected then
    begin

      Canvas.Pen.Color := FLightColor2;
      Canvas.MoveTo(ARect.Left, ARect.Top);
      Canvas.LineTo(ARect.Left, ARect.Bottom);

      Canvas.Pen.Color := FShadowColor;
      Canvas.MoveTo(ARect.Right, ARect.Top);
      Canvas.LineTo(ARect.Right, ARect.Bottom + 1);
    end
    else
    begin
      Canvas.Pen.Color := FixedColor;
      Canvas.MoveTo(ARect.Right, ARect.Top + 1);
      Canvas.LineTo(ARect.Right, ARect.Bottom);
    end;

    if Selected then
    begin
      Dec(ARect.Right, 1);
      Dec(ARect.Bottom, 1);
    end;

    if (goFixedHorzLine in Options) then
    begin
      Canvas.Pen.Color := FLightColor2;
      Canvas.MoveTo(ARect.Left, ARect.Top);
      Canvas.LineTo(ARect.Right, ARect.Top);

      Canvas.Pen.Color := FShadowColor;
      Canvas.MoveTo(ARect.Left, ARect.Bottom);
      Canvas.LineTo(ARect.Right + 1, ARect.Bottom);
    end;

    if (goFixedVertLine in Options) then
    begin
      Canvas.Pen.Color := FLightColor2;
      Canvas.MoveTo(ARect.Left, ARect.Top + 1);
      Canvas.LineTo(ARect.Left, ARect.Bottom);

      Canvas.Pen.Color := FShadowColor;
      Canvas.MoveTo(ARect.Right, ARect.Top);
      Canvas.LineTo(ARect.Right, ARect.Bottom);
    end;

    //
    if (ARow = 0) then
    begin
      Canvas.Pen.Color := FLightColor2;
      Canvas.MoveTo(ARect.Left, ARect.Top);
      Canvas.LineTo(ARect.Right + 1, ARect.Top);
    end;
    if (ARow = FixedRows - 1) then
    begin
      Canvas.Pen.Color := FShadowColor;
      Canvas.MoveTo(ARect.Left, ARect.Bottom);
      Canvas.LineTo(ARect.Right + 1, ARect.Bottom);
    end;
    if ACol = ColCount - 1 then
      if UseRightToLeftAlignment then
      begin
        Canvas.Pen.Color := FLightColor2;
        Canvas.MoveTo(ARect.Left, ARect.Top - 1);
        Canvas.LineTo(ARect.Left, ARect.Bottom);
      end;

    if not(goFixedVertLine in Options) and not(goFixedHorzLine in Options)
    then
    begin

      Canvas.Pixels[ARect.Left - 1, ARect.Top - 1] := FFixedColor;
      Canvas.Pixels[ARect.Left, ARect.Top - 1] := FFixedColor;
      Canvas.Pixels[ARect.Left - 1, ARect.Top] := FFixedColor;
    end;

    if Selected then
    begin
      Inc(ARect.Right, 1);
      Inc(ARect.Bottom, 1);
      Canvas.Pen.Color := FShadowColor2;
      Canvas.MoveTo(ARect.Left, ARect.Bottom);
      Canvas.LineTo(ARect.Right + 1, ARect.Bottom);

      Canvas.Pen.Color := FShadowColor2;
      Canvas.MoveTo(ARect.Right, ARect.Top);
      Canvas.LineTo(ARect.Right, ARect.Bottom);

    end;

    if not Selected then
    begin
      Dec(ARect.Top, 1);
      Canvas.Pen.Color := Parent.Brush.Color;
      Canvas.MoveTo(ARect.Left, 0);
      Canvas.LineTo(ARect.Right + 1, 0);
    end;
  end;
end;

procedure TAbjGrid.xDrawHeaderCol(ACol, ARow: integer; Rect: TRect;
  Selected: boolean);
var
  fRect: TRect;
  txtFormat: Longint;
  txt: string;
begin
  if FixedRows = 0 then
    exit;
  Rect := CellRect(ACol, ARow);
  fRect := Rect;

  txt := Cells[ACol, ARow];

  if (FFlat and Selected) or (FFlat and (goRowSelect in Options)) then
    Canvas.Brush.Color := NewColor(FFixedColor, 20)
    // GetShadeColor(FFixedColor, 20)//FFixedSelectColor
  else
    Canvas.Brush.Color := FixedColor;

  if FFlat then
  begin
    Inc(Rect.Bottom, 1);
    if ARow <> 0 then
      Dec(Rect.Top, 1);
    Dec(Rect.Left, 1);
    Inc(Rect.Right, 1);
    Canvas.Pen.Color := FShadowColor;
    Canvas.Rectangle(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom);
    Dec(Rect.Bottom, 1);
    if ARow <> 0 then
      Inc(Rect.Top, 1);
    Inc(Rect.Left, 1);
    Dec(Rect.Right, 1);

  end
  else
  begin
    if UseRightToLeftAlignment then
      if ACol = (ColCount - 1) then
      begin
        Dec(fRect.Left, 1);
      end;

    Canvas.FillRect(fRect);

    if (not Selected) and (BorderStyle <> bsNone) then
      Dec(fRect.Top, 1);
    if (not Selected) and (BorderStyle = bsNone) and (ARow = 0) then
      Inc(fRect.Top, 1);

    if Selected or (goRowSelect in Options) then
    begin

      Canvas.Pen.Color := FShadowColor2;
      Canvas.MoveTo(Rect.Right, Rect.Top);
      Canvas.LineTo(Rect.Right, Rect.Bottom);
      Canvas.LineTo(Rect.Left, Rect.Bottom);

      Canvas.Pen.Color := FShadowColor;
      Canvas.MoveTo(Rect.Right - 1, Rect.Top + 1);
      Canvas.LineTo(Rect.Right - 1, Rect.Bottom - 1);
      Canvas.LineTo(Rect.Left + 1, Rect.Bottom - 1);

      Canvas.Pen.Color := FLightColor; // clWhite;
      Canvas.MoveTo(fRect.Left, fRect.Bottom - 1);
      Canvas.LineTo(fRect.Left, fRect.Top);
      Canvas.LineTo(fRect.Right, fRect.Top);

    end
    else
    begin
      if ARow = 0 then
      begin
        Canvas.Pen.Color := Parent.Brush.Color;
        Canvas.MoveTo(fRect.Left, fRect.Top - 1);
        Canvas.LineTo(fRect.Right + 1, fRect.Top - 1);
      end;

      Canvas.Pen.Color := FShadowColor; // FFixedClellBorderColor;
      Canvas.MoveTo(fRect.Right, fRect.Top);
      Canvas.LineTo(fRect.Right, fRect.Bottom);
      Canvas.LineTo(fRect.Left - 1, fRect.Bottom);

      Canvas.Pen.Color := FLightColor; // clWhite;

      Canvas.MoveTo(fRect.Left, fRect.Bottom - 1);
      Canvas.LineTo(fRect.Left, fRect.Top);
      Canvas.LineTo(fRect.Right, fRect.Top);

    end;
  end;

  txtFormat := 0;
  if (Rect.Right - Rect.Left) > Canvas.TextWidth(txt) then
    txtFormat := DT_CENTER;

  txtFormat := DrawTextBiDiModeFlags(txtFormat);

  if txt <> '' then
  begin
    SetBkMode(Canvas.Handle, TRANSPARENT);
    Inc(Rect.Top, 1);
    if not FFlat and Selected then
      Dec(Rect.Top, 1);

    Canvas.Font.Assign(FColHeaderFont);
    if Selected then
      Canvas.Font.Color := GetShadeColor(Canvas.Brush.Color, 70)
    else
    begin
      Canvas.Font.Color := GetShadeColor(Canvas.Brush.Color, 30);
    end;

    fRect := Rect;
    Inc(fRect.Top, 1);
    Inc(fRect.Left, 2);

    DrawTextX(Canvas.Handle, txt, fRect, txtFormat);

    Dec(fRect.Top, 1);
    Dec(fRect.Left, 2);

    Canvas.Font.Assign(FColHeaderFont);
    DrawTextX(Canvas.Handle, txt, fRect, txtFormat);

  end;
  if FCurSortCol > -1 then
    DrawSort(FCurSortCol, FSortDec, FFixedColor);
  Canvas.Font.Assign(Font);
  Canvas.Brush.Color := Color;
  Canvas.Pen.Color := Color;

end;

function TAbjGrid.GetCellText(ACol, ARow: integer): string;
begin
  result := Cells[ACol, ARow]
end;

function TAbjGrid.GetCellHeight(ACol, ARow: integer): integer;
var
  H: integer;
  S: string;
  R: TRect;
  FontStyle: TFontStyles;
begin
  if GlobalCellMargin > 200 then // in case we did not init the value
    GlobalCellMargin := 0;
  result := DefaultRowHeight;
  S := GetCellText(ACol, ARow);

  R.Top := 0;
  R.Left := 0;
  R.Right := ColWidths[ACol] - GlobalCellMargin;
  R.Bottom := 100;
  FontStyle := Canvas.Font.Style;
  Canvas.Font.Style := [fsBold];
  H := DrawTextX(Canvas.Handle, S, R, DT_WORDBREAK + DT_CALCRECT);
  Canvas.Font.Style := FontStyle;

  if H > result then
    result := H + Round(Canvas.Font.Size * 1.5); // 14;

end;

function TAbjGrid.DetermineCellHeight(S: string; ACol, ARow: integer): integer;
var
  R: TRect;
  FontStyle: TFontStyles;
begin
  result := RowHeights[ARow];

  R.Top := 0;
  R.Left := 0;
  R.Right := ColWidths[ACol] - 20;
  R.Bottom := result;
  result := DrawTextX(Canvas.Handle, S, R, DT_WORDBREAK + DT_CALCRECT);
  result := result + (Canvas.TextHeight('|') div 2);
end;

procedure TAbjGrid.DrawCellFocus(ACol, ARow: integer);
var
  Rect: TRect;
  clr: TColor;
begin
  clr := SelectColor; // FixedColor;  ////clHighlight;
  Rect := CellRect(ACol, ARow);

  if UseRightToLeftAlignment then
  begin
    Dec(Rect.Left, 1);
    Dec(Rect.Right, 1);
  end;

  if (goRowSelect in Options) then
  begin
    Canvas.Brush.Color := FSelectColor;
    Canvas.FillRect(Rect);
    exit;
  end;
  Canvas.Pen.Width := 1;
  Canvas.Pen.Color := GetShadeColor(clr, 60); // Clr;
  Canvas.MoveTo(Rect.Left, Rect.Top + 1);
  Canvas.LineTo(Rect.Right - 1, Rect.Top + 1);
  Canvas.LineTo(Rect.Right - 1, Rect.Bottom - 1);
  Canvas.LineTo(Rect.Left, Rect.Bottom - 1);
  Canvas.LineTo(Rect.Left, Rect.Top + 0);

  Canvas.Pen.Color := GetShadeColor(clr, 30); // NewColor(Clr, 30);
  Canvas.MoveTo(Rect.Left + 1, Rect.Top + 2);
  Canvas.LineTo(Rect.Right - 2, Rect.Top + 2);
  Canvas.LineTo(Rect.Right - 2, Rect.Bottom - 2);
  Canvas.LineTo(Rect.Left + 1, Rect.Bottom - 2);
  Canvas.LineTo(Rect.Left + 1, Rect.Top + 2);

  Canvas.Pen.Width := 1;

  Canvas.Pen.Width := 1;
  Canvas.Pen.Color := FBackColor;
end;

procedure TAbjGrid.DrawCellText(Rect: TRect; const S: string;
  const Format: NativeInt);
begin
  Inc(Rect.Left, 1);
  Inc(Rect.Top, 3);
  Dec(Rect.Right, 3);
  Dec(Rect.Bottom, 2);

  Canvas.FillRect(Rect);
  DrawTextW(Canvas.Handle, PWideChar(S), Length(S), Rect, Format)

end;

procedure TAbjGrid.DoEnter;
var
  i: integer;
begin
  inherited;
  exit;
  // if not ShowColHeader then exit;
  DrawHeaderRow(col, row, True);
  for i := 0 to FixedRows - 1 do
    DrawHeaderCol(col, i, CellRect(col, i), True);

end;

procedure TAbjGrid.SelectAll;
begin
  Selection := TGridRect(Rect(1, 1, ColCount - 1, RowCount - 1));
end;

function TAbjGrid.SelectCell(ACol, ARow: integer): boolean;
var
  hRect: TRect;
  i: integer;
begin

  // inherited SelectCell(ACol, ARow);
  if (col <> ACol) then
  begin
    for i := 0 to FixedRows - 1 do
      DrawHeaderCol(col, i, CellRect(col, i), false);
    if UseRightToLeftAlignment then
      ChangeGridOrientation(True);

    for i := 0 to FixedRows - 1 do
      inherited DrawCell(col, i, CellRect(col, i), []);

    if UseRightToLeftAlignment then
      ChangeGridOrientation(false);

    for i := 0 to FixedRows - 1 do
      DrawHeaderCol(ACol, i, CellRect(ACol, i), True);

    if UseRightToLeftAlignment then
      ChangeGridOrientation(True);
    inherited DrawCell(ACol, 0, CellRect(ACol, 0), []);
    if UseRightToLeftAlignment then
      ChangeGridOrientation(false);
  end;

  // -----
  if FixedCols > 0 then
  begin
    if FFlat then
    begin
      Dec(hRect.Top);
      Inc(hRect.Bottom);
    end;
    for i := 0 to FixedCols - 1 do
      DrawHeaderRow(i, row, false);

    if UseRightToLeftAlignment then
      ChangeGridOrientation(True);
    for i := 0 to FixedCols - 1 do
    begin
      hRect := CellRect(i, row);
      inherited DrawCell(i, row, hRect, []);
    end;
    if UseRightToLeftAlignment then
      ChangeGridOrientation(false);

    // -----

    hRect := CellRect(0, ARow);
    if FFlat then
    begin
      Dec(hRect.Top);
      Inc(hRect.Bottom);
    end;

    begin
      for i := 0 to FixedCols - 1 do
        DrawHeaderRow(i, ARow, True);
      if UseRightToLeftAlignment then
        ChangeGridOrientation(True);
      for i := 0 to FixedCols - 1 do
      begin
        hRect := CellRect(i, ARow);
        inherited DrawCell(i, ARow, hRect, []);
      end;
      if UseRightToLeftAlignment then
        ChangeGridOrientation(false);
    end;
  end;
  result := True;
  Editor.DoHide;
  if assigned(OnSelectCell) then
    OnSelectCell(self, ACol, ARow, result);

end;

function TAbjGrid.Selected(ACol, ARow: integer): boolean;
begin
  result := (ACol >= Selection.Left) and (ACol <= Selection.Right) and
    (ARow >= Selection.Top) and (ARow <= Selection.Bottom);
end;

function TAbjGrid.SelectionRect: TRect;
begin
  Rect(CellRect(Selection.Left, Selection.Top).Left,
    CellRect(Selection.Left, Selection.Top).Top, CellRect(Selection.Right,
    Selection.Bottom).Right, CellRect(Selection.Right,
    Selection.Bottom).Bottom);

end;

procedure TAbjGrid.DrawCell(ACol, ARow: integer; Rect: TRect;
  State: TGridDrawState);
var
  Org, Ext: TPoint;
  X: integer;
  OldRect: TRect;
begin
  if ARow = -1 then
    ARow := 0;
  if UseRightToLeftAlignment then
  begin
    OldRect := Rect;
    Rect.Left := ClientWidth - Rect.Left;
    Rect.Right := ClientWidth - Rect.Right;
    X := Rect.Left;
    Rect.Left := Rect.Right;
    Rect.Right := X;

    Org := Point(0, 0);
    Ext := Point(1, 1);
    SetMapMode(Canvas.Handle, mm_Anisotropic);
    SetWindowOrgEx(Canvas.Handle, Org.X, Org.Y, nil);
    SetViewportExtEx(Canvas.Handle, ClientWidth, ClientHeight, nil);
    SetWindowExtEx(Canvas.Handle, Ext.X * ClientWidth,
      Ext.Y * ClientHeight, nil);
    GridDrawCell(ACol, ARow, Rect, State);

    inherited DrawCell(ACol, ARow, OldRect, State);
    Org := Point(ClientWidth, 0);
    Ext := Point(-1, 1);
    SetMapMode(Canvas.Handle, mm_Anisotropic);
    SetWindowOrgEx(Canvas.Handle, Org.X, Org.Y, nil);
    SetViewportExtEx(Canvas.Handle, ClientWidth, ClientHeight, nil);
    SetWindowExtEx(Canvas.Handle, Ext.X * ClientWidth,
      Ext.Y * ClientHeight, nil);
  end
  else
  begin
    GridDrawCell(ACol, ARow, Rect, State);

    inherited DrawCell(ACol, ARow, Rect, State);

  end;

end;

function TAbjGrid.DrawGridLine(ACol, ARow: integer; Selected: boolean): TRect;
var
  ARect: TRect;
  Color: TColor;
begin
  ARect := CellRect(ACol, ARow);
  InflateRect(ARect, 1, 1);
  if ACol = ColCount - 1 then
    Dec(ARect.Right, 1);
  if ARow = RowCount - 1 then
    Dec(ARect.Bottom, 1);


  Color := FBackColor;
  if (goHorzLine in Options) then
    Color := FGridLineColor;
  if (goVertLine in Options) then
    Color := FGridLineColor;

  Canvas.Pen.Color := Color;
  Canvas.MoveTo(ARect.Left, ARect.Top);
  Canvas.LineTo(ARect.Right, ARect.Top);
  Canvas.MoveTo(ARect.Left, ARect.Bottom);
  Canvas.LineTo(ARect.Right, ARect.Bottom);

  if not(goHorzLine in Options) then
    Color := FBackColor;
  if not(goVertLine in Options) then
    Color := FBackColor;
  Canvas.Pen.Color := Color;
  Canvas.MoveTo(ARect.Left, ARect.Top);
  Canvas.LineTo(ARect.Left, ARect.Bottom);
  Canvas.MoveTo(ARect.Right, ARect.Top);
  Canvas.LineTo(ARect.Right, ARect.Bottom);

  result := ARect;
end;

function TAbjGrid.DrawFixedGridLine(ACol, ARow: integer;
  Selected: boolean): TRect;
var
  ARect: TRect;

begin

  ARect := CellRect(ACol, ARow);

  Dec(ARect.Top, 1);
  if (ARow = 0) then
    if not FFlat then
      Inc(ARect.Top, 2)
    else
      Inc(ARect.Top, 1);
  if ACol < FixedCols then
  begin
    if not UseRightToLeftAlignment then
    begin
      if ACol = 0 then
        if (not Selected) and (not FFlat) then
          Inc(ARect.Left, 1);

      if (goFixedVertLine in Options) or (goVertLine in Options) then
        if ACol > 0 then
          Dec(ARect.Left, 1);

    end
    else
    begin
      Dec(ARect.Left, 1);
      Dec(ARect.Right, 1);

      if (goFixedVertLine in Options) or (goVertLine in Options) then
        if ACol > 0 then
          Inc(ARect.Right, 1);

    end;
  end;

  if ACol >= FixedCols then
  begin
    Dec(ARect.Left, 1);
    if (not FFlat) and (not Selected) then
      Inc(ARect.Top, 1);
  end;

  if FFlat then
    Canvas.Pen.Color := FGridLineColor
  else
    Canvas.Pen.Color := Canvas.Brush.Color;

  if goFixedHorzLine in Options then
  begin
    Canvas.Pen.Color := FGridLineColor;

    Canvas.MoveTo(ARect.Left, ARect.Bottom);
    Canvas.LineTo(ARect.Right, ARect.Bottom);

    if (ARow = row + 1) and not(FFlat) then
    else
    begin
      Canvas.MoveTo(ARect.Left, ARect.Top);
      Canvas.LineTo(ARect.Right, ARect.Top);
    end;
  end
  else
  begin
    Canvas.Pen.Color := Canvas.Brush.Color;
    Canvas.MoveTo(ARect.Left + 1, ARect.Bottom);
    Canvas.LineTo(ARect.Right, ARect.Bottom);
  end;

  if goFixedVertLine in Options then
  begin
    Canvas.Pen.Color := FGridLineColor;
    Canvas.MoveTo(ARect.Right, ARect.Top);
    Canvas.LineTo(ARect.Right, ARect.Bottom);
    Canvas.MoveTo(ARect.Left, ARect.Top);
    Canvas.LineTo(ARect.Left, ARect.Bottom);
  end
  else
  begin
    Canvas.Pen.Color := Canvas.Brush.Color;
    Canvas.MoveTo(ARect.Right, ARect.Top);
    Canvas.LineTo(ARect.Right, ARect.Bottom);
    Canvas.MoveTo(ARect.Left, ARect.Top);
    Canvas.LineTo(ARect.Left, ARect.Bottom);
  end;

  result := ARect;

end;

procedure TAbjGrid.SetFlat(const Value: boolean);
begin
  if FFlat <> Value then
  begin
    FFlat := Value;
    Invalidate_AbjGrid;

  end;
end;

function TAbjGrid.IsColInSelect(ACol: integer): boolean;
begin
  result := (ACol >= Selection.Left) and (ACol <= Selection.Right);
end;

procedure TAbjGrid.ColumnMoved(FromIndex, ToIndex: integer);
var
  C: integer;
  tmp: integer;
begin
  inherited;

  if FromIndex > ToIndex then
  begin
    tmp := FColState[FromIndex];
    for C := FromIndex downto ToIndex + 1 do
      FColState[C] := FColState[C - 1];
    FColState[ToIndex] := tmp;
  end;

  if FromIndex < ToIndex then
  begin
    tmp := FColState[FromIndex];
    for C := FromIndex to ToIndex - 1 do
      FColState[C] := FColState[C + 1];
    FColState[ToIndex] := tmp;
  end;
  // FColMoving := true;
  if assigned(OnColumnMoved) then
    OnColumnMoved(self, FromIndex, ToIndex);
end;

procedure TAbjGrid.LoadFromStringList(var L: TStringList; DrawHeader: boolean;
  Index: integer);
var
  S, S2: string;
  i, iCol, iRow, cRow, sCount, lCount: integer;
  Delimiter: string;

begin
  Screen.Cursor := crHourGlass;
  ClearAll;

  lCount := L.Count;
  Delimiter := GetFieldDelimiter;
  if DrawHeader then
    cRow := 0
  else
    cRow := FixedRows;

  for iRow := Index to lCount - 1 do
  begin
    iCol := FixedCols;
    S := L.Strings[iRow];
    sCount := Length(S);
    for i := 1 to sCount do
    begin
      if (S[i] <> Delimiter) and (S[i] <> FCommentDelimiter) then
        S2 := S2 + S[i]
      else
      begin
        if S[i] = FCommentDelimiter then
          break;
        begin
          Cells[iCol, cRow] := S2;
          Inc(iCol);
        end;
        S2 := ''
      end;
    end;
    Cells[iCol, cRow] := S2;
    S2 := '';
    if (iCol + 1) > ColCount then
      ColCount := iCol + 1;
    if S <> '' then
      if S[1] <> FCommentDelimiter then
        Inc(cRow);


    Application.ProcessMessages;
  end;
  if cRow > RowCount then  // ? is there memory issue for large files
    RowCount := cRow;

  if L.Count = 1 then
  begin
    S := Trim(L.Strings[0]);
    iCol := 0;
    if S[1] <> FCommentDelimiter then
      for i := 1 to Length(S) do
        if S[i] = Delimiter then
          Inc(iCol);
    ColCount := FixedCols + iCol;
  end;

  InitGrid; // New

  Screen.Cursor := crDefault;

end;



procedure TAbjGrid.LoadFromFile(const FileName: string; HasHeader: boolean);
var
  S, S2: string;
  i, col, iRow, cRow, sCount, lCount: integer;
  L: TStringList;
  Delimiter: string;
  isCSV: boolean;

begin

  Screen.Cursor := crHourGlass;
  isCSV := (UpperCase(ExtractFileExt(FileName)) = UpperCase('.CSV'));
  Clear;
  ColCount := 2;

  L := TStringList.Create;
  try
{$IFDEF UNICODE}
    L.LoadFromFile(FileName, TEncoding.UTF8);

    if L.Count < 2 then
    begin
      L.LoadFromFile(FileName, TEncoding.BigEndianUnicode);
    end;

    if L.Count < 2 then
    begin
      L.LoadFromFile(FileName, TEncoding.Unicode);
    end;

    if L.Count < 2 then
    begin
      L.LoadFromFile(FileName);
    end;

{$ELSE}
    L.LoadFromFile(FileName);
{$ENDIF}
    lCount := L.Count;

    Delimiter := GetFieldDelimiter;
    if isCSV then
      Delimiter := #44;

    if HasHeader then
      cRow := FixedRows - 1
    else
      cRow := FixedRows;

    for iRow := 0 to lCount - 1 do
    begin
      col := FixedCols;
      S := L.Strings[iRow];
      sCount := Length(S);
      for i := 1 to sCount do
      begin
        if (S[i] <> Delimiter) { and (s[i] <> FCommentDelimiter) } then
          S2 := S2 + S[i]
        else
        begin
          if S[i] = FCommentDelimiter then
            break;
          begin
            if isCSV then
              S2 := AnsiDequotedStr(S2, '"');
            // s2 := Remove_Quotation(s2);
            Cells[col, cRow] := S2;
            Inc(col);
          end;
          S2 := ''
        end;
      end;
      Cells[col, cRow] := S2;
      S2 := '';
      if (col + 1) > ColCount then
        ColCount := col + 1;
      if S <> '' then
        if S[1] <> FCommentDelimiter then
          Inc(cRow);

      Application.ProcessMessages;
    end;
    if cRow > RowCount then
      RowCount := cRow;
    InitGrid;
  finally
    L.Free;
    Screen.Cursor := crDefault;
  end;

end;

procedure StringSave(const FileName: TFileName; const Data: UTF8String);
// Fredrik Nordbakke : https://forums.embarcadero.com/message.jspa?messageID=217785
var
  FS: TFileStream;
begin
  FS := TFileStream.Create(FileName, fmCreate);
  try
    FS.Write(Pointer(Data)^, Length(Data));
  finally
    FS.Free;
  end;
end;

procedure TAbjGrid.SaveToFile(const FileName: string; IncludeHeader: boolean);
var
  sRow, S: WideString;
  S2: string;
  ACol, ARow, RowStart: integer;
  L: TStringList;
  Delimiter: string;
  Cancel: boolean;
begin
  if Trim(FileName) = '' then
    exit;

  Cancel := false;
  L := TStringList.Create;
  try

    Delimiter := GetFieldDelimiter;

    if IncludeHeader then
      RowStart := 0
    else
      RowStart := FixedRows;

    for ARow := RowStart to RowCount - 1 do
    begin
      sRow := '';

      for ACol := (FixedCols) to ColCount - 1 do
      begin
        S := Cells[ACol, ARow];
        S2 := S;
        if assigned(OnExport) then
          OnExport(self, etText, ACol, ARow, S2, Cancel);
        if Cancel then
        begin
          break;
        end;

        sRow := sRow + S + Delimiter;
      end;
      sRow := Copy(sRow, 1, Length(sRow)-1);
      L.Add(sRow);
      sRow := '';
    end;

{$IFDEF UNICODE}

    StringSave(FileName, UTF8Encode(Copy(L.Text, 1,
    Length(L.Text) - 2 { Length(LineBreak) } )));
{$ELSE}
    L.SaveToFile(FileName);
{$ENDIF}
  finally
    L.Free;
  end;

end;

procedure TAbjGrid.SaveToFile(const FileName: string; SaveOption: TSaveOption);
  function DoText(S: string): string;
  var
    i, X: integer;
    quoted: boolean;
  begin
    X := 0;

    quoted := false;
    for i := 1 to Length(S) do
    begin
      X := X + 1;
      case S[X] of
        #9:
          begin
            delete(S, X, 1);
            X := X - 1;
            quoted := True;
          end;
        #10:
          begin
            quoted := True;
          end;
        #13:
          begin
            S[X] := #10;
            quoted := True;
          end;

        ',', ';':
          quoted := True;
        '"':
          begin
            insert('"', S, X);
            X := X + 1;
            quoted := True;
          end;
      end;
    end;
    if quoted then
      S := '"' + S + '"';
    result := S; // trim(s);
  end;

var
  sRow, S: string;
  ACol, ARow: integer;
  L: TStringList;
  Delimiter: string;
  Cancel: boolean;
  RowStart: integer;
begin
  Cancel := false;
  L := TStringList.Create;
  try

    Delimiter := GetFieldDelimiter;
    if soIncludeHeaders in [SaveOption] then
      RowStart := 0
    else
      RowStart := FixedRows;
    for ARow := RowStart to RowCount - 1 do
    begin
      sRow := '';

      for ACol := (FixedCols) to ColCount - 1 do
      begin
        S := Cells[ACol, ARow];

        if assigned(OnExport) then
          OnExport(self, etText, ACol, ARow, S, Cancel);
        if Cancel then
        begin
          break;
        end;

        sRow := sRow + DoText(S) + Delimiter;
      end;
      L.Add(sRow);
      sRow := '';
    end;
{$IFDEF UNICODE}
    L.SaveToFile(FileName, TEncoding.UTF8);
{$ELSE}
    L.SaveToFile(FileName);
{$ENDIF}
  finally
    L.Free;
  end;
end;

procedure TAbjGrid.SaveToFile(const FileName: string; SaveOption: TSaveOption;
  Target: TExportTarget);
var
  sRow, S: string;
  ACol, ARow: integer;
  L: TStringList;
  Delimiter: string;
  Cancel: boolean;
  RowStart: integer;
  CurRow: integer;

  function DoText(S: string): string;
  var
    i, X: integer;
    quoted: boolean;
  begin
    X := 0;
    // s := trim(s);
    quoted := false;
    for i := 1 to Length(S) do
    begin
      X := X + 1;
      case S[X] of

        #9:
          begin
            delete(S, X, 1);
            X := X - 1;
            quoted := True;
          end;
        #10:
          begin
            quoted := True;
          end;
        #13:
          begin
            S[X] := #10;
            quoted := True;
          end;

        ',', ';':
          quoted := True;
        '"':
          begin
            insert('"', S, X);
            X := X + 1;
            quoted := True;
          end;
      end;
    end;
    if quoted then
      S := '"' + S + '"';
    result := S; // trim(s);
  end;

begin
  if Target = etCSV then
    FieldDelimiter := 'COMMA';

  Cancel := false;
  CurRow := row;
  L := TStringList.Create;
  try
    Delimiter := GetFieldDelimiter;
    if soIncludeHeaders in [SaveOption] then
      RowStart := 0
    else
      RowStart := FixedRows;

    for ARow := RowStart to RowCount - 1 do
    begin
      sRow := '';

      for ACol := (FixedCols) to ColCount - 1 do
      begin
        S := Cells[ACol, ARow];

        if assigned(OnExport) then
          OnExport(self, etText, ACol, ARow, S, Cancel);
        if Cancel then
        begin
          break;
        end;
        if Target = etCSV then
          S := DoText(S);
        sRow := sRow + S + Delimiter;
      end;
      // Delete Last Delimiter
      sRow := Copy(sRow, 1, Length(sRow) - 1);
      L.Add(sRow);
      sRow := '';
    end;

{$IFDEF UNICODE}
    L.SaveToFile(FileName, TEncoding.UTF8);
    // StringSave(Filename, UTF8Encode(Copy(L.Text, 1, Length(L.Text)-2{Length(LineBreak)})));
{$ELSE}
    L.SaveToFile(FileName);
{$ENDIF}
  finally
    L.Free;
    row := CurRow;
  end;

end;

// {$IFDEF VER5U}
procedure TAbjGrid.SaveToExcel(const FileName, Title, Footer: string;
  ShowApp: boolean);
var

  xlApp: variant;
  xlBook: variant;
  xlSheet: variant;
  ACol, ARow: integer;
  txt: String;
  ColWidth: integer;
  Cancel: boolean;

const
  xlCenter = $FFFFEFF4;
  CLASS_ExcelApplication: TGUID = '{00024500-0000-0000-C000-000000000046}';
begin

  Cancel := false;
  try
    xlApp := CreateOleObject('Excel.Application');


    xlApp.Visible := ShowApp;
    xlBook := xlApp.Workbooks.Add;
    xlSheet := xlBook.Worksheets[1];
    // xlApp.DisplayAlerts := false;
    xlSheet.DisplayRightToLeft := (BiDiMode = bdRightToLeft);

    xlSheet.Cells.Font.Size := 10;

    // --------Title
    txt := Title;
    xlSheet.Range['A1:A1'].Cells[1, 1].Value := txt;
    xlSheet.Range['A1:A1'].Font.Size := 10;
    xlSheet.Range['A1:A1'].Font.Bold := True;
    xlSheet.Range['A1:A1'].Font.Color := 16711680; // ;clBlue;
    xlSheet.Range['A1:A1'].Columns[1].ColumnWidth := 6;
    // ------ Col Title
    xlSheet.Range['B3:Z3'].RowHeight := 30;
    xlSheet.Range['B3:Z3'].WrapText := True;
    try
      // xlSheet.Range['B3:Z3'].HorizontalAlignment := xlCenter;
      // xlSheet.Range['B3:Z3'].VerticalAlignment := xlCenter;

    except

    end;
    xlSheet.PageSetup.PrintTitleRows := '$3:$3';
    // .PrintTitleColumns = ""

    // xlSheet.Range['A3:Z3'].AutoFit;
    xlSheet.Range['A3:Z3'].Font.Size := 10;
    xlSheet.Range['A3:Z3'].Font.Bold := false;
    xlSheet.Range['A3:Z3'].Font.Color := 8388608; // clNavy;
    xlSheet.Range['A4:Z4'].Font.Size := 10;

    // .Borders.LineStyle = xlThick
    for ACol := 0 to ColCount - 1 do
    begin
      ColWidth := ColWidths[ACol] div 7;
      if ColWidth < 0 then
        ColWidth := 0;
      if ColWidth > 50 then
        ColWidth := 50; // ColWidth div 3;
      xlSheet.Range['A1:Z1'].Columns[ACol + 1].ColumnWidth := ColWidth;
    end;
    // ----Data-------------

    for ARow := 0 to RowCount - 1 do
    begin
      try
        Application.ProcessMessages;
        for ACol := 1 to ColCount - 1 do
        begin
          txt := Cells[ACol, ARow];

          if assigned(OnExport) then
            OnExport(self, etExcel, ACol, ARow, txt, Cancel);
          if Cancel then
          begin
            break;
          end;

          if IsDate(txt) then
          begin
            if BiDiMode = bdRightToLeft then
            begin
              txt := MonthDayYearFormat(txt);
              xlSheet.Range['A4:Z4'].Cells[ARow, ACol + 1].NumberFormat :=
                {$IF CompilerVersion >= 22.0}
                'yyyy' + FormatSettings.DateSeparator + 'mm' + FormatSettings.DateSeparator + 'dd';
                {$ELSE}
                'yyyy' + DateSeparator + 'mm' + DateSeparator + 'dd';
                {$ENDIF}
            end;
          end;
          if Trim(txt) = '' then
            txt := ' ';

          xlSheet.Range['A4:Z4'].Cells[ARow, ACol + 1].Value := txt;
        end;
      except
      end;
    end;
    ARow := RowCount - 1;
    // ----------------
    // xlSheet.Range[Cells[4, 1], Cells[4 + rs.RecordCount,2+ rs.Fields.Count]).Font.Size := 12;
    // xlSheet.Range[Cells[1, 1], Cells[5, 3]].Font.Italic = True;

    // txt := #39 + DateToStr(Date);
    // xlSheet.Range['A4:Z4'].Cells[Row + 2 , 2].Value := txt;
    txt := Footer;
    xlSheet.Range['A4:Z4'].Cells[ARow + 3, 2].Value := txt;

    xlSheet.Cells[ARow + 5, 1].Font.Size := 10;
    xlSheet.Cells[ARow + 6, 1].Font.Size := 10;
    xlSheet.Cells[ARow + 5, 1].Font.Italic := True;
    xlSheet.Cells[ARow + 6, 1].Font.Italic := True;
    xlSheet.Cells[ARow + 5, 1].Font.Bold := True;
    xlSheet.Cells[ARow + 6, 1].Font.Bold := True;

    // ---------
    if Trim(FileName) <> '' then
      xlSheet.SaveAs(FileName);

    if not ShowApp then
      xlApp.Quit;
  except
    if not ShowApp then
      xlApp.Quit;
    Screen.Cursor := 0;
  end;
  Screen.Cursor := 0;

end;
// {$ENDIF}

procedure TAbjGrid.SaveToHTML(const FileName, Title, Footer: string);
var
  S, sRow: string;
  col, row: integer;
  L: TStringList;
  Cancel: boolean;
begin
  Cancel := false;
  L := TStringList.Create;
  try

    L.Add('<html>');
    L.Add('<head>');
    L.Add('<title>' + Title + '</title>');
    L.Add('</head>');
    L.Add('');
    L.Add('<style>');;

    L.Add('  table {background-color:silver; border-width:with=1; font-size : x-small;}');
    L.Add('  th {background-color: Lavender; vertical-align:top;}');
    L.Add('  td {background-color: GhostWhite; vertical-align:Top;}');

    S := '';
    if BiDiMode = bdRightToLeft then
      S := 'Direction = "rtl"; ';
    L.Add('  body {' + S +
      'font-family: Verdana, Tahoma, Arial, Helvetica, Geneva, Sans-Serif; font-size : x-small;}');

    L.Add('</style>');
    L.Add('');
    L.Add('<body>');
    L.Add('');
    if Title <> '' then
      L.Add('<h4>' + Title + '</h4>');
    L.Add('');
    L.Add('<table cellspacing = "1" cellpadding = "2">');
    L.Add('<CAPTION>' + Title + '</CAPTION>');

    sRow := '';
    for col := (FixedCols) to ColCount - 1 do
    begin
      S := Cells[col, 0];
      if Trim(S) = '' then
        S := '&nbsp;';
      sRow := sRow + '<th nowrap>' + S + '</th>';
    end;
    L.Add('<tr>' + sRow + '</tr>');

    for row := 1 to RowCount - 1 do
    begin

      sRow := '';
      for col := (FixedCols) to ColCount - 1 do
      begin
        S := Cells[col, row];
        if Trim(S) = '' then
          S := '&nbsp;';

        if assigned(OnExport) then
          OnExport(self, etHTML, col, row, S, Cancel);
        if Cancel then
        begin
          break;
        end;

        sRow := sRow + '<td nowrap>' + S + '</td>';
      end;
      L.Add('<tr>' + sRow + '</tr>');

      S := '';

      if Cancel then
        break;
    end;
    L.Add('</table>');
    L.Add('');
    if Footer <> '' then
      L.Add('<p><b>' + Footer + '</b></p>');
    L.Add('');
    L.Add('</body>');
    L.Add('</html>');

{$IFDEF UNICODE}
    L.SaveToFile(FileName, TEncoding.UTF8);
{$ELSE}
    L.SaveToFile(FileName);
{$ENDIF}
  finally
    L.Free;
  end;

end;

procedure TAbjGrid.SaveToXML(const FileName: string; Metadat: boolean);
var
  sHeader: string;
  sMeta: string;
  sDataRow: string;
  ARow, ACol: integer;
  L: TStringList;
  S: string;
  Cancel: boolean;
  wSF, wSV: WideString;

const
  NL = #13 + #10;

  function RS(S: string): string;
  var
    i: integer;
  begin
    for i := 1 to Length(S) do
      if S[i] = ' ' then
        S[i] := '_'; // _x0020_  replace it with space
    result := S;
  end;

  function Encode(S: string): string;
  begin
    S := ReplaceStr(S, '<', '&lt;');
    S := ReplaceStr(S, '>', '&gt;');
    S := ReplaceStr(S, '"', '&quot;');
    S := ReplaceStr(S, '&', '&amp;');
    result := S;
  end;

begin
  Cancel := false;
  L := TStringList.Create;
  try

    sHeader := '<?xml version="1.0" encoding="UTF-8"?> ' + NL +
      '<DATAPACKET Version="2.0">';
    L.Add(sHeader + NL);

    if Metadat then
    begin
      sMeta := '<METADATA><FIELDS>';

      sDataRow := '';
      for ACol := 1 to ColCount - 1 do
      begin
        wSF := RS(Cells[ACol, 0]);
        if Trim(wSF) = '' then
          wSF := 'R' + IntToStr(0) + 'C' + IntToStr(ACol);
        sMeta := sMeta + '<FIELD attrname= ' + '"' +
        // AnsiToUtf8(RS(Cells[ACol,0])) +
          wSF + '"' + ' fieldtype="string.uni" WIDTH=' + '"' +
          IntToStr(Length(Cells[ACol, 0])) + '"' + ' />' + NL;
      end;
      sMeta := sMeta + ' </FIELDS><PARAMS/></METADATA>';
      L.Add(sMeta + NL);
    end;

    L.Add('<ROWDATA>');
    for ARow := FixedRows to RowCount - 1 do
    begin
      for ACol := 1 to ColCount - 1 do
      begin
        S := Cells[ACol, ARow];
        if assigned(OnExport) then
          OnExport(self, etXML, ACol, ARow, S, Cancel);
        if Cancel then
        begin
          break;
        end;
        wSF := RS(Cells[ACol, 0]);
        wSV := Encode(S);
        if (Trim(wSF) <> '') or (Trim(wSV) <> '') then
        begin
          if Trim(wSF) = '' then
            wSF := 'R' + IntToStr(ARow) + 'C' + IntToStr(ACol);

          sDataRow := sDataRow + wSF + '=' + '"' + wSV + '" ';
        end;
      end;
      sDataRow := '<ROW ' + sDataRow + ' />';
      L.Add(sDataRow);
      sDataRow := '';
    end;
    L.Add('</ROWDATA></DATAPACKET>');

{$IFDEF UNICODE}
    L.SaveToFile(FileName, TEncoding.UTF8);
{$ELSE}
    L.SaveToFile(FileName);
{$ENDIF}
  finally
    L.Free;
  end;

end;

procedure TAbjGrid.SetAutoRowNumber(const Value: boolean);
begin
  if FAutoRowNumber <> Value then
  begin
    FAutoRowNumber := Value;
    Invalidate_AbjGrid;
  end;
end;

procedure TAbjGrid.SetFieldDelimiter(const Value: string);
begin
  FFieldDelimiter := Value;
end;

function TAbjGrid.GetFieldDelimiter: string;
begin
  result := #9; // Default
  if UpperCase(FFieldDelimiter) = 'TAB' then
    result := #9
  else if UpperCase(FFieldDelimiter) = 'SEMICOLON' then
    result := ';'
  else if UpperCase(FFieldDelimiter) = 'SPACE' then
    result := #32
  else if UpperCase(FFieldDelimiter) = 'COMMA' then
    result := ','
  else if Length(FFieldDelimiter) > 0 then
    result := FFieldDelimiter;
end;

function TAbjGrid.NewColor(clr: TColor; Value: integer): TColor;
var
  R, g, b: integer;
begin
  if Value > 100 then
    Value := 100;
  clr := ColorToRGB(clr);
  R := clr and $000000FF;
  g := (clr and $0000FF00) shr 8;
  b := (clr and $00FF0000) shr 16;

  R := R + Round((255 - R) * (Value / 100));
  g := g + Round((255 - g) * (Value / 100));
  b := b + Round((255 - b) * (Value / 100));

  // Result := Windows.GetNearestColor(ACanvas.Handle, RGB(r, g, b));
  result := RGB(R, g, b);
end;

function TAbjGrid.GetShadeColor(clr: TColor; Value: integer): TColor;
var
  R, g, b: integer;
begin
  clr := ColorToRGB(clr);
  R := clr and $000000FF;
  g := (clr and $0000FF00) shr 8;
  b := (clr and $00FF0000) shr 16;

  R := (R - Value);
  if R < 0 then
    R := 0;
  if R > 255 then
    R := 255;

  g := (g - Value) + 2;
  if g < 0 then
    g := 0;
  if g > 255 then
    g := 255;

  b := (b - Value);
  if b < 0 then
    b := 0;
  if b > 255 then
    b := 255;

  // Result := Windows.GetNearestColor(ACanvas.Handle, RGB(r, g, b));
  result := RGB(R, g, b);
end;

procedure TAbjGrid.SetAutoColWidth(const Value: boolean);
begin
  FAutoColWidth := Value;
end;

procedure TAbjGrid.SetCommentDelimiter(const Value: string);
begin
  FCommentDelimiter := Value;
end;

procedure TAbjGrid.SetEditor(const Value: TGridEditor);
begin
  FEditor := Value;
end;

procedure TAbjGrid.SetAutoRowHeight(const Value: boolean);
begin
  if FAutoRowHeight <> Value then
  begin
    FAutoRowHeight := Value;
    Invalidate_AbjGrid;
  end;
end;

procedure TAbjGrid.SetAutoRowIncrement(const Value: boolean);
begin
  FAutoRowIncrement := Value;
end;

procedure TAbjGrid.SetAllowEdit(const Value: boolean);
begin
  FAllowEdit := Value;
end;

procedure TAbjGrid.SetAllowPaste(const Value: boolean);
begin
  FAllowPaste := Value;
end;

procedure TAbjGrid.SetAllowRowDelete(const Value: boolean);
begin
  FAllowRowDelete := Value;
end;

procedure TAbjGrid.SetAllowRowInsert(const Value: boolean);
begin
  FAllowRowInsert := Value;
end;

procedure TAbjGrid.SetFixedColor(const Value: TColor);
begin
  if FFixedColor <> Value then
  begin
    FFixedColor := Value;
    InitColors;
    Invalidate_AbjGrid;
  end;
end;
{$IFNDEF VER6U}

function TAbjGrid.CellUnderMouse: TPoint;
var
  P: TPoint;
  ACol, ARow: integer;
begin
  P := Mouse.CursorPos;
  P := ScreenToClient(P);

  MouseToCell(P.X, P.Y, ACol, ARow);
  result.X := ACol;
  result.Y := ARow;
end;

function TAbjGrid.CellUnderMouse(var ACol, ARow: NativeInt): TPoint;
var
  P: TPoint;
begin
  P := CellUnderMouse;
  ACol := P.X;
  ARow := P.Y;
end;

procedure TAbjGrid.ChangeGridOrientation(RightToLeftOrientation: boolean);
var
  Org: TPoint;
  Ext: TPoint;
begin
  if RightToLeftOrientation then
  begin
    Org := Point(ClientWidth, 0);
    Ext := Point(-1, 1);
    SetMapMode(Canvas.Handle, mm_Anisotropic);
    SetWindowOrgEx(Canvas.Handle, Org.X, Org.Y, nil);
    SetViewportExtEx(Canvas.Handle, ClientWidth, ClientHeight, nil);
    SetWindowExtEx(Canvas.Handle, Ext.X * ClientWidth,
      Ext.Y * ClientHeight, nil);
  end
  else
  begin
    Org := Point(0, 0);
    Ext := Point(1, 1);
    SetMapMode(Canvas.Handle, mm_Anisotropic);
    SetWindowOrgEx(Canvas.Handle, Org.X, Org.Y, nil);
    SetViewportExtEx(Canvas.Handle, ClientWidth, ClientHeight, nil);
    SetWindowExtEx(Canvas.Handle, Ext.X * ClientWidth,
      Ext.Y * ClientHeight, nil);
  end;
end;
{$ENDIF}

procedure TAbjGrid.SetAlternatingBackColor(const Value: TColor);
begin
  if FAlternatingBackColor <> Value then
  begin
    FAlternatingBackColor := Value;
    Invalidate_AbjGrid;
  end;
end;

procedure TAbjGrid.SetBackColor(const Value: TColor);
begin
  if FBackColor <> Value then
  begin
    FBackColor := Value;
    Invalidate_AbjGrid;
  end;
end;

procedure TAbjGrid.SetSelectColor(const Value: TColor);
begin
  if FSelectColor <> Value then
  begin
    FSelectColor := Value;
    Invalidate_AbjGrid;
  end;

end;

procedure TAbjGrid.SetGridLineColor(const Value: TColor);
begin
  if FGridLineColor <> Value then
  begin
    FGridLineColor := Value;
    Invalidate_AbjGrid;
  end;

end;

procedure TAbjGrid.Invalidate_AbjGrid;
begin
{$IFDEF VER6U}
  InvalidateGrid;
{$ELSE}
  Invalidate;
{$ENDIF}
end;

procedure TAbjGrid.SetRowHeaderWidth(const Value: integer);
begin
  if FRowHeaderWidth <> Value then
  begin
    FRowHeaderWidth := Value;
    ColWidths[0] := FRowHeaderWidth;
  end;

end;

procedure TAbjGrid.SetColHeaderHeight(const Value: integer);
begin
  if FColHeaderHeight <> Value then
  begin
    FColHeaderHeight := Value;
    RowHeights[0] := FColHeaderHeight;
  end;

end;

procedure TAbjGrid.SetColHeaderFont(const Value: TFont);
begin
  FColHeaderFont.Assign(Value);
  Invalidate_AbjGrid;
end;

procedure TAbjGrid.SetRowHeaderFont(const Value: TFont);
begin
  FRowHeaderFont.Assign(Value);
  Invalidate_AbjGrid;
end;

procedure TAbjGrid.SetAllowSorting(const Value: boolean);
begin
  if FAllowSorting <> Value then
  begin
    FAllowSorting := Value;
    // Invalidate_AbjGrid;
  end;
end;

// ----
function CustomSort(List: TStringList; Index1, Index2: integer): integer;

  function IsDate(sDate: string): boolean;
  var
    d: TDateTime;
  begin
    try
      // StrToDate(sDate);
      result := TryStrToDate(sDate, d);
    except
      result := false;
    end;

  end;

var
  SortOrder: integer;
begin
  { TODO : Case Sentesive emplimentation }
  List[Index1] := UpperCase(List[Index1]);
  List[Index2] := UpperCase(List[Index2]);

  if FSortDec then
    SortOrder := -1
  else
    SortOrder := 1;
  if List[Index1] = List[Index2] then
  begin
    result := 0;
    exit;
  end;

  if (StrValue(List[Index1]) <> 0) and (StrValue(List[Index2]) <> 0) then
  begin
    if StrValue(List[Index1]) > StrValue(List[Index2]) then
      result := SortOrder // 1
    else
      result := -SortOrder; // -1;

  end
  else
  begin
    if (IsDate(List[Index1])) and (IsDate(List[Index2])) // New 5/12/2005
    then
      if StrToDate(List[Index1]) > StrToDate(List[Index2]) then
        result := SortOrder // 1
      else
        result := -SortOrder // -1;
    else
    begin
      if List[Index1] > List[Index2] then
        result := SortOrder // 1
      else
        result := -SortOrder; // -1;
    end;
  end;
end;

function ValInt(S: string): integer;
var
  i, Code: integer;
begin
  Val(S, i, Code);
  if Code = 0 then
    result := i
  else
    result := 0;

end;

function ValReal(S: string): Real;
var
  i: Real;
  Code: integer;
begin
  Val(S, i, Code);
  if Code = 0 then
    result := i
  else
    result := 0;

end;

function StrValue(S: string): Real;
begin
  if Pos('.', S) > 0 then
    result := ValReal(S)
  else
    result := ValInt(S)
end;

function HTMLColor(Color: TColor): string;
begin
  result := IntToHex(Color, 6);
  result := '"#' + Copy(result, 5, 2) + Copy(result, 3, 2) +
    Copy(result, 1, 2) + '"';
end;

function IsDate(d: string): boolean;
var
  dd: TDateTime;
begin
  result := GetDate(d, dd);
end;

function GetDate(S: string; var DT: TDateTime): boolean;
var
  i: integer;
  Year, Month, Day: Word;
  Y, M, d: string;
  LastPos: integer;
begin
  LastPos := 1;
  // suppose d/m/y
  for i := 1 to Length(S) do
  begin
    if S[i] = ' ' then
      delete(S, i, 1);
  end;

  for i := 1 to Length(S) do
  begin
    case S[i] of
      '/', '.', '\':
        S[i] := '-';
    end;
  end;
  for i := 1 to Length(S) do
  begin
    case S[i] of
      '-', '0' .. '9':
        ;
    else
      S[i] := ' ';
    end;
  end;

  d := '';
  for i := 1 to Length(S) do
    if S[i] <> '-' then
    begin
      d := d + S[i];
      LastPos := i + 1;
    end
    else
    begin
      LastPos := i + 1;
      break;
    end;

  M := '';
  for i := LastPos to Length(S) do
    if S[i] <> '-' then
    begin
      M := M + S[i];
      LastPos := i + 1;
    end
    else
    begin
      LastPos := i + 1;
      break;
    end;

  Y := '';
  for i := LastPos to Length(S) do
    if S[i] <> '-' then
      Y := Y + S[i]
    else
    begin
      // LastPos := i + 1;
      break;
    end;
  try

    Day := KVal(Trim(d));
    Month := KVal(Trim(M));
    Year := KVal(Trim(Y));
    if IsValidDate(Year, Month, Day) then
    begin
      DT := EncodeDate(Year, Month, Day);
      result := True;
    end
    else
      result := false;
  except
    DT := date;
    result := false;
  end;
end;

function KVal(S: string): integer;
var
  i, Code: integer;
begin
  Val(S, i, Code);
  if Code = 0 then
    result := i
  else
    result := 0;

end;


// ---------

procedure TAbjGrid.SortGrid(ACol: integer);
var
  l: TstringList;
  i,j : integer;
  Cancel: boolean;
begin
  if (RowCount - FixedRows) < 2 then
    exit;

  Cancel := false;
  if assigned(BeforeSort) then
    BeforeSort(self, col, not FSortDec, Cancel);

  if Cancel then
    exit;

  FGridOnBuild := True;
  Enabled := false;

  try
    screen.cursor := crHourGlass;
    GridSort(self, ACol, 0, not FSortDec);
    if assigned(AfterSort) then
      AfterSort(self, col, not FSortDec, Cancel);

  finally
    Enabled := True;
    FGridOnBuild := false;
    screen.cursor := crdefault;
  end;


{
  FGridOnBuild := true;
  Enabled := false;

  l:= TstringList.Create;
  try
    l.BeginUpdate ;
    l.Assign (Cols[ACol]);

    for i := 0 to l.Count -1 do
      l.Objects[i] := TObject(i);

    for i := 0 to FixedRows -1  do
      l.Delete(i);

    l.CustomSort(@CustomSort);


    for i := 0 to  l.Count - 1 do
    begin
      if integer(l.Objects[i]) <> i then
      begin
         RowMove(integer(l.Objects[i]), i + FixedCols);
         for j := 0 to l.Count -1 do
          if Integer(l.Objects[j]) < integer(l.Objects[i]) then
            l.Objects[j] := TObject(Integer(l.Objects[j])+1);
      end;
    end;
    if assigned(AfterSort) then
      AfterSort(Self, Col, not FSortDec, Cancel);

  finally
    l.Free;
    Enabled := true;
    FGridOnBuild := false;
  end;
 }


end;

procedure TAbjGrid.SortGrid(ACol: integer; SortDec: boolean);
begin
  FSortDec := SortDec;
  SortGrid(ACol);
end;

procedure TAbjGrid.DrawSort(ACol: integer; SortDec: boolean; AColor: TColor);
var
  ARect: TRect;
  X, Y, X2: integer;
begin
  ARect := CellRect(ACol, 0);

  X := ARect.Right - 4;
  Y := ARect.Top + 7;
  X2 := 5;

  if SortDec then
  begin
    Canvas.Pen.Color := GetShadeColor(AColor, 70);
    Canvas.MoveTo(X, Y);
    Canvas.LineTo(X - 10, Y);
    Canvas.LineTo(X - X2, Y + X2);
    if not FFlat then
      Canvas.Pen.Color := NewColor(AColor, 70);
    Canvas.LineTo(X, Y);
    if not FFlat then
      Canvas.Pixels[X, Y] := AColor;

  end
  else
  begin
    Canvas.Pen.Color := GetShadeColor(AColor, 70);
    Canvas.MoveTo(X, Y + X2);
    Canvas.LineTo(X - 10, Y + X2);
    Canvas.LineTo(X - X2, Y);
    if not FFlat then
      Canvas.Pen.Color := NewColor(AColor, 70);
    Canvas.LineTo(X, Y + X2);
    if not FFlat then
      Canvas.Pixels[X, Y + X2] := AColor;
  end;

end;

procedure TAbjGrid.ExchangeCells(Cell_1, Cell_2: TPoint);
var
  s1, S2: string;
begin
  s1 := Cells[Cell_1.X, Cell_1.Y];
  S2 := Cells[Cell_2.X, Cell_2.Y];

  Cells[Cell_1.X, Cell_1.Y] := S2;
  Cells[Cell_2.X, Cell_2.Y] := s1;

end;

procedure TAbjGrid.SetSampleData(const Value: boolean);
begin
  FSampleData := Value;
  if FSampleData then
  begin
    Rows[0].CommaText := ',Caption_A, Caption_B, Caption_D, Caption_C';
    Rows[1].CommaText := ',Column1, Column2, Column3, Column4';
    Rows[2].CommaText := ',Column1, Column2, Column3, Column4';
    Rows[3].CommaText := ',Column1, Column2, Column3, Column4';
    Rows[4].CommaText := ',Column1, Column2, Column3, Column4';
  end
  else
  begin
    Clear;
    ClearFixedCols;
    ClearFixedRows;
  end;
end;

function TAbjGrid.IndexOfCol(ACol: integer; S: string; Start: integer): integer;
begin
  for result := Start to RowCount - 1 do
    if Pos(S, Cols[ACol][result]) = 1 then
      exit;
  result := -1;
end;

procedure TAbjGrid.CalcSizingState(X, Y: integer; var State: TGridState;
  var Index, SizingPos, SizingOfs: integer; var FixedInfo: TGridDrawInfo);
begin
  inherited;
  GridState := State;
end;

procedure TAbjGrid.CopyToClipboard;
var
  ARow, ACol: integer;
  sRow, S: WideString;
  // Edit: TEdit;
begin
  sRow := '';
  for ARow := Selection.Top to Selection.Bottom do
  begin
    for ACol := Selection.Left to Selection.Right do
    begin
      sRow := sRow + #9 + Cells[ACol, ARow];
    end;
    delete(sRow, 1, 1);
    S := S + sRow + #13 + #10;
    sRow := '';
  end;
  delete(S, Length(S) - 1, 2);
  if ((Win32Platform and VER_PLATFORM_WIN32_NT) <> 0) then
    Clipboard_SetBuffer(CF_UNICODETEXT, PWideChar(S)^,
      (Length(S) + 1) * SizeOf(WideChar))
  else
    Clipboard.AsText := S;
end;

procedure TAbjGrid.PasteFromClipboard;
var
  List: TStringList;
  ARow, ACol, i: integer;
  BitMap: TBitmap;
  RowText, S: WideString;
  Cancel: boolean;

  function GetClipboardAsWideText: WideString;
  var
    Data: THandle;
  begin
    Clipboard.Open;
    Data := GetClipboardData(CF_UNICODETEXT);
    try
      if Data <> 0 then
        result := PWideChar(GlobalLock(Data))
      else
        result := '';
    finally
      if Data <> 0 then
        GlobalUnlock(Data);
      Clipboard.Close;
    end;
    if (Data = 0) or (result = '') then
      result := Clipboard.AsText

  end;

begin
  if not AllowPaste then
    exit;

  // if not (goEditing in options) then exit;
  if Clipboard.HasFormat(CF_TEXT) then
    if Clipboard.AsText <> '' then
    begin
      List := TStringList.Create;
      try
        List.Text := GetClipboardAsWideText; // Clipboard.AsText;
        for ARow := 0 to List.Count - 1 do
        begin
          RowText := List[ARow];
          S := '';
          ACol := col;
          for i := 1 to Length(RowText) do
          begin
            if RowText[i] = #9 then
            begin
              Cells[ACol, row + ARow] := S;
              if assigned(OnSetEditText) then
                OnSetEditText(self, ACol, row + ARow, S);
              S := '';
              Inc(ACol);
              if ACol > ColCount then
                break;
            end
            else
              S := S + RowText[i];
          end;

          if assigned(BeforePaste) then
            BeforePaste(self, ACol, ARow, S, Cancel);
          if not Cancel then
          begin
            Cells[ACol, row + ARow] := S;
            if assigned(OnSetEditText) then
              OnSetEditText(self, ACol, row + ARow, S);
          end;
          Cancel := false;
          if (row + ARow) > RowCount then
            break;
        end;
      finally
        List.Free;
      end;
    end;

  if Clipboard.HasFormat(CF_BITMAP) then
  begin
    BitMap := TBitmap.Create;
    try
      BitMap.Assign(Clipboard);
      Canvas.StretchDraw(CellRect(col, row), BitMap);

    finally
      BitMap.Free;
    end;
  end;

end;

procedure TAbjGrid.WMCopy(var Message: TMessage);
begin
  CopyToClipboard;
end;

procedure TAbjGrid.WMPaste(var Message: TMessage);
begin
  PasteFromClipboard;
end;

procedure Clipboard_SetBuffer(Format: Word; var Buffer; Size: integer);
begin
  THackClipboard(Clipboard).SetBuffer(Format, Buffer, Size);
end;

function DrawTextX(Handle: THandle; WS: WideString; ARect: TRect;
  Flags: Longint): integer;
var
  S: string;
begin
  Flags := Flags + DT_NOPREFIX;
  if ((Win32Platform and VER_PLATFORM_WIN32_NT) <> 0) then
    result := DrawTextW(Handle,
{$IFDEF VER8U} WS {$ELSE} PWideChar(WS)
      {$ENDIF}, Length(WS), ARect, Flags)
  else
  begin
    S := WS;
    result := Drawtext(Handle,
{$IFDEF VER8U} S {$ELSE} PChar(S)
      {$ENDIF}, Length(S), ARect, Flags);
  end;
end;

procedure TAbjGrid.MoveToNext;
var
  ARow: integer;
begin
  ARow := row + 1;
  if (ARow < 0) and (RowCount > 0) then
    ARow := RowCount - 1;

  if ARow > RowCount - 1 then
    ARow := RowCount - 1;

  row := ARow;

end;

procedure TAbjGrid.SetTitle(const Value: TStrings);
begin
  // FTitle := Value;
  FTitle.Assign(Value);
end;

procedure TAbjGrid.InvalidateCell(ACol, ARow: integer);
begin
  InvalidateCell(ACol, ARow);
end;

procedure TAbjGrid.InvalidateRowA(ARow: integer);
begin

  InvalidateRow(ARow);
end;

function TAbjGrid.RowUnderMouse(X, Y: integer): integer;
begin
  MouseToCell(X, Y, X, result);
end;

function TAbjGrid.ColUnderMouse(X, Y: integer): integer;
begin
  MouseToCell(X, Y, result, Y);
end;

procedure TAbjGrid.ColWidthsChanged;
begin
  inherited;
  DoOnColumnResize;
end;

procedure TAbjGrid.RowHeightsChanged;
begin
  inherited;
  DoOnRowResize;
end;

procedure TAbjGrid.DoOnColumnResize;
begin
  Editor.DoResize;
  if assigned(FonColumnResize) then
    FonColumnResize(self);
end;

procedure TAbjGrid.DoOnRowResize;
begin
  Editor.DoResize;
  if assigned(FonRowResize) then
    FonRowResize(self);
end;

procedure TAbjGrid.RowSelect(ARow: integer);
var
  SRect: TGridRect;
begin
  SRect.Left := 1;
  SRect.Top := ARow;
  SRect.Right := ColCount - 1;
  SRect.Bottom := ARow;
  Selection := SRect;
end;

function TAbjGrid.GetParentForm: TCustomForm;
var
  TopForm: boolean;
  Control: TControl;
begin
  TopForm := True;
  Control := self;
  while (TopForm or not(Control is TCustomForm)) and (Control.Parent <> nil) do
    Control := Control.Parent;
  if Control is TCustomForm then
    result := TCustomForm(Control)
  else
    result := nil;
end;

procedure Gradient(Col_D, Col_L: TColor; Bmp: TBitmap);
// Author: Doctor Nam
// Page: http://www.swissdelphicenter.ch/torry/showcode.php?id=433
type
  PixArray = array [1 .. 3] of Byte;
var
  rdiv, gdiv, bdiv, H, W: integer;
  P: ^PixArray;
begin
  rdiv := GetRValue(Col_L) - GetRValue(Col_D);
  gdiv := GetgValue(Col_L) - GetgValue(Col_D);
  bdiv := GetbValue(Col_L) - GetbValue(Col_D);

  Bmp.PixelFormat := pf24Bit;

  for H := 0 to Bmp.Height - 1 do
  begin
    P := Bmp.ScanLine[H];
    for W := 0 to Bmp.Width - 1 do
    begin
      P^[1] := GetbValue(Col_L) - Round((H / Bmp.Width) * bdiv);
      P^[2] := GetgValue(Col_L) - Round((H / Bmp.Width) * gdiv);
      P^[3] := GetRValue(Col_L) - Round((H / Bmp.Width) * rdiv);
      Inc(P);
    end;
  end;
end;

{ TGridEditor }

function TGridEditor.CharIndexFromPoint(P: TPoint): integer;
begin
  result := LoWord(Perform(EM_CHARFROMPOS, 0, MakeLParam(P.X, P.Y)));
end;

procedure TGridEditor.CMExit(var Message: TCMExit);
begin
  // visible := false;
end;

procedure TGridEditor.CNCommand(var Message: TWMCommand);
begin
  inherited;
  {
    if (Message.NotifyCode = EN_CHANGE) then
    begin
    if HandleAllocated then
    begin
    if assigned(Grid.EditChange) then
    Grid.EditChange(Self);
    Modified := true;

    end;
    end;
  }
end;

constructor TGridEditor.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Width := 1;
  Height := 1;
  Visible := false;
  Grid := TAbjGrid(AOwner);
  Grid.Options := Grid.Options - [goEditing];
  WantReturns := false;
  Parent := Grid;
  Ctl3D := false;
  Visible := false;
end;

procedure TGridEditor.DoHide;
begin
  try
    if Visible then
    begin
      Put_Content_In_Cell;
      Text := '';
      Visible := false
    end;
  except

  end;
end;

procedure TGridEditor.DoResize;
var
  Rect: TRect;
  ParentForm: TCustomForm;
begin
  if self = nil then
    exit;
  if Grid = nil then
    exit;
  if not Grid.Visible then
    exit;
  if not Visible then
    exit;

  ParentForm := nil;
  ParentForm := Grid.GetParentForm;
  if ParentForm = nil then
    exit;

  if TForm(ParentForm).Visible then
  begin
    Rect := Grid.CellRect(Grid.col, Grid.row);
    InflateRect(Rect, -1, -2);
    BoundsRect := Rect;
  end;

end;

procedure TGridEditor.KeyUp(var Key: Word; Shift: TShiftState);
var
  H: integer;
  P: integer;
begin
  case Key of
    13:
      begin
        Key := 0;
        if ssAlt in Shift then
        begin
          P := SelStart;
          Text := Copy(Text, 1, P) + #13 + #10 +
            Copy(Text, P + 1, Length(Text));
          SelStart := P + 2;
        end
        else
        begin
          Put_Content_In_Cell;
          Text := '';
          Visible := false;
          Grid.SetFocus;
        end;
      end;
    VK_ESCAPE:
      begin
        Modified := false;
        Text := '';
        Visible := false;
        Grid.SetFocus;
      end;

  end;

  H := Grid.DetermineCellHeight(Text, Grid.col, Grid.row);
  if (H > Height) and ((Top + H) < Grid.Top + Grid.Height) then
  begin
    Height := H;
  end;

  inherited KeyUp(Key, Shift);
end;

procedure TGridEditor.Put_Content_In_Cell;
begin
  if ReadOnly then
    exit;

  Modified := (Grid.Cells[Grid.col, Grid.row] <> Text);
  if Modified then
  begin
    Grid.Cells[Grid.col, Grid.row] := Text;
    if assigned(Grid.AfterEdit) then
      Grid.AfterEdit(self, Grid.col, Grid.row, Text);
  end;

  Application.ProcessMessages;

end;

procedure TGridEditor.Show_Edit_In_Cell;
var
  ACol, ARow: NativeInt;
begin
  Grid.CellUnderMouse(ACol, ARow);
  if (ARow <> -1) and (ACol <> -1) then
  begin
    Show_Edit_In_Cell(ACol, ARow, True, '')
  end;

end;

procedure TGridEditor.Show_Edit_In_Cell(ACol, ARow: integer;
  CaretUnderMosue: boolean; const AText: string);
var
  Rect: TRect;
  P: TPoint;
  S: string;
begin
  if ACol < Grid.FixedCols then
    exit;
  if ARow < Grid.FixedRows then
    exit;

  P := Mouse.CursorPos;

  Rect := Grid.CellRect(ACol, ARow);
  BoundsRect := Rect;
  Font.Assign(Grid.Font);
  S := Grid.Cells[ACol, ARow];

  if AText <> '' then
    S := S + AText;
  Text := S;
  Modified := false;
  // BiDiMode := Grid.BiDiMode;
  Visible := True;
  SetFocus;

  P := ScreenToClient(P);

  if CaretUnderMosue then
  begin
    SelStart := CharIndexFromPoint(P);
  end
  else
  begin
    SelStart := Length(Text);
  end;

  if assigned(Grid.BeforeEdit) then
    Grid.BeforeEdit(self, Grid.col, Grid.row, Text);
end;

function BlendColor(Color1, Color2: TColor; A: Byte): TColor;
var
  c1, c2: Longint;
  R, g, b, v1, v2: Byte;
begin
  A := Round(2.55 * A);
  c1 := ColorToRGB(Color1);
  c2 := ColorToRGB(Color2);
  v1 := Byte(c1);
  v2 := Byte(c2);
  R := A * (v1 - v2) shr 8 + v2;
  v1 := Byte(c1 shr 8);
  v2 := Byte(c2 shr 8);
  g := A * (v1 - v2) shr 8 + v2;
  v1 := Byte(c1 shr 16);
  v2 := Byte(c2 shr 16);
  b := A * (v1 - v2) shr 8 + v2;
  result := (b shl 16) + (g shl 8) + R;
end;

end.
