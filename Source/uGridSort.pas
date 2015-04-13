unit uGridSort;

{
  The core of this unit takent from:
  http://www.delphiforfun.org/programs/delphi_techniques/GridQuickSort.htm

  It has the following copright notice:
  "opyright  © 2002-2005, Gary Darby,  www.DelphiForFun.org
  This program may be used or modified for any non-commercial purpose
  so long as this original notice remains in place.
  All other rights are reserved"
}
interface

uses

  Windows, SysUtils, Messages, Classes, Graphics, Controls,
  Grids, StdCtrls, DateUtils, StrUtils, Types;

procedure GridSort(Grid: TStringGrid; const SortCol: integer;
  const DataType: integer; const Ascending: boolean);

procedure MergeSort(Grid: TStringGrid; var arrRows: array of integer;
  SortCol, DataType: integer; Ascending: boolean);

implementation

function ValInt(S: string): integer;
var
  I, Code: integer;
begin
  Val(S, I, Code);
  if Code = 0 then
    result := I
  else
    result := 0;
end;

function ValReal(S: string): Real;
var
  I: Real;
  Code: integer;
begin
  Val(S, I, Code);
  if Code = 0 then
    result := I
  else
    result := 0;
end;

function StrValue(S: string): Real;
begin
  S := ReplaceStr(S, ',', '');
  if Pos('.', S) > 0 then
    result := ValReal(S)
  else
    result := ValInt(S)
end;

function IsDate(sDate: string): boolean;
var
  d: TDateTime;
begin
  try
    result := TryStrToDate(sDate, d);
  except
    result := false;
  end;

end;

procedure GridSort(Grid: TStringGrid; const SortCol: integer;
  const DataType: integer; const Ascending: boolean);

var
  I: integer;
  GridTemp: TStringGrid;
  arrRows: array of integer;
begin
  GridTemp := TStringGrid.create(nil);

  GridTemp.RowCount := Grid.RowCount;
  GridTemp.ColCount := Grid.ColCount;
  GridTemp.FixedRows := Grid.FixedRows;

  Setlength(arrRows, Grid.RowCount - Grid.FixedRows);
  for I := Grid.FixedRows to Grid.RowCount - 1 do
  begin
    arrRows[I - Grid.FixedRows] := I;
    GridTemp.rows[I].assign(Grid.rows[I]);
  end;
  MergeSort(Grid, arrRows, SortCol, DataType, Ascending);

  Grid.DefaultRowHeight := Grid.DefaultRowHeight;
  for I := 0 to Grid.RowCount - Grid.FixedRows - 1 do
  begin
    Grid.rows[I + Grid.FixedRows].assign(GridTemp.rows[arrRows[I]])
  end;
  Grid.Row := Grid.FixedRows;

  GridTemp.free;
  Setlength(arrRows, 0);

end;

procedure MergeSort(Grid: TStringGrid; var arrRows: array of integer;
  SortCol, DataType: integer; Ascending: boolean);

var
  arrTempRows: array of integer;

  // ---------------------------------------------------
  function Compare(val1, val2: string): integer;
  var
    SortOrder: integer;
  begin
    // Hear detecting data type case by case, it's slower then pre determind type
    // but give more logic results
    // if Ascending then SortOrder := 1 else SortOrder := -1;
    SortOrder := 1;

    if val1 = val2 then
    begin
      result := 0;
      exit;
    end;

    if val1 = '' then
      val1 := ' ';

    if val1[1] in ['.', ',', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    then
    else
    begin
      val1 := Copy(val1, 1, 100); // to test against speed
      val2 := Copy(val2, 1, 100);

      val1 := UpperCase(val1);
      val2 := UpperCase(val2);

      if val1 > val2 then
        result := SortOrder // 1
      else
        result := -SortOrder; // -1;
      exit;
    end;

    if (StrValue(val1) <> 0) and (StrValue(val2) <> 0) then
    begin
      if StrValue(val1) > StrValue(val2) then
        result := SortOrder // 1
      else
        result := -SortOrder; // -1;

    end
    else
    begin
      if IsDate(val1) and IsDate(val2) // New 5/12/2005
      then
        if StrToDate(val1) > StrToDate(val2) then
          result := SortOrder // 1
        else
          result := -SortOrder // -1;
      else
      begin
        if val1 > val2 then
          result := SortOrder // 1
        else
          result := -SortOrder; // -1;
      end;
    end;

  end;

{ ---------- Merge ------------- }
  procedure Merge(ALo, AMid, AHi: integer);
  var
    I, j, k, m, n: integer;
  begin
    I := 0;
    Setlength(arrTempRows, AMid - ALo + 1);
    for j := ALo to AMid do
    begin
      { copy lower half of Vals into temporary array AVals }
      arrTempRows[I] := arrRows[j];
      inc(I);
    end;

    I := 0;
    j := AMid + 1;
    k := ALo;
    while ((k < j) and (j <= AHi)) do
    begin
      { Merge: Compare upper half to copied verasion of the lower half and move the
        appropriate value (smallest for ascending, largest for descending) into
        the lower half positions, for equals use Avals to preserve original order }
      // with Grid do
      n := Compare(Grid.Cells[SortCol, arrRows[j]],
        Grid.Cells[SortCol, arrTempRows[I]]);

      if Ascending and (n >= 0) or ((not Ascending) and (n <= 0)) then
      begin
        arrRows[k] := arrTempRows[I];
        inc(I);
        inc(k);
      end
      else
      begin
        arrRows[k] := arrRows[j];
        inc(k);
        inc(j);
      end;
    end;

    { copy any remaining, unsorted, elements }
    for m := k to j - 1 do
    begin
      arrRows[m] := arrTempRows[I];
      inc(I);
    end;
  end;

{ ------------ PerformMergeSort ------------ }
  procedure PerformMergeSort(ALo, AHi: integer);
  { recursively split the split the value into 2 pieces and merge them back
    together as we unwind the recursion }
  var
    AMid: integer;
  begin
    if (ALo < AHi) then
    begin
      AMid := (ALo + AHi) shr 1;
      PerformMergeSort(ALo, AMid);
      PerformMergeSort(AMid + 1, AHi);
      Merge(ALo, AMid, AHi);
    end;
  end;

begin
  PerformMergeSort(0, High(arrRows));
end;

end.
