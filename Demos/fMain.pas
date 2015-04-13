unit fMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.Grids, AbjGrid;

type
  TForm2 = class(TForm)
    Panel1: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GridMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
  private
    { Private declarations }
    Grid: TAbjGrid;
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

procedure TForm2.FormCreate(Sender: TObject);
begin
  Grid := TAbjGrid.Create(self);
  Grid.Parent := self;

  Grid.AutoRowNumber := true;
  Grid.AllowRowDelete := true;
  Grid.AllowRowInsert := true;
  Grid.AllowSorting := true;
  Grid.AllowPaste := true;
  Grid.AllowEdit := true;
  Grid.AutoRowHeight := true;
  Grid.AutoColWidth := true;

  Grid.AutoRowIncrement := true;
  Grid.AlternatingBackColor := $00F2F2F2;

//  Grid.ColHeaderFont.Size := 10;
  Grid.ColHeaderFont.Color := clNavy;
  Grid.RowHeaderFont.Color := clBlue;

  //Grid.SelectColor := $0096C8FE;

  Grid.Options := Grid.Options + [goRowSizing, goRowMoving, goColSizing, goColMoving,
                   goRangeSelect, goDrawFocusSelected];


  Grid.Align := alClient;

  Grid.OnMouseMove := GridMouseMove;

  Grid.LoadFromFile('SampleData.txt', true);
  Grid.Visible := true;

end;


procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Grid.Free;
end;

procedure TForm2.GridMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var
  P: TPoint;
begin
  P := Grid.CellUnderMouse;
  if (P.X > -1) and (P.Y > -1) then
    Panel1.Caption := Grid.Cells[P.X, P.Y]
  else
    Panel1.Caption := '';
end;

end.
