unit DBGridHelper;

interface

uses
  Vcl.DBGrids, Vcl.Graphics, System.Classes, System.Types, Vcl.Themes, Vcl.Grids, Data.DB;

type
  TDBGridHelper = class helper for TDBGrid
  private
    procedure DrawCheckbox(ARect: TRect; Checked: Boolean; Canvas: TCanvas);
    procedure CustomDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure CustomColEnter(Sender: TObject);
  public
    procedure EnableCheckboxForBooleanField;
    procedure SetEditingEnabled(Value: Boolean); // Método para configurar o dgEditing
  end;

implementation

{ TDBGridHelper }

procedure TDBGridHelper.SetEditingEnabled(Value: Boolean);
begin
  if Value then
    Self.Options := Self.Options + [dgEditing]
  else
    Self.Options := Self.Options - [dgEditing];
end;

procedure TDBGridHelper.EnableCheckboxForBooleanField;
begin
  Self.OnDrawColumnCell := CustomDrawColumnCell;
  Self.OnColEnter := CustomColEnter;
end;

procedure TDBGridHelper.CustomColEnter(Sender: TObject);
begin
  // Desabilitar edição para colunas booleanas
  if Assigned(Self.SelectedField) and (Self.SelectedField.DataType = ftBoolean) then
    SetEditingEnabled(False)
  else
    SetEditingEnabled(True);
end;

procedure TDBGridHelper.CustomDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  IsChecked: Boolean;
  Field: TField;
begin
  Field := Column.Field;
  if Assigned(Field) and (Field.DataType = ftBoolean) then
  begin
    Self.Canvas.FillRect(Rect);
    IsChecked := Field.AsBoolean;
    DrawCheckbox(Rect, IsChecked, Self.Canvas);
    Exit;
  end;

  DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TDBGridHelper.DrawCheckbox(ARect: TRect; Checked: Boolean; Canvas: TCanvas);
var
  CheckboxRect: TRect;
  Details: TThemedElementDetails;
begin
  CheckboxRect := ARect;
  CheckboxRect.Left := ARect.Left + (ARect.Width - 16) div 2;
  CheckboxRect.Top := ARect.Top + (ARect.Height - 16) div 2;
  CheckboxRect.Right := CheckboxRect.Left + 16;
  CheckboxRect.Bottom := CheckboxRect.Top + 16;

  if TStyleManager.IsCustomStyleActive then
  begin
    if Checked then
      Details := StyleServices.GetElementDetails(tbCheckBoxCheckedNormal)
    else
      Details := StyleServices.GetElementDetails(tbCheckBoxUncheckedNormal);
    StyleServices.DrawElement(Canvas.Handle, Details, CheckboxRect);
  end
  else
  begin
    Canvas.Rectangle(CheckboxRect);
    if Checked then
    begin
      Canvas.Pen.Color := clBlack;
      Canvas.MoveTo(CheckboxRect.Left, CheckboxRect.Top);
      Canvas.LineTo(CheckboxRect.Right, CheckboxRect.Bottom);
      Canvas.MoveTo(CheckboxRect.Left, CheckboxRect.Bottom);
      Canvas.LineTo(CheckboxRect.Right, CheckboxRect.Top);
    end;
  end;
end;


end.

