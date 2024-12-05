unit GDExportExcelHelper;

interface

uses
  System.SysUtils, System.Classes, Data.DB, System.Win.ComObj, System.IOUtils, System.StrUtils, System.Variants;

type
  TGDExportExcel = class
  public
    class procedure ExportToExcel(const ADataSet: TDataSet; const AFileNamePrefix: string;
      const AExcludeFields: TArray<string>; const AIncludeHeader: Boolean = True);
  end;

implementation

uses
  Winapi.ShellAPI, Winapi.Windows;

class procedure TGDExportExcel.ExportToExcel(const ADataSet: TDataSet; const AFileNamePrefix: string;
  const AExcludeFields: TArray<string>; const AIncludeHeader: Boolean);
var
  ExcelApp: Variant;
  Worksheet: Variant;
  Column, Row, FieldIndex: Integer;
  SaveFilePath, UniqueFileName, ReportsDir: string;
  Field: TField;
begin
  if not Assigned(ADataSet) then
    raise Exception.Create('Dataset is not assigned.');

  if ADataSet.IsEmpty then
    raise Exception.Create('Dataset is empty, nothing to export.');

  // Criar o diretório "relatorios" na mesma pasta do executável
  ReportsDir := TPath.Combine(ExtractFilePath(ParamStr(0)), 'relatorios');
  if not TDirectory.Exists(ReportsDir) then
    TDirectory.CreateDirectory(ReportsDir);

  // Gerar nome único do arquivo
  UniqueFileName := Format('%s_%s.xlsx', [AFileNamePrefix, TGUID.NewGuid.ToString.Replace('{', '').Replace('}', '').Replace('-', '')]);

  // Definir caminho completo do arquivo
  SaveFilePath := TPath.Combine(ReportsDir, UniqueFileName);

  // Criação do aplicativo Excel
  ExcelApp := CreateOleObject('Excel.Application');
  try
    ExcelApp.Visible := False; // Torne o Excel invisível durante a operação
    Worksheet := ExcelApp.Workbooks.Add(1).WorkSheets[1];

    // Adicionar cabeçalho das colunas, se necessário
    Row := 1;
    if AIncludeHeader then
    begin
      Column := 1;
      for FieldIndex := 0 to ADataSet.FieldCount - 1 do
      begin
        Field := ADataSet.Fields[FieldIndex];
        if not MatchText(Field.FieldName, AExcludeFields) then
        begin
          Worksheet.Cells[Row, Column] := Field.FieldName;
          Inc(Column);
        end;
      end;
      Inc(Row);
    end;

    // Adicionar dados do DataSet
    ADataSet.First;
    while not ADataSet.Eof do
    begin
      Column := 1;
      for FieldIndex := 0 to ADataSet.FieldCount - 1 do
      begin
        Field := ADataSet.Fields[FieldIndex];
        if not MatchText(Field.FieldName, AExcludeFields) then
        begin
          Worksheet.Cells[Row, Column] := Field.AsString;
          Inc(Column);
        end;
      end;
      Inc(Row);
      ADataSet.Next;
    end;

    // Salvar o arquivo Excel no caminho especificado
    Worksheet.Parent.SaveAs(SaveFilePath);

    // Abrir o arquivo gerado
    ShellExecute(0, 'open', PChar(SaveFilePath), nil, nil, SW_SHOWNORMAL);

  finally
    ExcelApp.Quit; // Encerrar o Excel
  end;
end;

end.
