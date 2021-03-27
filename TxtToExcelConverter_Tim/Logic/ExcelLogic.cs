using ClosedXML.Excel;
using TxtToExcelConverter_Tim.Models;

namespace TxtToExcelConverter_Tim.Logic
{
    public static class ExcelLogic
    {
        public static XLWorkbook GetExcel(TableModel[] models)
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet sheet = wb.Worksheets.Add("Sheet1");

            // ЗАГОЛОВКИ
            IXLRange range = sheet.Range(sheet.Cell(1, "A"), sheet.Cell(2, "I"));
            // заливка
            range.Style.Fill.SetBackgroundColor(XLColor.FromHtml("#fcecdc"));
            // жирный шрифт
            range.Style.Font.SetBold();
            // выравнивание
            range.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

            // ширина заголовков
            double factor = 1.145; // множитель (чтобы ширина была как у оригинального excel)

            sheet.Column("A").Width = 24.13 * factor;
            sheet.Column("B").Width = 33.63 * factor;
            sheet.Column("C").Width = 11.88 * factor;
            sheet.Column("D").Width = 9.25 * factor;
            sheet.Column("E").Width = 10.5 * factor;
            sheet.Column("F").Width = 52.75 * factor;
            sheet.Column("G").Width = 8.75 * factor;
            sheet.Column("H").Width = 8.75 * factor;
            sheet.Column("I").Width = 33.25 * factor;

            // подписи у заголовков и объединение ячеек для заголовков
            sheet.Range(sheet.Cell(1, "A"), sheet.Cell(2, "A"))
                .Merge().Value = "Component type";
            sheet.Range(sheet.Cell(1, "B"), sheet.Cell(2, "B"))
                .Merge().Value = "Name";
            sheet.Range(sheet.Cell(1, "C"), sheet.Cell(2, "C"))
                .Merge().Value = "Nominal";
            sheet.Range(sheet.Cell(1, "D"), sheet.Cell(2, "D"))
                .Merge().Value = "Deviation";
            sheet.Range(sheet.Cell(1, "E"), sheet.Cell(2, "E"))
                .Merge().Value = "Case type";
            sheet.Range(sheet.Cell(1, "F"), sheet.Cell(2, "F"))
                .Merge().Value = "Comment";
            sheet.Range(sheet.Cell(1, "G"), sheet.Cell(1, "H"))
                .Merge().Value = "quantity";
            sheet.Cell(2, "G").Value = "00";
            sheet.Cell(2, "H").Value = "01";
            sheet.Range(sheet.Cell(1, "I"), sheet.Cell(2, "I"))
                .Merge().Value = "Remark:";

            // границы у заголовков
            XLBorderStyleValues styleBorder = XLBorderStyleValues.Medium;

            range.Style.Border.SetOutsideBorder(styleBorder);
            range.Style.Border.SetInsideBorder(styleBorder);

            #region ЛОГИКА ВЫВОДА ДАННЫХ

            int currentRow = 3;
            IXLCells cellRange;

            for (int i = 0; i < models.Length; i++)
            {
                TableModel tm = models[i];

                #region ЛОГИКА ОТДЕЛЕНИЯ ГРУПП КОМПОНЕНТОВ

                // чтобы не выйти за пределы массива
                if (i > 0 && tm.ComponentType != models[i - 1].ComponentType)
                {
                    range = sheet.Range(sheet.Cell(currentRow, "A"), sheet.Cell(currentRow, "i")).Merge();

                    // границы
                    range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);

                    // заливка
                    range.Style.Fill.SetBackgroundColor(XLColor.FromHtml("#f8f4f4"));

                    currentRow++;
                }

                #endregion

                cellRange = sheet.Row(currentRow).Cells("A", "I");

                // выравнивание текста
                cellRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                cellRange.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                // границы
                styleBorder = XLBorderStyleValues.Thin;

                cellRange.Style.Border.SetInsideBorder(styleBorder);
                cellRange.Style.Border.SetOutsideBorder(styleBorder);

                sheet.Cell(currentRow, "A").Value = tm.ComponentType;
                sheet.Cell(currentRow, "B").Value = tm.Name;
                sheet.Cell(currentRow, "C").Value = tm.Nominal;
                sheet.Cell(currentRow, "D").Value = tm.Deviation;
                sheet.Cell(currentRow, "E").SetValue(tm.CaseType);
                sheet.Cell(currentRow, "F").Value = tm.Comment;
                sheet.Cell(currentRow, "G").Value = tm.Quanity.ZeroZero;
                sheet.Cell(currentRow, "H").Value = tm.Quanity.ZeroOne;

                currentRow++;
            }

            #endregion

            // СТИЛИ ДЛЯ ВСЕХ ЯЧЕЕК
            cellRange = sheet.CellsUsed();

            // шрифт
            cellRange.Style.Font.SetFontName("Arial Cyr");
            cellRange.Style.Font.SetFontSize(9);

            return wb;
        }
    }
}
