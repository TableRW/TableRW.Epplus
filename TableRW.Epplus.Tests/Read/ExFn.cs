

using OfficeOpenXml;

namespace TableRW.Read.Epplus.Tests;

static class ExcelEx {
    public static void WriteCells(this ExcelWorksheet sheet, (int row, int col) start, object?[][] values) {
        for (int iRow = 0; iRow < values.Length; iRow++) {
            for (int iCol = 0; iCol < values[iRow].Length; iCol++) {
                sheet.Cells[start.row + iRow, start.col + iCol].Value = values[iRow][iCol];
            }
        }
    }
}