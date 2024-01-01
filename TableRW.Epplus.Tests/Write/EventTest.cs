using OfficeOpenXml;

namespace TableRW.Write.Epplus.Tests;
public class EventTest : IDisposable {

    record EntityA(int Int1, int Int2, int? NullableInt, string Str1, DateTime Date);

    List<EntityA> DataList = new() {
        new(1, 11, 1000, "aaa", new(2021, 1, 11)),
        new(3, 33, null, "ccc", new(2022, 3, 13)),
        new(7, 77, 7000, "bbb", new(2023, 7, 17)),
    };

    public void Dispose() => _excel.Dispose();
    readonly ExcelPackage _excel;
    readonly ExcelWorksheet _sheet;
    readonly ExcelRange _cells;

    public EventTest() {
        _excel = new ExcelPackage(new MemoryStream(2 * 1024));
        _sheet = _excel.Workbook.Worksheets.Add("sheet1");
        _cells = _sheet.Cells;
    }

    [Fact]
    public void StartWritingTable() {
        var writer = new ExcelRangeWriter<EntityA>()
            .SetStart(1, 2)
            .OnStartWritingTable(it => {
                Assert.Equal(2, it.iCol);
                it.Src[1, 1].Value = "AA1";
            })
            .AddColumns((s, e) => s(e.Int1, e.Str1));

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_cells, DataList);

        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);
        Assert.Equal("AA1", _cells[1, 1].Value);

        for (var i = 0; i < DataList.Count;) {
            var col = 2;
            var e = DataList[i];
            i++;
            Assert.Equal(e.Int1, _cells[i, col++].Value);
            Assert.Equal(e.Str1, _cells[i, col++].Value);
        }
    }

    [Fact]
    public void StartWritingRow() {
        var writer = new ExcelRangeWriter<EntityA>()
            .SetStart(1, 2)
            .OnStartWritingRow(it => it.Src[it.iRow ,it.iCol - 1].Value = 222)
            .AddColumns((s, e) => s(e.Int1, e.Int2, e.Str1));

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_cells, DataList);

        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);

        for (var i = 0; i < DataList.Count;) {
            var col = 1;
            var e = DataList[i];
            i++;
            Assert.Equal(222, _cells[i, col++].Value);
            Assert.Equal(e.Int1, _cells[i, col++].Value);
            Assert.Equal(e.Int2, _cells[i, col++].Value);
            Assert.Equal(e.Str1, _cells[i, col++].Value);
        }
    }

    [Fact]
    public void EndWritingRow() {
        var writer = new ExcelRangeWriter<EntityA>()
            .AddColumns((s, e) => s(e.Int1, e.Int2, e.Date.Day))
            .OnEndWritingRow(it => it.Src[it.iRow, it.iCol].Value = 222);

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_cells, DataList);

        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);

        for (var i = 0; i < DataList.Count;) {
            var col = 1;
            var e = DataList[i];
            i++;
            Assert.Equal(e.Int1, _cells[i, col++].Value);
            Assert.Equal(e.Int2, _cells[i, col++].Value);
            Assert.Equal(222, _cells[i, col++].Value);
        }
    }

    [Fact]
    public void EndWritingTable() {
        var writer = new ExcelRangeWriter<EntityA>()
            .AddColumns((s, e) => s(e.Int1, e.Int2, e.Date.Day))
            .OnEndWritingTable(it => it.Src[it.iRow, it.iCol + 1].Value = 444);

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_cells, DataList);

        Assert.Equal(DataList.Count + 1, _sheet.Dimension.Rows);

        for (var i = 0; i < DataList.Count;) {
            var col = 1;
            var e = DataList[i];
            i++;
            Assert.Equal(e.Int1, _cells[i, col++].Value);
            Assert.Equal(e.Int2, _cells[i, col++].Value);
            Assert.Equal(e.Date.Day, _cells[i, col++].Value);
        }
        Assert.Equal(444, _cells[DataList.Count + 1, 4].Value);
    }
}
