using OfficeOpenXml;

namespace TableRW.Write.Epplus.Tests;

public class ExcelRangeExTest : IDisposable {

    record EntityA(int Int1, int Int2, int? NullableInt, string Str1, DateTime Date);

    List<EntityA> DataList = new() {
        new(10, 11, 1000, "aaa", new(2023, 1, 1)),
        new(30, 33, 3000, "ccc", new(2023, 3, 3)),
        new(70, 77, 7000, "bbb", new(2023, 7, 7)),
    };

    public void Dispose() => _excel.Dispose();
    readonly ExcelPackage _excel;
    readonly ExcelWorksheet _sheet;
    readonly ExcelRange _cells;

    public ExcelRangeExTest() {
        _excel = new ExcelPackage(new MemoryStream(2 * 1024));
        _sheet = _excel.Workbook.Worksheets.Add("sheet1");
        _cells = _sheet.Cells;
    }


    [Fact]
    public void WriteFrom() {
        _cells.WriteFrom(DataList, cacheKey: 0, writer => {
            writer.AddColumns((s, e) =>
                s(e.Int1, s.Skip(1), e.NullableInt, e.Str1));

            var lmd = writer.Lambda();
            return lmd.Compile();
        });

        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);
        for (var i = 0; i < DataList.Count;) {
            var col = 1;
            var e = DataList[i];
            i++;
            Assert.Equal(e.Int1, _cells[i, col++].Value);
            Assert.Equal(null, _cells[i, col++].Value);
            Assert.Equal(e.NullableInt, _cells[i, col++].Value);
            Assert.Equal(e.Str1, _cells[i, col++].Value);
            Assert.Equal(null, _cells[i, col++].Value);
        }
    }

    [Fact]
    public void WriteFrom_AnotherKey() {
        _cells.WriteFrom(DataList, cacheKey: 1, writer => {
            writer.AddColumns((s, e) =>
                s(e.Int1, s.Skip(1), e.NullableInt, e.Str1));

            var lmd = writer.Lambda();
            return lmd.Compile();
        });
        _sheet.DeleteRow(1, 3);

        // 使用另一种方式，和上面的缓存不同
        _cells.WriteFrom(DataList, cacheKey: 2, writer => {
            writer.AddColumns((s, e) =>
                s(e.Int2, e.Int1, s.Skip(2), e.Str1));

            var lmd = writer.Lambda();
            return lmd.Compile();
        });

        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);
        for (var i = 0; i < DataList.Count;) {
            var e = DataList[i];
            var col = 1;
            i++;
            Assert.Equal(e.Int2, _cells[i, col++].Value);
            Assert.Equal(e.Int1, _cells[i, col++].Value);
            Assert.Equal(null, _cells[i, col++].Value);
            Assert.Equal(null, _cells[i, col++].Value);
            Assert.Equal(e.Str1, _cells[i, col++].Value);
            Assert.Equal(null, _cells[i, col++].Value);
        }
    }

    [Fact]
    public void WriteFrom_WithData() {
        _cells[1, 1].Value = "AA1";
        _cells.WriteFrom<EntityA, string?>(DataList, cacheKey: 3, writer => {
            writer
                .InitData(src => src[1, 1].Value.ToString())
                .AddColumn(it => it.SetColumnValue(it.Data))
                .AddColumns((s, e) => s(e.Int1, e.Int2, e.Str1));

            var lmd = writer.Lambda();
            return lmd.Compile();
        });

        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);
        for (var i = 0; i < DataList.Count;) {
            var e = DataList[i];
            var col = 1;
            i++;
            Assert.Equal("AA1", _cells[i, col++].Value);
            Assert.Equal(e.Int1, _cells[i, col++].Value);
            Assert.Equal(e.Int2, _cells[i, col++].Value);
            Assert.Equal(e.Str1, _cells[i, col++].Value);
        }
    }

}
