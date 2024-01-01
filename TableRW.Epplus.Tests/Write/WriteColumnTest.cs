using OfficeOpenXml;

namespace TableRW.Write.Epplus.Tests;

public class ExcelRangeWriterTest : IDisposable {

    record EntityA(int Int1, int Int2, int? NullableInt, string Str1, DateTime Date);

    List<EntityA> DataList = new() {
        new(1, 11, 1000, "aaa", new(2023, 1, 1)),
        new(3, 33, null, "ccc", new(2023, 3, 3)),
        new(7, 77, 7000, "bbb", new(2023, 7, 7)),
    };

    public void Dispose() => _excel.Dispose();
    readonly ExcelPackage _excel;
    readonly ExcelWorksheet _sheet;
    readonly ExcelRange _cells;

    public ExcelRangeWriterTest() {
        _excel = new ExcelPackage(new MemoryStream(2 * 1024));
        _sheet = _excel.Workbook.Worksheets.Add("sheet1");
        _cells = _sheet.Cells;
    }


    [Fact]
    public void AddColumns() {
        var writer = new ExcelRangeWriter<EntityA>()
            .AddColumns((s, e) => s(e.Int1, s.Skip(1), e.NullableInt, e.Str1));

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_cells, DataList);

        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);

        var iCol = 1;
        Assert.Equal(DataList[0].Int1, _cells[1, iCol++].Value);
        Assert.Equal(null, _cells[1, iCol++].Value);
        Assert.Equal(DataList[0].NullableInt, _cells[1, iCol++].Value);
        Assert.Equal(DataList[0].Str1, _cells[1, iCol++].Value);

        iCol = 1;
        Assert.Equal(DataList[1].Int1, _cells[2, iCol++].Value);
        Assert.Equal(null, _cells[2, iCol++].Value);
        Assert.Equal(null, _cells[2, iCol++].Value);
        Assert.Equal(DataList[1].Str1, _cells[2, iCol++].Value);
    }


    [Fact]
    public void AddColumns_Compute() {
        var writer = new ExcelRangeWriter<EntityA>()
            .AddColumns((s, e) => s(
                e.Int1, e.Int1 + e.Int2, e.NullableInt + 1000,
                $"{e.Str1} -- {e.Date.Month}",
                "AA " + DateTime.Now.Month
            ));

        var writeLmd = writer.Lambda();
        var writeFn = writeLmd.Compile();
        writeFn(_cells, DataList);

        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);


        var (row, col) = (1, 1);
        var e = DataList[row - 1];
        Assert.Equal(e.Int1, _cells[row, col++].Value);
        Assert.Equal(e.Int1 + e.Int2, _cells[row, col++].Value);
        Assert.Equal(e.NullableInt + 1000, _cells[row, col++].Value);
        Assert.Equal($"{e.Str1} -- {e.Date.Month}", _cells[row, col++].Value);
        Assert.Equal("AA " + DateTime.Now.Month, _cells[row, col++].Value);

        (row, col) = (2, 1);
        e = DataList[row - 1];
        Assert.Equal(e.Int1, _cells[row, col++].Value);
        Assert.Equal(e.Int1 + e.Int2, _cells[row, col++].Value);
        Assert.Equal(null, _cells[row, col++].Value);
        Assert.Equal($"{e.Str1} -- {e.Date.Month}", _cells[row, col++].Value);
        Assert.Equal("AA " + DateTime.Now.Month, _cells[row, col++].Value);

        (row, col) = (3, 1);
        e = DataList[row - 1];
        Assert.Equal(e.Int1, _cells[row, col++].Value);
        Assert.Equal(e.Int1 + e.Int2, _cells[row, col++].Value);
        Assert.Equal(e.NullableInt + 1000, _cells[row, col++].Value);
        Assert.Equal($"{e.Str1} -- {e.Date.Month}", _cells[row, col++].Value);
        Assert.Equal("AA " + DateTime.Now.Month, _cells[row, col++].Value);

    }


    [Fact]
    public void AddSkipColumn() {
        var writer = new ExcelRangeWriter<EntityA>()
            .AddSkipColumn(1)
            .AddColumns((s, e) => s(e.Int1))
            .AddSkipColumn(1)
            .AddColumns((s, e) => s(e.Str1))
            .AddSkipColumn(1)
            .AddAction(it => Assert.Equal(5, it.iCol))
            .AddAction(it => it.SetColumnValue("E"))
            .AddColumn(it => it.SetColumnValue(it.Entity.Str1));

        var writeLmd = writer.Lambda();
        var fn = writeLmd.Compile();
        fn(_cells, DataList);

        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);

        for (var i = 0; i < DataList.Count;) {
            var e = DataList[i];
            var col = 1;
            i++;
            Assert.Equal(null, _cells[i, col++].Value);
            Assert.Equal(e.Int1, _cells[i, col++].Value);
            Assert.Equal(null, _cells[i, col++].Value);
            Assert.Equal(e.Str1, _cells[i, col++].Value);
            Assert.Equal("E", _cells[i, col++].Value);
            Assert.Equal(e.Str1, _cells[i, col++].Value);
        }

    }

    [Fact]
    public void AddActionWrite() {
        var writer = new ExcelRangeWriter<EntityA>()
            .AddAction(it => it.SetColumnValue(1111))
            .AddSkipColumn(2)
            .AddColumns((s, e) => s(e.Int1, e.Str1))
            .AddAction(it => it.SetColumnValue(it.Entity.Str1 + "-A"))
            .AddColumn(it => it.SetColumnValue(it.Entity.Str1))
            .AddAction(it => Assert.Equal(5, it.iCol));

        var writeLmd = writer.Lambda();
        var fn = writeLmd.Compile();
        fn(_cells, DataList);

        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);

        for (var i = 0; i < DataList.Count;) {
            var e = DataList[i];
            var col = 1;
            i++;
            Assert.Equal(1111, _cells[i, col++].Value);
            Assert.Equal(null, _cells[i, col++].Value);
            Assert.Equal(e.Int1, _cells[i, col++].Value);
            Assert.Equal(e.Str1 + "-A", _cells[i, col++].Value);
            Assert.Equal(e.Str1, _cells[i, col++].Value);
        }
    }

    [Fact]
    public void SetStart() {
        var (startRow, startCol) = (2, 2);
        var writer = new ExcelRangeWriter<EntityA>()
            .SetStart(startRow, startCol)
            .AddColumns((s, e) => s(e.Int1, e.Str1));

        var writeLmd = writer.Lambda();
        var fn = writeLmd.Compile();
        fn(_cells, DataList);

        Assert.Equal(startRow, _sheet.Dimension.Start.Row);
        Assert.Equal(DataList.Count, _sheet.Dimension.Rows);

        for (var i = 0; i < DataList.Count;) {
            var e = DataList[i];
            var col = 1;
            i += startRow;
            Assert.Equal(null, _cells[i, col++].Value);
            Assert.Equal(e.Int1, _cells[i, col++].Value);
            Assert.Equal(e.Str1, _cells[i, col++].Value);
        }
    }


}
