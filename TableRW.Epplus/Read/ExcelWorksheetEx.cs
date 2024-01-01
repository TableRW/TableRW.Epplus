
using OfficeOpenXml;
using TableRW.Read.Epplus;
using TableRW.Utils.Ex;

namespace TableRW.Read;

public static class ExcelWorksheetEx {
    public static List<TEntity> ReadToList<TEntity>(
        this ExcelWorksheet sheet, int headerRow
    ) where TEntity : new() {
        if (sheet == null) { throw new ArgumentNullException(nameof(sheet)); }
        if (sheet.Dimension == null) { throw new ArgumentNullException("sheet.Dimension == null"); }

        if (CacheReadFn<TEntity>.FnUseHeader is var fn && fn != null) {
            return fn(sheet);
        }

        var header = GetHeader();
        if (header.Count == 0) {
            throw new InvalidOperationException("No data read: The number of column headers is 0");
        }

        var (iCol, m0) = header[0];
        var reader = new SheetReader<TEntity>();
        reader.SetStart(headerRow + 1, iCol);
        reader.AddColumn(m0);

        foreach (var (i, m) in header.Skip(1)) {
            if (i - iCol > 1) {
                reader.AddSkipColumn(i - iCol - 1);
            }
            reader.AddColumn(m);
            iCol = i;
        }

        var readLmd = reader.Lambda();
        CacheReadFn<TEntity>.FnUseHeader = fn = readLmd.Compile();
        return fn(sheet);

        List<(int i, MemberInfo member)> GetHeader() {
            var t_entity = typeof(TEntity);
            var props = t_entity.GetProperties().Where(p => p.CanWrite)
                .Concat<MemberInfo>(t_entity.GetFields().Where(f => !f.IsInitOnly))
                .Where(m => m.HasAttribute<IgnoreReadAttribute>() == false)
                .ToDictionary(m => m.Name);

            return Range(1, sheet.Dimension.Columns + 1)
                .Select(i => (i, sheet.Cells[headerRow, i].Text))
                .Select((t) => (t.i, member: props.GetValueOr(t.Text, null!)))
                .Where(t => t.member != null)
                .ToList();
        }
    }

    public static List<TEntity> ReadToList<TEntity>(
        this ExcelWorksheet sheet,
        int cacheKey,
        Func<SheetReader<TEntity>, Func<ExcelWorksheet, List<TEntity>>> buildRead
    ) where TEntity : new() {
        if (CacheReadFn<TEntity>.DicFn is var dic && !dic.TryGetValue(cacheKey, out var fn)) {
            dic[cacheKey] = fn = buildRead(new());
        }

        return fn(sheet);
    }

    public static List<TEntity> ReadToList<TEntity, TData>(
        this ExcelWorksheet sheet,
        int cacheKey,
        Func<SheetReader<TEntity, TData>, Func<ExcelWorksheet, List<TEntity>>> buildRead
    ) where TEntity : new() {
        if (CacheReadFn<TEntity>.DicFn is var dic && !dic.TryGetValue(cacheKey, out var fn)) {
            dic[cacheKey] = fn = buildRead(new());
        }

        return fn(sheet);
    }

}

static class CacheReadFn<T> {
    internal static Func<ExcelWorksheet, List<T>>? FnUseHeader;

    internal static Dictionary<int, Func<ExcelWorksheet, List<T>>> DicFn = new();
}
