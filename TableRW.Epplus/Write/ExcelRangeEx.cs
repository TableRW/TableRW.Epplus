
using OfficeOpenXml;
using TableRW.Write.Epplus;
using TableRW.Utils.Ex;

namespace TableRW.Write;

public static class ExcelRangeEx {

    public static void WriteFrom<TEntity>(
        this ExcelRange tbl,
        IEnumerable<TEntity> enumerable,
        int cacheKey,
        Func<ExcelRangeWriter<TEntity>, Action<ExcelRange, IEnumerable<TEntity>>> buildWrite
    ) {
        if (CacheFn<TEntity>.DicFn is var dic && !dic.TryGetValue(cacheKey, out var fn)) {
            dic[cacheKey] = fn = buildWrite(new());
        }

        fn(tbl, enumerable);
    }

    public static void WriteFrom<TEntity, TData>(
        this ExcelRange tbl,
        IEnumerable<TEntity> enumerable,
        int cacheKey,
        Func<ExcelRangeWriter<TEntity, TData>, Action<ExcelRange, IEnumerable<TEntity>>> buildWrite
    ) {
        if (CacheFn<TEntity>.DicFn is var dic && !dic.TryGetValue(cacheKey, out var fn)) {
            dic[cacheKey] = fn = buildWrite(new());
        }

        fn(tbl, enumerable);
    }
}

static class CacheFn<T> {

    internal static Dictionary<int, Action<ExcelRange, IEnumerable<T>>> DicFn = new();
}
