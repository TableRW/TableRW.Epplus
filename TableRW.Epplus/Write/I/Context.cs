using OfficeOpenXml;

namespace TableRW.Write.I.Epplus {
    public class Context<TEntity>(ExcelRange src)
    : I.Context<ExcelRange, object?, TEntity>(src) { }

    public class Context<TEntity, TData>(ExcelRange src)
    : I.Context<ExcelRange, object?, TEntity, TData>(src) { }
}

namespace TableRW.Write.Epplus {
    public static class ContextEx {
        public static void SetColumnValue<TEntity>(
            this I.IContext<ExcelRange, object?, TEntity> it, object? value)
        => it.Src[it.iRow, it.iCol].Value = value;
    }
}