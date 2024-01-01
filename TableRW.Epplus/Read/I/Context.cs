using OfficeOpenXml;

namespace TableRW.Read.I.Epplus;

public class Context<TEntity>(ExcelWorksheet src)
: I.Context<ExcelWorksheet, TEntity>(src) { }

public class Context<TEntity, TData>(ExcelWorksheet src)
: I.Context<ExcelWorksheet, TEntity, TData>(src) { }
