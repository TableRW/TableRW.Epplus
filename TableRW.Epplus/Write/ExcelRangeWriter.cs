
using TableRW.Write.I.Epplus;

namespace TableRW.Write.Epplus;

public class ExcelRangeWriter<TEntity> : ExcelRangeWriterImpl<Context<TEntity>> { }

public class ExcelRangeWriter<TEntity, TData> : ExcelRangeWriterImpl<Context<TEntity, TData>> { }
