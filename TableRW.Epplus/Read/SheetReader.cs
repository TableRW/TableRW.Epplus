using TableRW.Read.I.Epplus;

namespace TableRW.Read.Epplus;

public class SheetReader<TEntity> : SheetReaderImpl<Context<TEntity>> { }

public class SheetReader<TEntity, TData> : SheetReaderImpl<Context<TEntity, TData>> { }
