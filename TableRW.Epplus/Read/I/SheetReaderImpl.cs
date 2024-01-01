using OfficeOpenXml;
using E = System.Linq.Expressions.Expression;

namespace TableRW.Read.I.Epplus;

public class SheetReaderImpl<C> : TableReaderImpl<C> {

    static SheetReaderImpl() {
        ReadSource<ExcelWorksheet>.SetDefaultStart(1, 1);
        ReadSource<ExcelWorksheet>.Impl(
            ReadSrcValueByIndex,
            (src, iRow) => iRow > src.Dimension.End.Row,
            (src, iRow, iCol) => src.Cells[iRow, iCol].Value == null);
    }

    public static Expression ReadSrcValueByIndex(Expression ctx, Type valueType)
    => ConvertSrcValue(GetSrcValueByIndex(ctx), valueType);

    public static Expression GetSrcValueByIndex(Expression ctx)
    => Utils.Expr.ExtractBody((ISource<ExcelWorksheet> ctx) =>
       ctx.Src.Cells[ctx.iRow, ctx.iCol], ctx);

    public static Expression ConvertSrcValue(Expression srcValue, Type valueType)
    => valueType == typeof(string)
       ? E.Property(srcValue, nameof(ExcelRange.Text))
       : E.Call(srcValue, nameof(ExcelRange.GetValue), [valueType]);

}
