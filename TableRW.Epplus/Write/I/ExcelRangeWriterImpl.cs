using OfficeOpenXml;
using E = System.Linq.Expressions.Expression;

namespace TableRW.Write.I.Epplus;

public class ExcelRangeWriterImpl<C> : TableWriterImpl<C> {

    static ExcelRangeWriterImpl() {
        WriteSource<ExcelRange>.DefaultStart = (1, 1);
        WriteSource<ExcelRange>.Impl(WriteSrcValue);
    }

    public static Expression WriteSrcValue(Expression ctx, Expression value) {
        var convertVal = value.Type.IsValueType
            ? E.Convert(value, typeof(object))
            : value;

        var set = // ctx.Src[ctx.iRow,ctx.iCol].Value = (object)value
            E.Assign(E.Property(
                E.Call(E.Property(ctx, "Src"),
                "get_Item", [], E.Property(ctx, "iRow"), E.Property(ctx, "iCol")),
            "Value"), convertVal);
        return set;
    }

}
