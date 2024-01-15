# TableRW.Epplus (English | [中文](Doc/README.zh-CN.md))
[![NuGet Version](https://img.shields.io/nuget/v/TableRW.Epplus.svg?label=NuGet)](https://www.nuget.org/packages/TableRW.Epplus)

A library for reading and writing table data, using expression trees to generate delegates (Lambda), quickly and conveniently reading and writing data to entity objects (Entity), and mapping multi-layer entities to read and write.

```
dotnet add package TableRW.Epplus
```

## Read from `Worksheets` to Entity

### Add namespace
``` cs
using TableRW.Read;
using TableRW.Read.Epplus;
```

### Simple reading (not cached)
```cs
public class Entity {
    public long Id { get; set; }
    public string Name;
    public string Tel; // it can be of a field
    public int? NullableInt { get; set; } // or a property
}

var reader = new SheetReader<Entity>()
    .AddColumns((s, e) => s(e.Id, e.Name, e.Tel, e.NullableInt));

// When debugging, you can view the generated expression tree
var readLmd = reader.Lambda(); // Expression<Func<ExcelWorksheet, List<Entity>>>
var readFn = readLmd.Compile(); // Func<ExcelWorksheet, List<Entity>>
var list = readFn(excelWorksheet); // List<Entity>
```

### Reading with subtables (not cached)
``` cs
public class EntityA {
    public int Id { get; set; }
    public string Name { get; set; }
    public DateTime Date { get; set; }
    public List<EntityB> SubB { get; set; }
}
public class EntityB {
    public int Id { get; set; }
    public string Text { get; set; }
    public string Remark { get; set; }
}
var reader2 = new SheetReader<EntityA>()
    .AddColumns((s, e) => s(s.RowKey(e.Id), e.Name, e.Date))
    .AddSubTable(e => e.SubList, (s, e) => s(e.Id, e.Text, e.Remark));

var readLmd = reader2.Lambda(); // Expression<Func<ExcelWorksheet, List<EntityA>>>
var readFn = readLmd.Compile(); // Func<ExcelWorksheet, List<EntityA>>

// excelWorksheet
// | 10  | name1 | 101  | text101 | remark101
// | 10  | name1 | 102  | text102 | remark102
// | 20  | name2 | 201  | text201 | remark201
var list = readFn(excelWorksheet); // List<EntityA>
_ = list.Count == 2;
_ = list[0].SubB.Count == 2;
_ = list[1].SubB.Count == 1;

```

### Cache Generated delegate
The above `reader` compiles the expression tree every time it is executed, and should actually cache the resulting `readFn` and call the delegate directly afterwards.
``` cs
// The user needs to create a new class to manage the Cache.
static class CacheReadFn<T> {
    internal static Func<ExcelWorksheet, List<T>>? Fn;
}

static class CacheReadSheet {
    public static List<T> Read<T>(ExcelWorksheet sheet, Action<SheetReader<T>> buildRead) {
        if (CacheReadFn<T>.Fn == null) {
            var reader = new SheetReader<T>();
            buildRead(reader);

            // When debugging, you can view the generated expression tree
            var readLmd = reader.Lambda();
            CacheReadFn<T>.Fn = readLmd.Compile();
        }
        return CacheReadFn<T>.Fn(sheet);
    }
}

var list = CacheReadSheet.Read<Entity>(sheet, reader => {
    reader.AddColumns((s, e) => s(e.Id, e.Name, e.Tel, e.NullableInt));
});
```

### Use the cache provided by the library
This library also has some simple encapsulation for user-friendly invocation:
``` cs
using TableRW.Read;
using TableRW.Read.Epplus;

void Example(ExcelWorksheet sheet) {
    // Use the column name of the ExcelWorksheet as the property mapping.
    // The column name and property name must be the same.
    var list1 = sheet.ReadToList<Entity>(headerRow: 1); // List<Entity>

    var list2 = sheet.ReadToList<Entity>(cacheKey: 0, reader => {
        // Handle the mapping of properties and columns yourself
        reader.AddColumns((s, e) => s(e.Id, e.Name, e.Tel, e.NullableInt));

        // When debugging, you can view the generated expression tree
        var lmd = reader.Lambda();
        return lmd.Compile();
    });
}
```

### Events on read
``` cs
static void Example(ExcelWorksheet sheet) {
var list2 = sheet.ReadToList<Entity>(cacheKey: 0, reader => {
    reader.AddColumns((s, e) =>
        s(e.Id, e.Name, e.Tel, e.NullableInt))
        .OnStartReadingTable(it => {
            // Return false ends the read
            return true;
        })
        .OnStartReadingRow(it => {
            // If the SkipRow parameter is true, the row will be skipped.
            return it.SkipRow(true);
        })
        .OnEndReadingRow(it => {
            // If the SkipRow parameter is true, the row will be skipped.
            return it.SkipRow(true);
        })
        .OnEndReadingTable(it => { });

    var lmd = reader.Lambda();
    return lmd.Compile();
});
}
```

### Adjust the generated Lambda
``` cs
var reader = new SheetReader<Entity>()
    .AddColumns((s, e) => s(e.Id, e.Name, e.Tel, e.NullableInt));

// When debugging, you can view the generated expression tree
var lmd1 = reader.Lambda();
var fn1 = lmd1.Compile(); // Func<ExcelWorksheet, List<Entity>>
fn1(sheet);


var lmd2 = reader.Lambda(f => f.StartRow());
var fn2 = lmd2.Compile(); // Func<ExcelWorksheet, int, List<Entity>>
var startRow = 3; // Start reading from row 3
fn2(sheet, startRow);


var lmd3 = reader.Lambda(f => f.Start());
var fn3 = lmd3.Compile(); // Func<ExcelWorksheet, int, int, List<Entity>>
(startRow, var startCol) = (3, 2); // Start reading from row 3, column 2
fn3(sheet, startRow, startCol);

var lmd4 = reader.Lambda(f => f.ToDictionary(entity => entity.Id));
var fn4 = lmd4.Compile(); // Func<ExcelWorksheet, Dictionary<long, Entity>>
// Returns a Dictionary with entity.Id as key
var dic4 = fn4(sheet); // Dictionary<long, Entity>

// multiple combinations
var lmd5 = reader.Lambda(f => f.StartRow().ToDictionary(entity => entity.Id));
var fn5 = lmd5.Compile(); // Func<ExcelWorksheet, int, int, Dictionary<long, Entity>>
startRow = 2;
var dic5 = fn5(sheet, startRow);
```

### More ways to read
``` cs
static void Example(ExcelWorksheet sheet) {
var list = sheet.ReadToList<Entity>(cacheKey: 0, reader => {
    var x = reader
        // Set the starting position to read
        .SetStart(row: 3, column: 2)
        // Add several column mapping reads
        .AddColumns((s, e) => s(e.Id, e.Name))
        // Skip 2 columns to read
        .AddSkipColumn(2)
        // Convert the value of this column to DateTime, and then execute a function
        .AddColumnRead((DateTime val) => it => {
            if (val.Year < 2000) {
                // If Year < 2000, skip reading this row
                return it.SkipRow();
            }
            it.Entity.Year = val.Year;
            return null; // No action to be done
        })
        //Add a few more columns to read
        .AddColumns((s, e) => s(e.Text1, e.Text2))
        // Execute an Action. There is no data column read here, and the entity can be processed.
        .AddActionRead(it => {
            it.Entity.Remark1 = it.Entity.Text1 + it.Entity.Text2;
            it.Entity.Remark2 = it.Entity.Id + " - " + it.Entity.Year;
        });


    var lmd = reader.Lambda();
    return lmd.Compile();
});
}
```

## Write `ExcelRange`

### Add namespace
``` cs
using TableRW.Write;
using TableRW.Write.Epplus;
```

### Simple write (not cached)
``` cs
public class Entity {
    public long Id { get; set; }
    public string Name;
    public string Tel; // it can be of a field
    public int? NullableInt { get; set; } // or a property
}

var writer = new ExcelRangeWriter<Entity>()
    .AddColumns((s, e) => s(e.Id, s.Skip(1), e.Name, e.Tel, e.NullableInt));

// When debugging, you can view the generated expression tree
var writeLmd = writer.Lambda(); // Expression<Action<ExcelRange, IEnumerable<Entity>>>
var writeFn = writeLmd.Compile(); // Action<ExcelRange, IEnumerable<Entity>>
IEnumerable<Entity> data = new List<Entity>();
writeFn(cells, data);
```

### Cache Generated delegate
The above `writer` compiles the expression tree for each execution, and should actually cache the resulting `writeFn`, and call the delegate directly thereafter.
``` cs
// The user needs to create a new class to manage the Cache.
static class CacheWriteFn<T> {
    internal static Action<ExcelRange, IEnumerable<T>>? Fn;
}

static class CacheWriteExcel {
    public static void WriteFrom<T>(
        ExcelRange cells, IEnumerable<TEntity> data, Action<ExcelRangeWriter<T>> buildWrite
    ) {
        if (CacheWriteFn<T>.Fn == null) {
            var writer = new ExcelRangeWriter<T>();
            buildWrite(writer);

            // When debugging, you can view the generated expression tree
            var writeLmd = writer.Lambda();
            CacheWriteFn<T>.Fn = writeLmd.Compile();
        }
        CacheWriteFn<T>.Fn(cells);
    }
}

var list = new List<Entity>();
CacheWriteExcel.WriteFrom<Entity>(cells, list, writer => {
    writer.AddColumns((s, e) => s(e.Id, e.Name, e.Tel, e.NullableInt));
});
```

### Use the cache provided by the library
This library also has some simple encapsulation for user-friendly invocation:
``` cs
using TableRW.Write;
using TableRW.Write.Epplus;

void Example(ExcelRange cells, List<Entity> data) {
    cells.WriteFrom(data, cacheKey: 0, writer => {
        writer.AddColumns((s, e) => s(e.Id, e.Name, e.Tel, e.NullableInt));

        // When debugging, you can view the generated expression tree
        var lmd = writer.Lambda();
        return lmd.Compile();
    );
    // Data is written to cells
}
```

### Events on write
``` cs
static void Example(ExcelRange cells, List<Entity> data) {

cells.WriteFrom(data, cacheKey: 0, writer => {
    writer.AddColumns((s, e) =>
        s(e.Id, e.Name, e.Tel, e.NullableInt))
        .OnStartWritingTable(it => {
            it.Src[1, 1].Value = "set column";
        })
        .OnStartWritingRow(it => {
            it.Src[1, 1].Value = "set column";
        })
        .OnEndWritingRow(it => {
            it.Src[1, 1].Value = "set column";
        })
        .OnEndWritingTable(it => { });

    var lmd = writer.Lambda();
    return lmd.Compile();
});
}
```

