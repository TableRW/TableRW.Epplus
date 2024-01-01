global using System;
global using System.Collections.Generic;
global using System.Linq;
global using System.Reflection;
global using System.Linq.Expressions;

global using static GUsing;
static class GUsing {
    public static IEnumerable<int> Range(int start, int count) => Enumerable.Range(start, count);
}