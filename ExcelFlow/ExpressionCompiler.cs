using System.Linq.Expressions;
using System.Reflection;

namespace ExcelFlow;

public static class ExpressionCompiler
{
    public static Action<T, object> CompileSetter<T>(PropertyInfo propertyInfo)
    {
        ParameterExpression instanceParam = Expression.Parameter(typeof(T), "instance");
        ParameterExpression valueParam = Expression.Parameter(typeof(object), "value");
        
        UnaryExpression convertedValue = Expression.Convert(valueParam, propertyInfo.PropertyType);
        
        MethodCallExpression setterCall = Expression.Call(instanceParam, propertyInfo.GetSetMethod(true)!, convertedValue);
        
        Expression<Action<T, object>> lambda = Expression.Lambda<Action<T, object>>(setterCall, instanceParam, valueParam);
        
        return lambda.Compile();
    }

    public static Func<T, object> CompileGetter<T>(PropertyInfo property)
    {
        ParameterExpression instance = Expression.Parameter(typeof(T), "instance");
        
        MemberExpression propertyAccess = Expression.Property(instance, property);
        
        UnaryExpression convertedProperty = Expression.Convert(propertyAccess, typeof(object));
        
        Expression<Func<T, object>> lambda = Expression.Lambda<Func<T, object>>(convertedProperty, instance);
        
        return lambda.Compile();
    }
}