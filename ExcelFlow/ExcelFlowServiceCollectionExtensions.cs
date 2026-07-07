namespace Microsoft.Extensions.DependencyInjection;

public static class ExcelFlowServiceCollectionExtensions
{
    /// <summary>
    /// Registers ExcelFlow services into the Dependency Injection container.
    /// </summary>
    public static IServiceCollection AddExcelFlow(this IServiceCollection services)
    {
        services.AddSingleton<global::ExcelFlow.IExcelFlowService, global::ExcelFlow.ExcelFlowService>();
        return services;
    }
}
