using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFlow;

/// <summary>
/// Defines a contract for models that provide compile-time generated Excel column definitions.
/// Required for NativeAOT and Trimming compatibility.
/// </summary>
/// <typeparam name="T">The type of the model implementing this interface.</typeparam>
public interface IExcelFlowSerializable<T> where T : new()
{
    static abstract IEnumerable<ExcelColumnDefinition<T>> GetDefinitions();
}
