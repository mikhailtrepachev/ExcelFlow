using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFlow;

/// <summary>
/// Marks a class to be processed by the ExcelFlow Source Generator.
/// The generator will create a highly optimized, AOT-compatible mapper for this class.
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class ExcelFlowSerializableAttribute : Attribute
{
}
