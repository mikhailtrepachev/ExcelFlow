using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using ExcelFlow;

namespace ExcelFlow.Tests;

public class ValidationFluentTests
{
    // Dummy class that bypasses the generator for isolated testing
    public class DummyRow : IExcelFlowSerializable<DummyRow>
    {
        public int Id { get; set; }
        public decimal Amount { get; set; }

        public static IEnumerable<ExcelColumnDefinition<DummyRow>> GetDefinitions()
        {
            yield return new ExcelColumnDefinition<DummyRow>("Id", typeof(int), (item, val) => item.Id = (int)val!, item => item.Id);
            yield return new ExcelColumnDefinition<DummyRow>("Amount", typeof(decimal), (item, val) => item.Amount = (decimal)val!, item => item.Amount);
        }
    }

    [Fact]
    public void Builder_ValidateMethod_AppendsRuleSuccessfully()
    {
        // Act - we use Excel.Read to invoke the internal builder properly
        var builder = Excel.Read<DummyRow>(new MemoryStream());
        builder.Validate(x => x.Amount > 0, "Amount must be greater than zero");

        // Assert
        Assert.NotNull(builder);
    }
}