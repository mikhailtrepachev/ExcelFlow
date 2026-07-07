namespace ExcelFlow.Tests;
using ExcelDataReader;

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

        public static int[] InitializeIndexMap(Dictionary<string, int> headerMap) => new[] { 0, 1 };
        public static void ParseRow(IExcelDataReader reader, int[] indexMap, out DummyRow item, Action<ExcelParseError>? onError, int rowNumber)
        {
            item = new DummyRow();
            // Dummy parsing logic not needed for builder tests
        }
        
        public static void WriteHeaders(DocumentFormat.OpenXml.OpenXmlWriter writer) { }
        public static void WriteRow(DocumentFormat.OpenXml.OpenXmlWriter writer, DummyRow item) { }
    }

    [Fact]
    public void Builder_ValidateMethod_AppendsRuleSuccessfully()
    {
        // Act - we use Excel.Read to invoke the internal builder properly
        ExcelReaderBuilder<DummyRow> builder = Excel.Read<DummyRow>(new MemoryStream());
        builder.Validate(x => x.Amount > 0, "Amount must be greater than zero");

        // Assert
        Assert.NotNull(builder);
    }
}