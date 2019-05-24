using System;
using System.Collections.Generic;
using Xunit;

namespace codingdojo
{
    public class MessageEnricherTest
    {
        [Theory]
        [InlineData("No matches found")]
        [InlineData("Circular Reference xxx")]
        [InlineData("Missing Formula")]
        [InlineData("General Error")]
        public void SimpleExceptionShouldReturnMessage(string message)
        {
            var enricher = new MessageEnricher();

            var worksheet = new SpreadsheetWorkbook();

            var e = new Exception(message);

            var actual = enricher.EnrichError(worksheet, e);

            Assert.Equal(message, actual.Message);
        }
        [Fact]
        public void SpreadsheetExceptionShouldReturnNoMatchesFound()
        {
            var enricher = new MessageEnricher();

            var worksheet = new SpreadsheetWorkbook();

            var e = new SpreadsheetException("No matches found", new List<string>() { "C1", "C2", "C3" }, "token");

            var actual = enricher.EnrichError(worksheet, e);

            Assert.Equal("No match found for token [" + e.Token + "] related to formula '" + worksheet.GetFormulaName() + "'.", actual.Message);
        }
        [Fact]
        public void ShouldReturnCircularReference()
        {
            var enricher = new MessageEnricher();

            var worksheet = new SpreadsheetWorkbook();

            var e = new SpreadsheetException("Circular Reference xxx", new List<string>() { "C1", "C2", "C3" }, "token");

            var actual = enricher.EnrichError(worksheet, e);

            Assert.Equal("Circular Reference in spreadsheet related to formula '" + worksheet.GetFormulaName() + "'. Cells: " + e.Cells, actual.Message);
        }
        [Fact]
        public void ShouldReturnInvalidExpression()
        {
            var enricher = new MessageEnricher();

            var worksheet = new SpreadsheetWorkbook();

            var e = new ExpressionParseException("Object reference not set to an instance of an object");

            var actual = enricher.EnrichError(worksheet, e);

            Assert.Equal("Invalid expression found in tax formula [" + worksheet.GetFormulaName() + "]. Check that separators and delimiters use the English locale.", actual.Message);
        }
    }
}