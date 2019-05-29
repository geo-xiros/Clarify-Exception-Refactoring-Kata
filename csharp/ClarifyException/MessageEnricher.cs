using System;

namespace codingdojo
{
    public class MessageEnricher
    {
        public ErrorResult EnrichError(SpreadsheetWorkbook spreadsheetWorkbook, Exception e)
        {
            var formulaName = spreadsheetWorkbook.GetFormulaName();
            var error = e.Message;

            if (InvalidExpression(e))
                error = $"Invalid expression found in tax formula [{formulaName}]. Check that separators and delimiters use the English locale.";
            else if (CircularReference(e))
                error = parseCircularReferenceException(e as SpreadsheetException, formulaName);
            else if (NoMatchesFound(e))
                error = parseNoMatchException(e as SpreadsheetException, formulaName);
            else if (MissingFormula(e))
                error = parseMissingFormula(e as SpreadsheetException, formulaName);
            else if (ObjectReferenceNotSet(e))
                error = "Missing Lookup Table";

            return new ErrorResult(formulaName, error, spreadsheetWorkbook.GetPresentation());
        }

        private static bool ObjectReferenceNotSet(Exception e) => "Object reference not set to an instance of an object".Equals(e.Message) && e.StackTrace.Contains("VLookup");
        private static bool MissingFormula(Exception e) => "Missing Formula".Equals(e.Message) && e.GetType() == typeof(SpreadsheetException);
        private static bool NoMatchesFound(Exception e) => "No matches found".Equals(e.Message) && e.GetType() == typeof(SpreadsheetException);
        private bool InvalidExpression(Exception e) => e.GetType() == typeof(ExpressionParseException);
        private bool CircularReference(Exception e) => e.Message.StartsWith("Circular Reference") && e.GetType() == typeof(SpreadsheetException);

        private string parseNoMatchException(SpreadsheetException e, string formulaName)
        {
            return $"No match found for token [{e.Token}] related to formula '{formulaName}'.";
        }

        private string parseCircularReferenceException(SpreadsheetException e, string formulaName)
        {
            return $"Circular Reference in spreadsheet related to formula '{formulaName}'. Cells: {e.Cells}";
        }

        private string parseMissingFormula(SpreadsheetException e, string formulaName)
        {
            return $"Invalid expression found in tax formula [{formulaName}]. Check for merged cells near {e.Cells}";
        }
    }
}