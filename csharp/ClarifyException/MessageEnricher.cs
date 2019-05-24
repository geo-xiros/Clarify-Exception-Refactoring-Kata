using System;

namespace codingdojo
{
    public class MessageEnricher
    {
        public ErrorResult EnrichError(SpreadsheetWorkbook spreadsheetWorkbook, Exception e)
        {
            var formulaName = spreadsheetWorkbook.GetFormulaName();
            var error = e.Message;

            if (e.GetType() == typeof(ExpressionParseException))
            {
                error = $"Invalid expression found in tax formula [{formulaName}]. Check that separators and delimiters use the English locale.";
            }
            else if (e.Message.StartsWith("Circular Reference") && e.GetType() == typeof(SpreadsheetException))
            {
                error = parseCircularReferenceException(e as SpreadsheetException, formulaName);
            }
            else if ("No matches found".Equals(e.Message) && e.GetType() == typeof(SpreadsheetException))
            {
                error = parseNoMatchException(e as SpreadsheetException, formulaName);
            }
            else if ("Missing Formula".Equals(e.Message) && e.GetType() == typeof(SpreadsheetException))
            {
                error = parseMissingFormula(e as SpreadsheetException, formulaName);
            }
            else if ("Object reference not set to an instance of an object".Equals(e.Message))
            {
                error = StackTraceContains(e, "VLookup");
            }
            
            return new ErrorResult(formulaName, error, spreadsheetWorkbook.GetPresentation());
        }

        private string StackTraceContains(Exception e, string message)
        {
            foreach (var ste in e.StackTrace.Split('\n'))
            {
                if (ste.Contains(message))
                    return "Missing Lookup Table";
            }
            return e.Message;
        }

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