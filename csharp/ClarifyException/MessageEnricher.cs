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
            else if (e.Message.StartsWith("Circular Reference"))
            {
                error = parseCircularReferenceException(e, formulaName);
            }
            else if ("Object reference not set to an instance of an object".Equals(e.Message)
                && StackTraceContains(e, "VLookup"))
            {
                error = "Missing Lookup Table";
            }
            else if ("No matches found".Equals(e.Message))
            {
                error = parseNoMatchException(e, formulaName);
            }

            return new ErrorResult(formulaName, error, spreadsheetWorkbook.GetPresentation());
        }

        private bool StackTraceContains(Exception e, string message)
        {
            foreach (var ste in e.StackTrace.Split('\n'))
            {
                if (ste.Contains(message))
                    return true;
            }
            return false;
        }

        private string parseNoMatchException(Exception e, string formulaName)
        {
            if (e.GetType() == typeof(SpreadsheetException))
            {
                var we = (SpreadsheetException)e;
                return "No match found for token [" + we.Token + "] related to formula '" + formulaName + "'.";
            }

            return e.Message;
        }

        private string parseCircularReferenceException(Exception e, string formulaName)
        {
            if (e.GetType() == typeof(SpreadsheetException))
            {
                var we = (SpreadsheetException)e;
                return "Circular Reference in spreadsheet related to formula '" + formulaName + "'. Cells: " +
                       we.Cells;
            }

            return e.Message;
        }
    }
}