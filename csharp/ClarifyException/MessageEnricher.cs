using System;
using System.Collections.Generic;
using System.Linq;

namespace codingdojo
{
    public class MessageEnricher
    {
        private List<IErrorValidator> _errorValidators;
        public MessageEnricher()
        {
            _errorValidators = new List<IErrorValidator>()
            {
                new ObjectReferenceNotSetErrorValidator(),
                new MissingFormulaErrorValidator(),
                new NoMatchesFoundErrorValidator(),
                new InvalidExpressionErrorValidator(),
                new CircularReferenceErrorValidator()
            };
        }

        public ErrorResult EnrichError(SpreadsheetWorkbook spreadsheetWorkbook, Exception e)
        {
            var formulaName = spreadsheetWorkbook.GetFormulaName();

            var error = _errorValidators
                .FirstOrDefault(ev => ev.Validate(e))
                ?.ErrorMessage(e, formulaName) 
                ?? e.Message;

            return new ErrorResult(formulaName, error, spreadsheetWorkbook.GetPresentation());

        }

    }
    public interface IErrorValidator
    {
        bool Validate(Exception e);
        string ErrorMessage(Exception e, string formulaName);
    }
    public class ObjectReferenceNotSetErrorValidator : IErrorValidator
    {
        public string ErrorMessage(Exception e, string formulaName)
            => "Missing Lookup Table";

        public bool Validate(Exception e)
            => e.Message.Equals("Object reference not set to an instance of an object") && e.StackTrace.Contains("VLookup");
    }
    public class MissingFormulaErrorValidator : IErrorValidator
    {
        public string ErrorMessage(Exception e, string formulaName)
            => $"Invalid expression found in tax formula [{formulaName}]. Check for merged cells near {((SpreadsheetException)e).Cells}";

        public bool Validate(Exception e)
            => e.Message.Equals("Missing Formula") && e.GetType() == typeof(SpreadsheetException);
    }
    public class NoMatchesFoundErrorValidator : IErrorValidator
    {
        public string ErrorMessage(Exception e, string formulaName)
            => $"No match found for token [{((SpreadsheetException)e).Token}] related to formula '{formulaName}'.";

        public bool Validate(Exception e)
            => e.Message.Equals("No matches found") && e.GetType() == typeof(SpreadsheetException);
    }
    public class InvalidExpressionErrorValidator : IErrorValidator
    {
        public string ErrorMessage(Exception e, string formulaName)
            => $"Invalid expression found in tax formula [{formulaName}]. Check that separators and delimiters use the English locale.";

        public bool Validate(Exception e)
            => e.GetType() == typeof(ExpressionParseException);
    }
    public class CircularReferenceErrorValidator : IErrorValidator
    {
        public string ErrorMessage(Exception e, string formulaName)
            => $"Circular Reference in spreadsheet related to formula '{formulaName}'. Cells: {((SpreadsheetException)e).Cells}";

        public bool Validate(Exception e)
            => e.Message.StartsWith("Circular Reference") && e.GetType() == typeof(SpreadsheetException);
    }
}