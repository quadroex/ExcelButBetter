using Antlr4.Runtime.Misc;
using ExcelButBetter.Grammar;

namespace ExcelButBetter.Logic
{
    public class CalculatorVisitor : LabCalculatorBaseVisitor<double>
    {
        private readonly Func<string, double> _getCellValue;

        public CalculatorVisitor(Func<string, double> getCellValue) { _getCellValue = getCellValue; }

        private double ProcessResult(double value)
        {
            if (double.IsInfinity(value)) throw new DivideByZeroException("OVERFLOW");
            if (double.IsNaN(value)) throw new DivideByZeroException("NAN");

            if (value == 0.0) return 0.0;

            return Math.Round(value, 10);
        }

        public override double VisitCompileUnit([NotNull] LabCalculatorParser.CompileUnitContext context) { return Visit(context.expression()); }

        public override double VisitNumberExpr([NotNull] LabCalculatorParser.NumberExprContext context)
        {
            var text = context.GetText();
            return double.Parse(text, System.Globalization.CultureInfo.InvariantCulture);
        }

        public override double VisitIdentifierExpr([NotNull] LabCalculatorParser.IdentifierExprContext context)
        {
            var coordinate = context.GetText();
            return _getCellValue(coordinate);
        }

        public override double VisitParenthesizedExpr([NotNull] LabCalculatorParser.ParenthesizedExprContext context) { return Visit(context.expression()); }

        public override double VisitIncExpr([NotNull] LabCalculatorParser.IncExprContext context) => ProcessResult(Visit(context.expression()) + 1);

        public override double VisitDecExpr([NotNull] LabCalculatorParser.DecExprContext context) => ProcessResult(Visit(context.expression()) - 1);

        public override double VisitNotExpr([NotNull] LabCalculatorParser.NotExprContext context)
        {
            double value = Visit(context.expression());
            return value != 0 ? 0.0 : 1.0;
        }

        public override double VisitUnaryPlusExpr([NotNull] LabCalculatorParser.UnaryPlusExprContext context) { return Visit(context.expression()); }

        public override double VisitUnaryMinusExpr([NotNull] LabCalculatorParser.UnaryMinusExprContext context)
        {
            double val = Visit(context.expression());
            if (val == 0.0) return 0.0;
            return ProcessResult(-val);
        }

        public override double VisitMultiplicativeExpr([NotNull] LabCalculatorParser.MultiplicativeExprContext context)
        {
            var left = Visit(context.expression(0));
            var right = Visit(context.expression(1));

            if (context.MULTIPLY() != null) return ProcessResult(left * right);
            if (right == 0) throw new DivideByZeroException("DIV BY ZERO");

            return ProcessResult(left / right);
        }

        public override double VisitAdditiveExpr([NotNull] LabCalculatorParser.AdditiveExprContext context)
        {
            var left = Visit(context.expression(0));
            var right = Visit(context.expression(1));

            if (context.PLUS() != null) return ProcessResult(left + right);
            return ProcessResult(left - right);
        }

        public override double VisitRelationalExpr([NotNull] LabCalculatorParser.RelationalExprContext context)
        {
            var left = Visit(context.expression(0));
            var right = Visit(context.expression(1));
            bool result = false;

            if (context.GT() != null) result = left > right;
            else if (context.LT() != null) result = left < right;
            else if (context.EQ() != null) result = Math.Abs(left - right) < 1e-9;

            return result ? 1.0 : 0.0;
        }
    }
}
