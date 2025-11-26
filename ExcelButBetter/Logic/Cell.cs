namespace ExcelButBetter.Logic
{
    public class Cell
    {
        public string Coordinate { get; set; }
        public string Expression { get; set; }
        public double Value { get; set; }
        public string Error { get; set; } = "";
        
        public bool IsText { get; set; } = false;

        public double? TemporaryValue { get; set; }
        public bool IsEvaluating { get; set; }

        public Cell(string coordinate)
        {
            Coordinate = coordinate;
            Expression = "";
            Value = 0;
            Error = "";
            IsText = false;
        }
    }
}