using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using Antlr4.Runtime;
using ExcelButBetter.Grammar;

namespace ExcelButBetter.Logic
{
    public class GridManager
    {
        private readonly Dictionary<string, Cell> _cells = new();

        private string _savedState = "";

        public bool IsDirty
        {
            get
            {
                string currentState = GetCsvContent();
                return currentState != _savedState;
            }
        }

        // update state
        public void UpdateSavedState() { _savedState = GetCsvContent(); }

        public int RowCount { get; private set; } = 0;
        public int ColCount { get; private set; } = 0;
        private const string GenericError = "#ERROR!";

        public void SetDimensions(int rows, int cols)
        {
            RowCount = rows;
            ColCount = cols;

            if (string.IsNullOrEmpty(_savedState)) UpdateSavedState();
        }

        public void LoadFromCsvContent(string csv)
        {
            try
            {
                _cells.Clear();
                var lines = csv.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                if (lines.Length > 0)
                {
                    int newRowCount = lines.Length;
                    int newColCount = 0;

                    for (int r = 0; r < newRowCount; r++)
                    {
                        var fields = ParseCsvLine(lines[r]);
                        newColCount = Math.Max(newColCount, fields.Count);

                        for (int c = 0; c < fields.Count; c++)
                        {
                            if (!string.IsNullOrWhiteSpace(fields[c]))
                            {
                                string address = GetColumnName(c) + (r + 1);
                                var cell = GetOrCreateCell(address);
                                cell.Expression = fields[c];
                            }
                        }
                    }

                    RowCount = newRowCount;
                    ColCount = newColCount;
                }

                else
                {
                    // do nothing
                }

                RecalculateAll();
                UpdateSavedState(); 
            }
            catch (Exception ex) { throw new Exception($"Помилка читання: {ex.Message}"); }
        }

        // read
        public string GetCsvContent()
        {
            var sb = new StringBuilder();

            for (int row = 0; row < RowCount; row++)
            {
                var line = new List<string>();

                for (int col = 0; col < ColCount; col++)
                {
                    string address = GetColumnName(col) + (row + 1);
                    string content = _cells.ContainsKey(address) ? _cells[address].Expression : "";

                    if (content.Contains(",") || content.Contains("\"") || content.Contains("\n")) content = "\"" + content.Replace("\"", "\"\"") + "\"";
                    line.Add(content);
                }
                sb.AppendLine(string.Join(",", line));
            }
            return sb.ToString();
        }

        // should clean up
        private List<string> ParseCsvLine(string line) 
        { /*...*/ 
            var result = new List<string>(); 
            bool inQuotes = false; 
            StringBuilder currentField = new StringBuilder(); 

            for (int i = 0; i < line.Length; i++) 
            { 
                char c = line[i]; 

                if (inQuotes) 
                { 
                    if (c == '"') 
                    { 
                        if (i + 1 < line.Length && line[i + 1] == '"') 
                        { 
                            currentField.Append('"'); 
                            i++; 
                        } 

                        else inQuotes = false; 
                    } 

                    else currentField.Append(c); 
                } 

                else 
                { 
                    if (c == ',') 
                    { 
                        result.Add(currentField.ToString()); 
                        currentField.Clear(); 
                    } 

                    else if (c == '"') inQuotes = true; 

                    else currentField.Append(c); 
                } 
            } 

            result.Add(currentField.ToString()); 
            return result; 
        }

        public void CleanUpOutOfBounds() 
        { 
            var toRemove = _cells.Keys.Where(k => !IsCoordinateValid(k)).ToList(); 
            foreach (var key in toRemove) _cells.Remove(key); 
        }

        public IEnumerable<Cell> GetAllCells() => _cells.Values;

        public Cell GetOrCreateCell(string coordinate) 
        { 
            if (!_cells.ContainsKey(coordinate)) _cells[coordinate] = new Cell(coordinate.ToUpper());
            return _cells[coordinate]; 
        }

        public void UpdateCell(string coordinate, string expression) 
        { 
            if (!IsCoordinateValid(coordinate)) return; 

            var cell = GetOrCreateCell(coordinate); 

            cell.Expression = expression; 
            RecalculateAll(); 
        }

        public void RecalculateAll() 
        { 
            foreach (var cell in _cells.Values) 
            { 
                cell.TemporaryValue = null; 
                cell.IsEvaluating = false; 
                cell.Error = ""; 
            } 

            var cellsCopy = _cells.Values.ToList(); 
            
            foreach (var cell in cellsCopy) 
            { 
                if (cell.TemporaryValue == null) 
                { 
                    try { EvaluateCell(cell); } 
                    catch { } 
                } 
            } 

            foreach (var cell in _cells.Values) if (cell.TemporaryValue.HasValue) cell.Value = cell.TemporaryValue.Value; 
        }

        private double EvaluateCell(Cell cell) 
        { 
            if (cell.TemporaryValue.HasValue) return cell.TemporaryValue.Value; 

            if (cell.IsEvaluating) 
            { 
                cell.Error = GenericError; 
                throw new Exception("Cycle"); 
            } 

            cell.IsEvaluating = true; 

            try 
            { 
                string input = cell.Expression.Trim(); 
                
                if (string.IsNullOrEmpty(input)) 
                { 
                    cell.TemporaryValue = 0; 
                    cell.IsText = false; 
                } 

                else if (input.StartsWith("=")) 
                { 
                    cell.IsText = false; 
                    string formula = input.Substring(1); 
                    CalculateWithAntlr(cell, formula); 
                } 

                else 
                { 
                    if (double.TryParse(input, NumberStyles.Float, CultureInfo.InvariantCulture, out double val)) 
                    { 
                        cell.TemporaryValue = val; 
                        cell.IsText = false;
                    } 
                    else 
                    { 
                        cell.TemporaryValue = 0; 
                        cell.IsText = true; 
                    } 
                } 
            } 

            catch
            {
                cell.Error = GenericError; 
                throw new Exception("Error"); 
            } 

            finally { cell.IsEvaluating = false; }
            return cell.TemporaryValue ?? 0; 
        }

        private void CalculateWithAntlr(Cell cell, string formula) 
        { 
            var lexer = new LabCalculatorLexer(new AntlrInputStream(formula)); 
            lexer.RemoveErrorListeners(); 
            lexer.AddErrorListener(new LexerErrorListener()); 

            var parser = new LabCalculatorParser(new CommonTokenStream(lexer)); 
            parser.RemoveErrorListeners(); 
            parser.AddErrorListener(new ThrowExceptionErrorListener()); 

            var tree = parser.compileUnit(); 
            var visitor = new CalculatorVisitor((depCoord) => GetDependencyValue(depCoord)); 
            cell.TemporaryValue = visitor.Visit(tree); 
        }

        private double GetDependencyValue(string coordinate) 
        { 
            if (!IsCoordinateValid(coordinate)) throw new Exception("Ref Error"); 
            var cell = GetOrCreateCell(coordinate); 

            if (!string.IsNullOrEmpty(cell.Error)) throw new Exception("Dep Error"); 
            double val = EvaluateCell(cell); 

            if (cell.IsText) throw new Exception("Value Error"); 
            return val; 
        }

        private bool IsCoordinateValid(string coordinate) 
        { 
            var match = Regex.Match(coordinate.ToUpper(), @"^([A-Z]+)([0-9]+)$"); 

            if (!match.Success) return false; 

            string colPart = match.Groups[1].Value; 
            string rowPart = match.Groups[2].Value; 

            if (!int.TryParse(rowPart, out int row)) return false; 

            int col = GetColumnIndex(colPart) + 1; 

            return row <= RowCount && col <= ColCount; 
        }

        private string GetColumnName(int colIndex) 
        { 
            int divident = colIndex + 1; 
            string columnName = string.Empty; 

            while (divident > 0) 
            { 
                int letIndex = (divident - 1) % 26; 
                columnName = Convert.ToChar(65 + letIndex) + columnName; 
                divident = (int)((divident - letIndex) / 26); 
            } 

            return columnName; 
        }

        private int GetColumnIndex(string colName) 
        { 
            int index = 0; 

            foreach (char c in colName) 
            { 
                index *= 26; 
                index += (c - 'A' + 1); 
            } 

            return index - 1; 
        }
    }

    public class ThrowExceptionErrorListener : BaseErrorListener 
    { 
        public override void SyntaxError(TextWriter output, IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e) => throw new Exception("Syntax"); 
    }

    // not working without it
    public class LexerErrorListener : IAntlrErrorListener<int> 
    { 
        public void SyntaxError(TextWriter output, IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e) => throw new Exception("Syntax"); 
    }
}

//using System.Text;
//using System.Text.RegularExpressions;
//using System.Globalization;
//using Antlr4.Runtime;
//using ExcelButBetter.Grammar;

//namespace ExcelButBetter.Logic
//{
//    public class GridManager
//    {
//        private readonly Dictionary<string, Cell> _cells = new();

//        public int RowCount { get; private set; } = 0;
//        public int ColCount { get; private set; } = 0;

//        // Єдиний код помилки, як ти просив
//        private const string GenericError = "#ERROR!";

//        public void SetDimensions(int rows, int cols)
//        {
//            RowCount = rows;
//            ColCount = cols;
//        }

//        // --- CSV ---
//        public string GetCsvContent()
//        {
//            var sb = new StringBuilder();
//            for (int row = 0; row < RowCount; row++)
//            {
//                var line = new List<string>();
//                for (int col = 0; col < ColCount; col++)
//                {
//                    string address = GetColumnName(col) + (row + 1);
//                    string content = _cells.ContainsKey(address) ? _cells[address].Expression : "";
//                    if (content.Contains(",") || content.Contains("\"") || content.Contains("\n"))
//                    {
//                        content = "\"" + content.Replace("\"", "\"\"") + "\"";
//                    }
//                    line.Add(content);
//                }
//                sb.AppendLine(string.Join(",", line));
//            }
//            return sb.ToString();
//        }

//        public void LoadFromCsvContent(string csv)
//        {
//            try
//            {
//                _cells.Clear();
//                var lines = csv.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
//                if (lines.Length == 0) return;

//                int newRowCount = lines.Length;
//                int newColCount = 0;

//                for (int r = 0; r < newRowCount; r++)
//                {
//                    var fields = ParseCsvLine(lines[r]);
//                    newColCount = Math.Max(newColCount, fields.Count);
//                    for (int c = 0; c < fields.Count; c++)
//                    {
//                        if (!string.IsNullOrWhiteSpace(fields[c]))
//                        {
//                            string address = GetColumnName(c) + (r + 1);
//                            var cell = GetOrCreateCell(address);
//                            cell.Expression = fields[c];
//                        }
//                    }
//                }
//                RowCount = newRowCount;
//                ColCount = newColCount;
//                RecalculateAll();
//            }
//            catch
//            {
//                // Ігноруємо помилки завантаження, просто створюємо пусту або часткову
//            }
//        }

//        private List<string> ParseCsvLine(string line)
//        {
//            var result = new List<string>();
//            bool inQuotes = false;
//            StringBuilder currentField = new StringBuilder();
//            for (int i = 0; i < line.Length; i++)
//            {
//                char c = line[i];
//                if (inQuotes)
//                {
//                    if (c == '"')
//                    {
//                        if (i + 1 < line.Length && line[i + 1] == '"') { currentField.Append('"'); i++; }
//                        else inQuotes = false;
//                    }
//                    else currentField.Append(c);
//                }
//                else
//                {
//                    if (c == ',') { result.Add(currentField.ToString()); currentField.Clear(); }
//                    else if (c == '"') inQuotes = true;
//                    else currentField.Append(c);
//                }
//            }
//            result.Add(currentField.ToString());
//            return result;
//        }

//        // --- LOGIC ---

//        public void CleanUpOutOfBounds()
//        {
//            var toRemove = _cells.Keys.Where(k => !IsCoordinateValid(k)).ToList();
//            foreach (var key in toRemove) _cells.Remove(key);
//        }

//        public Cell GetOrCreateCell(string coordinate)
//        {
//            if (!_cells.ContainsKey(coordinate)) _cells[coordinate] = new Cell(coordinate.ToUpper());
//            return _cells[coordinate];
//        }

//        public void UpdateCell(string coordinate, string expression)
//        {
//            if (!IsCoordinateValid(coordinate)) return;
//            var cell = GetOrCreateCell(coordinate);
//            cell.Expression = expression;
//            RecalculateAll();
//        }

//        public void RecalculateAll()
//        {
//            // 1. Скидання
//            foreach (var cell in _cells.Values)
//            {
//                cell.TemporaryValue = null;
//                cell.IsEvaluating = false;
//                cell.Error = ""; // Очищуємо помилку перед перерахунком!
//            }

//            // 2. Обчислення
//            var cellsCopy = _cells.Values.ToList();
//            foreach (var cell in cellsCopy)
//            {
//                if (cell.TemporaryValue == null)
//                {
//                    try { EvaluateCell(cell); } catch { }
//                }
//            }

//            // 3. Фіксація
//            foreach (var cell in _cells.Values)
//            {
//                if (cell.TemporaryValue.HasValue)
//                {
//                    cell.Value = cell.TemporaryValue.Value;
//                }
//                // Якщо TemporaryValue null, то там вже записаний Error в EvaluateCell
//            }
//        }

//        private double EvaluateCell(Cell cell)
//        {
//            if (cell.TemporaryValue.HasValue) return cell.TemporaryValue.Value;
//            if (cell.IsEvaluating) { cell.Error = GenericError; throw new Exception("Cycle"); }

//            cell.IsEvaluating = true;
//            try
//            {
//                string input = cell.Expression.Trim();

//                if (string.IsNullOrEmpty(input))
//                {
//                    cell.TemporaryValue = 0;
//                    cell.IsText = false;
//                }
//                else if (input.StartsWith("="))
//                {
//                    // Це формула
//                    cell.IsText = false;
//                    string formula = input.Substring(1);
//                    CalculateWithAntlr(cell, formula);
//                }
//                else
//                {
//                    // Без дорівнює. 
//                    // Вимога: Тільки "чисте" число (без дужок) вважається числом.
//                    // (5) буде текстом.
//                    // NumberStyles.Float дозволяє десяткову крапку та експоненту, але НЕ дужки.
//                    if (double.TryParse(input, NumberStyles.Float, CultureInfo.InvariantCulture, out double val))
//                    {
//                        cell.TemporaryValue = val;
//                        cell.IsText = false;
//                    }
//                    else
//                    {
//                        // Все інше - текст
//                        cell.TemporaryValue = 0;
//                        cell.IsText = true;
//                    }
//                }
//            }
//            catch
//            {
//                // Будь-яка помилка (синтаксис, ділення на нуль, рекурсія) -> #ERROR!
//                cell.Error = GenericError;
//                throw new Exception("Error");
//            }
//            finally { cell.IsEvaluating = false; }

//            return cell.TemporaryValue ?? 0;
//        }

//        private void CalculateWithAntlr(Cell cell, string formula)
//        {
//            var lexer = new LabCalculatorLexer(new AntlrInputStream(formula));
//            lexer.RemoveErrorListeners(); lexer.AddErrorListener(new LexerErrorListener());
//            var parser = new LabCalculatorParser(new CommonTokenStream(lexer));
//            parser.RemoveErrorListeners(); parser.AddErrorListener(new ThrowExceptionErrorListener());
//            var tree = parser.compileUnit();
//            var visitor = new CalculatorVisitor((depCoord) => GetDependencyValue(depCoord));
//            cell.TemporaryValue = visitor.Visit(tree);
//        }

//        private double GetDependencyValue(string coordinate)
//        {
//            // Якщо координати не існує (видалили рядок) -> Помилка
//            if (!IsCoordinateValid(coordinate)) throw new Exception("Ref Error");

//            var cell = GetOrCreateCell(coordinate);

//            // Якщо в тій клітинці вже є помилка -> Поширюємо помилку
//            if (!string.IsNullOrEmpty(cell.Error)) throw new Exception("Dep Error");

//            // Рахуємо (якщо ще не пораховано)
//            double val = EvaluateCell(cell);

//            // Якщо там текст -> Помилка (не можна додавати текст)
//            if (cell.IsText) throw new Exception("Value Error");

//            return val;
//        }

//        private bool IsCoordinateValid(string coordinate)
//        {
//            var match = Regex.Match(coordinate.ToUpper(), @"^([A-Z]+)([0-9]+)$");
//            if (!match.Success) return false;
//            string colPart = match.Groups[1].Value;
//            string rowPart = match.Groups[2].Value;
//            if (!int.TryParse(rowPart, out int row)) return false;
//            int col = GetColumnIndex(colPart) + 1;
//            return row <= RowCount && col <= ColCount;
//        }

//        private string GetColumnName(int colIndex)
//        {
//            int dividend = colIndex + 1;
//            string columnName = string.Empty;
//            while (dividend > 0)
//            {
//                int modulo = (dividend - 1) % 26;
//                columnName = Convert.ToChar(65 + modulo) + columnName;
//                dividend = (int)((dividend - modulo) / 26);
//            }
//            return columnName;
//        }

//        private int GetColumnIndex(string colName)
//        {
//            int index = 0;
//            foreach (char c in colName) { index *= 26; index += (c - 'A' + 1); }
//            return index - 1;
//        }
//    }

//    public class ThrowExceptionErrorListener : BaseErrorListener
//    {
//        public override void SyntaxError(TextWriter output, IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e) => throw new Exception("Syntax");
//    }
//    public class LexerErrorListener : IAntlrErrorListener<int>
//    {
//        public void SyntaxError(TextWriter output, IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e) => throw new Exception("Syntax");
//    }
//}