using CommunityToolkit.Maui.Core;
using CommunityToolkit.Maui.Extensions;
using CommunityToolkit.Maui.Storage;
using CommunityToolkit.Maui.Views; // ?
using ExcelButBetter.Logic;
using System.Globalization;

namespace ExcelButBetter
{
    // start
    public partial class MainPage : ContentPage
    {
        int _rowCount = 5;
        int _colCount = 5;
        GridManager _gridManager;

        public static MainPage? Instance { get; private set; }

        public MainPage()
        {
            InitializeComponent();
            Instance = this;
            _gridManager = new GridManager();

            BtnSave.Text = "Зберегти";
            BtnLoad.Text = "Відкрити";
            BtnAddRow.Text = "+ Рядок";
            BtnDelRow.Text = "- Рядок";
            BtnAddCol.Text = "+ Стовп";
            BtnDelCol.Text = "- Стовп";
            BtnHelp.Text = "Довідка";

            UpdateGridDimensions();
            CreateGrid();

            _gridManager.UpdateSavedState();
        }

        // save (problems but ok)
        public async Task<bool> AskToSaveIfDirty()
        {
            if (!_gridManager.IsDirty) return true;

            var popup = new UnsavedChangesPopup();
            // var result = await this.ShowPopupAsync(popup); // may cause problems
            // if (result is UserChoice choice)

            await this.ShowPopupAsync(popup);

            switch (popup.Result)
            {
                case UserChoice.Save:
                    return await SaveFile();

                case UserChoice.Discard:
                    return true;

                case UserChoice.Cancel:
                default:
                    return false;
            }
        }

        private async Task<bool> SaveFile()
        {
            try
            {
                string defaultFileName = "table.csv";
                string csvContent = _gridManager.GetCsvContent();

                using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(csvContent));
                var fileSaverResult = await FileSaver.Default.SaveAsync(defaultFileName, stream, CancellationToken.None);

                if (fileSaverResult.IsSuccessful)
                {
                    await DisplayAlertAsync("Успіх", $"Збережено: {fileSaverResult.FilePath}", "ОК");
                    _gridManager.UpdateSavedState();
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                await DisplayAlertAsync("Помилка", $"Збій збереження: {ex.Message}", "ОК");
                return false;
            }
        }

        private async void SaveButton_Clicked(object? sender, EventArgs e) => await SaveFile();

        private async void LoadButton_Clicked(object? sender, EventArgs e)
        {
            if (!await AskToSaveIfDirty()) return;

            try
            {
                var fileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
                {
                    { DevicePlatform.WinUI, new[] { ".csv", ".txt", ".xlsm" } },
                    { DevicePlatform.MacCatalyst, new[] { "csv", "txt" } }
                });

                var result = await FilePicker.Default.PickAsync(new PickOptions { PickerTitle = "Оберіть файл", FileTypes = fileTypes });

                if (result != null)
                {
                    string csvContent = await File.ReadAllTextAsync(result.FullPath);
                    _gridManager.LoadFromCsvContent(csvContent);
                    _rowCount = _gridManager.RowCount;
                    _colCount = _gridManager.ColCount;
                    CreateGrid();
                    RefreshGrid();
                }
            }
            catch (Exception ex) { await DisplayAlertAsync("Помилка", $"Не вдалося відкрити: {ex.Message}", "ОК"); }
        }

        // should clean up
        private void UpdateGridDimensions() => _gridManager.SetDimensions(_rowCount, _colCount);

        private void CreateGrid()
        {
            GridWidget.Children.Clear(); 
            GridWidget.RowDefinitions.Clear(); 
            GridWidget.ColumnDefinitions.Clear();

            GridWidget.ColumnDefinitions.Add(new ColumnDefinition { Width = 50 });
            GridWidget.RowDefinitions.Add(new RowDefinition { Height = 30 });
            GridWidget.Add(new BoxView { Color = Colors.LightGray }, 0, 0);

            for (int col = 0; col < _colCount; col++) AddColumnHeader(col);
            for (int row = 0; row < _rowCount; row++) AddRow(row);
        }

        private void AddColumnHeader(int colIndex) 
        { 
            GridWidget.ColumnDefinitions.Add(new ColumnDefinition { Width = 100 }); 

            var label = new Label 
            { 
                Text = GetColumnName(colIndex), 
                HorizontalOptions = LayoutOptions.Center, 
                VerticalOptions = LayoutOptions.Center, 
                FontAttributes = FontAttributes.Bold, 
                TextColor = Colors.Black 
            }; 

            GridWidget.Add(new BoxView { Color = Colors.LightGray }, colIndex + 1, 0); 
            GridWidget.Add(label, colIndex + 1, 0); 
        }

        private void AddRow(int rowIndex) 
        { 
            GridWidget.RowDefinitions.Add(new RowDefinition { Height = 30 }); 

            var label = new Label 
            { 
                Text = (rowIndex + 1).ToString(), 
                HorizontalOptions = LayoutOptions.Center, 
                VerticalOptions = LayoutOptions.Center, 
                FontAttributes = FontAttributes.Bold, 
                TextColor = Colors.Black 
            }; 
            
            GridWidget.Add(new BoxView { Color = Colors.LightGray }, 0, rowIndex + 1); 
            GridWidget.Add(label, 0, rowIndex + 1); 

            for (int col = 0; col < _colCount; col++) CreateCellEntry(col, rowIndex); 
        }

        private void CreateCellEntry(int col, int row) 
        { 
            string address = GetColumnName(col) + (row + 1); 

            var entry = new Entry 
            { 
                HorizontalTextAlignment = TextAlignment.Start, 
                BackgroundColor = Colors.White, 
                TextColor = Colors.Black 
            }; 
            
            entry.BindingContext = address; 
            entry.Focused += Entry_Focused; 
            entry.Unfocused += Entry_Unfocused; 
            entry.Completed += Entry_Completed; 

            GridWidget.Add(entry, col + 1, row + 1); 
        }

        private void AddRow_Clicked(object sender, EventArgs e) 
        { 
            AddRow(_rowCount++); 
            UpdateGridDimensions();
            _gridManager.RecalculateAll();
            RefreshGrid();
        }

        private void AddCol_Clicked(object sender, EventArgs e) 
        { 
            AddColumnHeader(_colCount); 
            
            for (int row = 0; row < _rowCount; row++) CreateCellEntry(_colCount, row); 

            _colCount++; 
            UpdateGridDimensions();
            _gridManager.RecalculateAll();
            RefreshGrid();
        }

        private void DelRow_Clicked(object sender, EventArgs e) 
        { 
            if (_rowCount > 0) 
            { 
                _rowCount--; 
                GridWidget.RowDefinitions.RemoveAt(GridWidget.RowDefinitions.Count - 1); 

                var toRemove = GridWidget.Children.Where(c => Grid.GetRow((BindableObject)c) == _rowCount + 1).ToList(); 
                foreach (var el in toRemove) GridWidget.Children.Remove(el); 

                UpdateGridDimensions(); 
                _gridManager.CleanUpOutOfBounds(); 
                _gridManager.RecalculateAll(); 
                RefreshGrid(); 
            } 
        }

        private void DelCol_Clicked(object sender, EventArgs e) 
        { 
            if (_colCount > 0) 
            { 
                _colCount--; 
                GridWidget.ColumnDefinitions.RemoveAt(GridWidget.ColumnDefinitions.Count - 1); 
                
                var toRemove = GridWidget.Children.Where(c => Grid.GetColumn((BindableObject)c) == _colCount + 1).ToList(); 
                foreach (var el in toRemove) GridWidget.Children.Remove(el); 
                
                UpdateGridDimensions(); 
                _gridManager.CleanUpOutOfBounds(); 
                _gridManager.RecalculateAll(); 
                RefreshGrid(); 
            } 
        }

        private void Entry_Focused(object? sender, FocusEventArgs e) 
        { 
            if (sender is Entry entry) 
            { 
                string address = (string)entry.BindingContext; 
                var cell = _gridManager.GetOrCreateCell(address); 

                entry.Text = cell.Expression; 
                entry.TextColor = Colors.Black; 
            } 
        }

        private void Entry_Unfocused(object? sender, FocusEventArgs e) { if (sender is Entry entry) ProcessEntry(entry); }

        private void Entry_Completed(object? sender, EventArgs e) { if (sender is Entry entry) entry.Unfocus(); }

        private void ProcessEntry(Entry entry) 
        { 
            string address = (string)entry.BindingContext; 

            _gridManager.UpdateCell(address, entry.Text); 
            RefreshGrid(); 
        }

        private void RefreshGrid() 
        { 
            foreach (var child in GridWidget.Children) 
            { 
                if (child is Entry entry) 
                { 
                    if (entry.IsFocused) continue; 

                    string address = (string)entry.BindingContext; 
                    var cell = _gridManager.GetOrCreateCell(address); 

                    if (!string.IsNullOrEmpty(cell.Error)) 
                    { 
                        entry.Text = cell.Error; 
                        entry.TextColor = Colors.Red; 
                    } 

                    else 
                    { 
                        if (cell.IsText) entry.Text = cell.Expression;
                        else entry.Text = string.IsNullOrEmpty(cell.Expression) ? "" : cell.Value.ToString("G15", CultureInfo.InvariantCulture); // G15 may cause problems

                        entry.TextColor = Colors.Black; 
                    } 
                } 
            } 
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
        private async void HelpButton_Clicked(object? sender, EventArgs e) 
        { 
            await DisplayAlertAsync("Інфо",
                "Лабораторна робота #1. Варіант 35.\n" +
                "Виконав: Грицина Антон, група К-27.\n" +
                "\n" +
                "Операції:\n" +
                "1) +, -, *, / (бінарні).\n" +
                "3) +, - (унарні).\n" +
                "5) inc, dec.\n" +
                "8) =, <, >.\n" +
                "10) not.\n" +
                "\n" +
                "Для того, щоб використати формулу потрібно написати '=' на початку клітини.\n" +
                "Наприклад:\n" +
                "= not(3 > A2)\n" +
                "= inc(5)\n" +
                "= C6 * 7\n" +
                "Число 0 вважається логічним false, число 1 вважається логічним true.\n" +
                "Запис десяткових дробів можливий лише через крапку (.)."
            , "ОК"); 
        }
    }
}
