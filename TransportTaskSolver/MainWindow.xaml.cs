using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using ExcelDataReader;
using ClosedXML.Excel;

namespace TransportTaskSolver
{
    public partial class MainWindow : Window
    {
        private int[] _origSupplies;
        private int[] _origDemands;
        private int[,] _origCosts;

        private int[] supplies;
        private int[] demands;
        private int[,] costs;

        private bool _fictSupp;
        private bool _fictDem;

        private TextBox[,] _costBoxes;
        private TextBox[] _supBoxes;
        private TextBox[] _demBoxes;

        public MainWindow()
        {
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            BuildInputGrid();
        }

        private void BuildInputGrid()
        {
            if (CmbRows.SelectedItem == null || CmbCols.SelectedItem == null) return;

            int rows = int.Parse(((ComboBoxItem)CmbRows.SelectedItem).Content.ToString());
            int cols = int.Parse(((ComboBoxItem)CmbCols.SelectedItem).Content.ToString());

            GridInput.Children.Clear();
            GridInput.ColumnDefinitions.Clear();
            GridInput.RowDefinitions.Clear();

            
            GridInput.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });

            for (int c = 0; c < cols + 1; c++)
            {
                GridInput.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(70) });
            }

           
            for (int r = 0; r <= rows + 1; r++)
            {
                GridInput.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto, MinHeight = 35 });
            }

            _costBoxes = new TextBox[rows, cols];
            _supBoxes = new TextBox[rows];
            _demBoxes = new TextBox[cols];

           
            Action<string, int, int> addLabel = (text, r, c) =>
            {
                var tb = new TextBlock
                {
                    Text = text,
                    Margin = new Thickness(5),
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                    FontWeight = FontWeights.Bold
                };
                Grid.SetRow(tb, r);
                Grid.SetColumn(tb, c);
                GridInput.Children.Add(tb);
            };

         
            for (int j = 0; j < cols; j++)
                addLabel($"B{j + 1}", 0, j + 1);
            addLabel("Запас", 0, cols + 1); 

            
            for (int i = 0; i < rows; i++)
            {
                addLabel($"A{i + 1}", i + 1, 0);

                for (int j = 0; j < cols; j++)
                {
                    var tb = new TextBox
                    {
                        Width = 55,
                        Margin = new Thickness(2),
                        Text = "0",
                        HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center,
                        VerticalContentAlignment = System.Windows.VerticalAlignment.Center
                    };
                    Grid.SetRow(tb, i + 1);
                    Grid.SetColumn(tb, j + 1);
                    GridInput.Children.Add(tb);
                    _costBoxes[i, j] = tb;
                }

                var sup = new TextBox
                {
                    Width = 55,
                    Margin = new Thickness(2),
                    Text = "0",
                    HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center,
                    VerticalContentAlignment = System.Windows.VerticalAlignment.Center
                };
                Grid.SetRow(sup, i + 1);
                Grid.SetColumn(sup, cols + 1);
                GridInput.Children.Add(sup);
                _supBoxes[i] = sup;
            }

       
            addLabel("Потребность", rows + 1, 0);

            for (int j = 0; j < cols; j++)
            {
                var tb = new TextBox
                {
                    Width = 55,
                    Margin = new Thickness(2),
                    Text = "0",
                    HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center,
                    VerticalContentAlignment = System.Windows.VerticalAlignment.Center
                };
                Grid.SetRow(tb, rows + 1);
                Grid.SetColumn(tb, j + 1);
                GridInput.Children.Add(tb);
                _demBoxes[j] = tb;
            }
        }

        private void BtnBuildGrid_Click(object sender, RoutedEventArgs e)
        {
            BuildInputGrid();
            TxtStatus.Text = "Таблица построена под выбранный размер";
        }

        private bool LoadData()
        {
            try
            {
                int rows = _costBoxes.GetLength(0);
                int cols = _costBoxes.GetLength(1);

                var sup = new List<int>();
                for (int i = 0; i < rows; i++)
                    sup.Add(int.Parse(_supBoxes[i].Text));

                var dem = new List<int>();
                for (int j = 0; j < cols; j++)
                    dem.Add(int.Parse(_demBoxes[j].Text));

                var cost = new int[rows, cols];
                for (int i = 0; i < rows; i++)
                    for (int j = 0; j < cols; j++)
                        cost[i, j] = int.Parse(_costBoxes[i, j].Text);

                _origSupplies = sup.ToArray();
                _origDemands = dem.ToArray();
                _origCosts = cost;

                var supList = new List<int>(sup);
                var demList = new List<int>(dem);
                var costList = new List<List<int>>();
                for (int i = 0; i < rows; i++)
                {
                    var row = new List<int>();
                    for (int j = 0; j < cols; j++) row.Add(cost[i, j]);
                    costList.Add(row);
                }

                _fictSupp = false; _fictDem = false;
                int sumA = supList.Sum(), sumB = demList.Sum();
                if (sumA > sumB)
                {
                    demList.Add(sumA - sumB);
                    for (int i = 0; i < supList.Count; i++) costList[i].Add(0);
                    _fictDem = true;
                }
                else if (sumB > sumA)
                {
                    supList.Add(sumB - sumA);
                    var newRow = new List<int>();
                    for (int j = 0; j < demList.Count; j++) newRow.Add(0);
                    costList.Add(newRow);
                    _fictSupp = true;
                }

                supplies = supList.ToArray();
                demands = demList.ToArray();
                costs = new int[supplies.Length, demands.Length];
                for (int i = 0; i < supplies.Length; i++)
                    for (int j = 0; j < demands.Length; j++)
                        costs[i, j] = costList[i][j];

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка ввода: {ex.Message}", "Ошибка");
                return false;
            }
        }

        private void DisplayResult(int[,] plan, int[] currentDemands)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Поставщик");
            for (int j = 0; j < demands.Length; j++)
            {
                string name = $"B{j + 1}";
                if (_fictDem && j == demands.Length - 1) name += "*";
                dt.Columns.Add(name);
            }

            int totalCost = 0;
            for (int i = 0; i < supplies.Length; i++)
            {
                var row = dt.NewRow();
                string rowName = $"A{i + 1}";
                if (_fictSupp && i == supplies.Length - 1) rowName += "*";
                row["Поставщик"] = rowName;

                for (int j = 0; j < demands.Length; j++)
                {
                    string colName = $"B{j + 1}" + (_fictDem && j == demands.Length - 1 ? "*" : "");
                    row[colName] = plan[i, j].ToString();
                    totalCost += plan[i, j] * costs[i, j];
                }
                dt.Rows.Add(row);
            }

            var demandRow = dt.NewRow();
            demandRow["Поставщик"] = "Спрос";
            for (int j = 0; j < currentDemands.Length; j++)
            {
                string colName = $"B{j + 1}" + (_fictDem && j == demands.Length - 1 ? "*" : "");
                demandRow[colName] = currentDemands[j].ToString();
            }
            dt.Rows.Add(demandRow);

            ResultGrid.ItemsSource = dt.DefaultView;
            TxtTotalCost.Text = $"Общая стоимость: {totalCost}" +
                (_fictDem || _fictSupp ? " (с фиктивным элементом)" : "");
        }

        private void BtnNorthWest_Click(object sender, RoutedEventArgs e)
        {
            if (!LoadData()) return;
            int[,] plan = new int[supplies.Length, demands.Length];
            int[] ts = (int[])supplies.Clone();
            int[] td = (int[])demands.Clone();

            int i = 0, j = 0;
            while (i < supplies.Length && j < demands.Length)
            {
                int min = Math.Min(ts[i], td[j]);
                plan[i, j] = min;
                ts[i] -= min; td[j] -= min;
                if (ts[i] == 0) i++;
                if (td[j] == 0) j++;
            }
            TxtStatus.Text = "Выполнен расчёт: метод СЗУ";
            DisplayResult(plan, td);
        }

        private void BtnMinElement_Click(object sender, RoutedEventArgs e)
        {
            if (!LoadData()) return;
            var plan = BuildMinElementPlan();
            TxtStatus.Text = "Выполнен расчёт: метод минимальных элементов";
            DisplayResult(plan, demands);
        }

        private int[,] BuildMinElementPlan()
        {
            int[,] plan = new int[supplies.Length, demands.Length];
            int[] ts = (int[])supplies.Clone();
            int[] td = (int[])demands.Clone();

            var cells = new List<Tuple<int, int, int>>();
            for (int i = 0; i < supplies.Length; i++)
                for (int j = 0; j < demands.Length; j++)
                    cells.Add(Tuple.Create(costs[i, j], i, j));

            cells.Sort((a, b) => a.Item1.CompareTo(b.Item1));

            foreach (var cell in cells)
            {
                if (ts[cell.Item2] > 0 && td[cell.Item3] > 0)
                {
                    int min = Math.Min(ts[cell.Item2], td[cell.Item3]);
                    plan[cell.Item2, cell.Item3] = min;
                    ts[cell.Item2] -= min; td[cell.Item3] -= min;
                }
            }
            return plan;
        }

        private void BtnPotential_Click(object sender, RoutedEventArgs e)
        {
            if (!LoadData()) return;
            int[,] plan = BuildMinElementPlan();
            bool opt = false; int iter = 0;
            while (!opt && iter < 100)
            {
                iter++;
                opt = CheckOptimality(plan, out int br, out int bc);
                if (!opt)
                {
                    TxtStatus.Text = $"Итерация {iter}: улучшаем план...";
                    if (!ImprovePlan(plan, br, bc)) { TxtStatus.Text = "Цикл не найден."; break; }
                }
            }
            if (opt) TxtStatus.Text = $"План оптимален. Итераций: {iter}";
            DisplayResult(plan, demands);
        }

        private bool CheckOptimality(int[,] plan, out int bestRow, out int bestCol)
        {
            int m = supplies.Length, n = demands.Length;
            int[] u = new int[m], v = new int[n];
            bool[] uS = new bool[m], vS = new bool[n];
            for (int i = 0; i < m; i++) u[i] = int.MinValue;
            for (int j = 0; j < n; j++) v[j] = int.MinValue;
            u[0] = 0; uS[0] = true;

            bool ch = true;
            while (ch)
            {
                ch = false;
                for (int i = 0; i < m; i++)
                    for (int j = 0; j < n; j++)
                        if (plan[i, j] > 0)
                        {
                            if (uS[i] && !vS[j]) { v[j] = costs[i, j] - u[i]; vS[j] = true; ch = true; }
                            else if (vS[j] && !uS[i]) { u[i] = costs[i, j] - v[j]; uS[i] = true; ch = true; }
                        }
            }
            for (int i = 0; i < m; i++) if (!uS[i]) u[i] = 0;
            for (int j = 0; j < n; j++) if (!vS[j]) v[j] = 0;

            bestRow = -1; bestCol = -1; int maxD = 0;
            for (int i = 0; i < m; i++)
                for (int j = 0; j < n; j++)
                    if (plan[i, j] == 0)
                    {
                        int d = u[i] + v[j] - costs[i, j];
                        if (d > maxD) { maxD = d; bestRow = i; bestCol = j; }
                    }
            return maxD <= 0;
        }

        private bool ImprovePlan(int[,] plan, int sr, int sc)
        {
            int m = supplies.Length, n = demands.Length;
            int[,] signs = new int[m, n];
            var path = new List<Tuple<int, int>>();

            bool FindCycle(int r, int c, int depth, bool isRowMove)
            {
                if (depth >= 3)
                {
                    if (isRowMove && r == sr && c != sc) return true;
                    if (!isRowMove && c == sc && r != sr) return true;
                }
                if (isRowMove)
                {
                    for (int j = 0; j < n; j++)
                        if (j != c && signs[r, j] == 0 && plan[r, j] > 0)
                        {
                            signs[r, j] = (depth + 1) % 2 == 0 ? 1 : -1;
                            path.Add(Tuple.Create(r, j));
                            if (FindCycle(r, j, depth + 1, false)) return true;
                            signs[r, j] = 0; path.RemoveAt(path.Count - 1);
                        }
                }
                else
                {
                    for (int i = 0; i < m; i++)
                        if (i != r && signs[i, c] == 0 && plan[i, c] > 0)
                        {
                            signs[i, c] = (depth + 1) % 2 == 0 ? 1 : -1;
                            path.Add(Tuple.Create(i, c));
                            if (FindCycle(i, c, depth + 1, true)) return true;
                            signs[i, c] = 0; path.RemoveAt(path.Count - 1);
                        }
                }
                return false;
            }

            signs[sr, sc] = 1; path.Add(Tuple.Create(sr, sc));
            if (!FindCycle(sr, sc, 0, true) && !FindCycle(sr, sc, 0, false)) return false;

            int minVal = int.MaxValue;
            for (int i = 1; i < path.Count; i += 2)
                minVal = Math.Min(minVal, plan[path[i].Item1, path[i].Item2]);
            for (int i = 0; i < path.Count; i++)
                plan[path[i].Item1, path[i].Item2] += signs[path[i].Item1, path[i].Item2] * minVal;
            return true;
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            if (_costBoxes == null) return;
            int rows = _costBoxes.GetLength(0), cols = _costBoxes.GetLength(1);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++) _costBoxes[i, j].Text = "0";
                _supBoxes[i].Text = "0";
            }
            for (int j = 0; j < cols; j++) _demBoxes[j].Text = "0";

            ResultGrid.ItemsSource = null;
            TxtStatus.Text = "Данные очищены";
            TxtTotalCost.Text = "Общая стоимость: 0";
        }

        private void BtnLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm", Title = "Выберите файл" };

            if (dlg.ShowDialog() != true) return;
            try
            {
                using (var stream = File.Open(dlg.FileName, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var table = reader.AsDataSet().Tables[0];

                        int m = table.Rows.Count - 1;
                        int n = table.Columns.Count - 1;
                        if (m < 2 || m > 6 || n < 2 || n > 6)
                        {
                            MessageBox.Show("В файле должно быть от 2 до 6 поставщиков и потребителей.", "Ошибка");
                            return;
                        }

                        ((ComboBoxItem)CmbRows.Items[m - 2]).IsSelected = true;
                        ((ComboBoxItem)CmbCols.Items[n - 2]).IsSelected = true;
                        BuildInputGrid();

                        for (int i = 0; i < m; i++)
                        {
                            for (int j = 0; j < n; j++)
                                _costBoxes[i, j].Text = Convert.ToInt32(table.Rows[i][j]).ToString();
                            _supBoxes[i].Text = Convert.ToInt32(table.Rows[i][n]).ToString();
                        }
                        for (int j = 0; j < n; j++)
                            _demBoxes[j].Text = Convert.ToInt32(table.Rows[m][j]).ToString();

                        TxtStatus.Text = "Данные из Excel загружены в таблицу";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка чтения Excel: {ex.Message}", "Ошибка");
            }
        }

        private void BtnSaveExcel_Click(object sender, RoutedEventArgs e)
        {
            if (ResultGrid.ItemsSource == null)
            {
                MessageBox.Show("Сначала выполните расчёт.", "Внимание");
                return;
            }

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = "TransportTaskResult.xlsx"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Результат");
                    var dv = ResultGrid.ItemsSource as DataView;
                    if (dv == null) return;

                    var dt = dv.Table;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ws.Cell(1, j + 1).Value = dt.Columns[j].ColumnName;
                        ws.Cell(1, j + 1).Style.Font.Bold = true;
                        ws.Cell(1, j + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }
                    for (int i = 0; i < dt.Rows.Count; i++)
                        for (int j = 0; j < dt.Columns.Count; j++)
                            ws.Cell(i + 2, j + 1).Value = dt.Rows[i][j].ToString();

                    int last = dt.Rows.Count + 3;
                    ws.Cell(last, 1).Value = TxtTotalCost.Text;
                    ws.Cell(last, 1).Style.Font.Bold = true;
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(dlg.FileName);
                    MessageBox.Show("Файл сохранён.", "Успех");
                    TxtStatus.Text = "Результат сохранён";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}", "Ошибка");
            }
        }
    }
}