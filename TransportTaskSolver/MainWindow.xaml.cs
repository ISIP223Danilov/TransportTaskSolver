using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TransportTaskSolver
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private int[] supplies;
        private int[] demands;
        private int[,] costs;

        public MainWindow()
        {
            InitializeComponent();
        }

        private bool LoadData()
        {
            try
            {
                supplies = TxtSupplies.Text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                          .Select(int.Parse).ToArray();
                demands = TxtDemands.Text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                         .Select(int.Parse).ToArray();

                string[] lines = TxtCosts.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length != supplies.Length)
                    throw new Exception("Количество строк тарифов не совпадает с количеством поставщиков.");

                costs = new int[supplies.Length, demands.Length];
                for (int i = 0; i < supplies.Length; i++)
                {
                    string[] values = lines[i].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (values.Length != demands.Length)
                        throw new Exception($"В строке {i + 1} неверное количество тарифов.");

                    for (int j = 0; j < demands.Length; j++)
                        costs[i, j] = int.Parse(values[j]);
                }

                if (supplies.Sum() != demands.Sum())
                {
                    MessageBox.Show("Задача открытая! Сумма запасов не равна сумме потребностей.", "Ошибка");
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка ввода: {ex.Message}", "Ошибка");
                return false;
            }
        }

        private void DisplayResult(int[,] plan, int[] currentSupplies, int[] currentDemands)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Поставщик");
            for (int j = 0; j < demands.Length; j++)
                dt.Columns.Add($"B{j + 1}");

            int totalCost = 0;

            for (int i = 0; i < supplies.Length; i++)
            {
                DataRow row = dt.NewRow();
                row["Поставщик"] = $"A{i + 1}";
                for (int j = 0; j < demands.Length; j++)
                {
                    if (plan[i, j] > 0)
                    {
                        row[$"B{j + 1}"] = $"{plan[i, j]} (x{costs[i, j]})";
                        totalCost += plan[i, j] * costs[i, j];
                    }
                    else
                    {
                        row[$"B{j + 1}"] = "0";
                    }
                }
                dt.Rows.Add(row);
            }

            DataRow demandRow = dt.NewRow();
            demandRow["Поставщик"] = "Спрос";
            for (int j = 0; j < currentDemands.Length; j++)
                demandRow[$"B{j + 1}"] = currentDemands[j] == 0 ? "0" : currentDemands[j].ToString();
            dt.Rows.Add(demandRow);

            ResultGrid.ItemsSource = dt.DefaultView;
            TxtTotalCost.Text = $"Общая стоимость: {totalCost}";
        }

        private void BtnNorthWest_Click(object sender, RoutedEventArgs e)
        {
            if (!LoadData()) return;
            int[,] plan = new int[supplies.Length, demands.Length];
            int[] tempSupplies = (int[])supplies.Clone();
            int[] tempDemands = (int[])demands.Clone();

            int i = 0, j = 0;
            while (i < supplies.Length && j < demands.Length)
            {
                int min = Math.Min(tempSupplies[i], tempDemands[j]);
                plan[i, j] = min;
                tempSupplies[i] -= min;
                tempDemands[j] -= min;
                if (tempSupplies[i] == 0) i++;
                if (tempDemands[j] == 0) j++;
            }
            TxtStatus.Text = "Выполнен расчет: Метод СЗУ";
            DisplayResult(plan, tempSupplies, tempDemands);
        }

        private void BtnMinElement_Click(object sender, RoutedEventArgs e)
        {
            if (!LoadData()) return;
            int[,] plan = BuildMinElementPlan();
            TxtStatus.Text = "Выполнен расчет: Метод мин. элементов";
            DisplayResult(plan, supplies, demands);
        }

        private int[,] BuildMinElementPlan()
        {
            int[,] plan = new int[supplies.Length, demands.Length];
            int[] tempSupplies = (int[])supplies.Clone();
            int[] tempDemands = (int[])demands.Clone();

            var cells = new List<(int cost, int row, int col)>();
            for (int i = 0; i < supplies.Length; i++)
                for (int j = 0; j < demands.Length; j++)
                    cells.Add((costs[i, j], i, j));

            cells.Sort((a, b) => a.cost.CompareTo(b.cost));

            foreach (var cell in cells)
            {
                if (tempSupplies[cell.row] > 0 && tempDemands[cell.col] > 0)
                {
                    int min = Math.Min(tempSupplies[cell.row], tempDemands[cell.col]);
                    plan[cell.row, cell.col] = min;
                    tempSupplies[cell.row] -= min;
                    tempDemands[cell.col] -= min;
                }
            }
            return plan;
        }

        // =========================================================
        // МЕТОД ПОТЕНЦИАЛОВ (ЖЕЛЕЗОБЕТОННЫЙ АЛГОРИТМ)
        // =========================================================

        private void BtnPotential_Click(object sender, RoutedEventArgs e)
        {
            if (!LoadData()) return;

            int[,] plan = BuildMinElementPlan();

            bool isOptimal = false;
            int iteration = 0;

            while (!isOptimal && iteration < 50)
            {
                iteration++;
                isOptimal = CheckOptimality(plan, out int bestRow, out int bestCol);

                if (!isOptimal)
                {
                    TxtStatus.Text = $"Итерация {iteration}: Улучшаем план...";
                    if (!ImprovePlan(plan, bestRow, bestCol))
                    {
                        TxtStatus.Text = "Ошибка: цикл не найден.";
                        break;
                    }
                }
            }

            if (isOptimal)
                TxtStatus.Text = $"План оптимален! Итераций: {iteration}";

            DisplayResult(plan, supplies, demands);
        }

        private bool CheckOptimality(int[,] plan, out int bestRow, out int bestCol)
        {
            int m = supplies.Length;
            int n = demands.Length;
            int[] u = new int[m];
            int[] v = new int[n];
            bool[] uSet = new bool[m];
            bool[] vSet = new bool[n];

            for (int i = 0; i < m; i++) u[i] = int.MinValue;
            for (int j = 0; j < n; j++) v[j] = int.MinValue;

            u[0] = 0;
            uSet[0] = true;

            bool changed = true;
            while (changed)
            {
                changed = false;
                for (int i = 0; i < m; i++)
                {
                    for (int j = 0; j < n; j++)
                    {
                        if (plan[i, j] > 0)
                        {
                            if (uSet[i] && !vSet[j])
                            {
                                v[j] = costs[i, j] - u[i];
                                vSet[j] = true;
                                changed = true;
                            }
                            else if (vSet[j] && !uSet[i])
                            {
                                u[i] = costs[i, j] - v[j];
                                uSet[i] = true;
                                changed = true;
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < m; i++) if (!uSet[i]) u[i] = 0;
            for (int j = 0; j < n; j++) if (!vSet[j]) v[j] = 0;

            bestRow = -1;
            bestCol = -1;
            int maxDelta = 0;

            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    if (plan[i, j] == 0)
                    {
                        int delta = u[i] + v[j] - costs[i, j];
                        if (delta > maxDelta)
                        {
                            maxDelta = delta;
                            bestRow = i;
                            bestCol = j;
                        }
                    }
                }
            }

            return maxDelta <= 0;
        }

        private bool ImprovePlan(int[,] plan, int startRow, int startCol)
        {
            int m = supplies.Length;
            int n = demands.Length;

            int[,] signs = new int[m, n];
            List<(int r, int c)> path = new List<(int r, int c)>();

            bool FindCycle(int r, int c, int depth, bool isRowMove)
            {
                // Если сделали минимум 3 шага и можем вернуться в старт по прямой — цикл замкнут
                if (depth >= 3)
                {
                    if (isRowMove && r == startRow && c != startCol)
                        return true;
                    if (!isRowMove && c == startCol && r != startRow)
                        return true;
                }

                if (isRowMove)
                {
                    for (int j = 0; j < n; j++)
                    {
                        if (j != c && signs[r, j] == 0 && plan[r, j] > 0)
                        {
                            signs[r, j] = (depth + 1) % 2 == 0 ? 1 : -1;
                            path.Add((r, j));

                            if (FindCycle(r, j, depth + 1, false)) return true;

                            signs[r, j] = 0;
                            path.RemoveAt(path.Count - 1);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < m; i++)
                    {
                        if (i != r && signs[i, c] == 0 && plan[i, c] > 0)
                        {
                            signs[i, c] = (depth + 1) % 2 == 0 ? 1 : -1;
                            path.Add((i, c));

                            if (FindCycle(i, c, depth + 1, true)) return true;

                            signs[i, c] = 0;
                            path.RemoveAt(path.Count - 1);
                        }
                    }
                }

                return false;
            }

            // Инициализируем стартовую клетку один раз
            signs[startRow, startCol] = 1;
            path.Add((startRow, startCol));

            if (!FindCycle(startRow, startCol, 0, true) && !FindCycle(startRow, startCol, 0, false))
                return false;

            // Находим минимум среди отрицательных вершин (нечётные индексы в path)
            int minVal = int.MaxValue;
            for (int i = 1; i < path.Count; i += 2)
            {
                minVal = Math.Min(minVal, plan[path[i].r, path[i].c]);
            }

            // Пересчитываем план
            for (int i = 0; i < path.Count; i++)
            {
                plan[path[i].r, path[i].c] += signs[path[i].r, path[i].c] * minVal;
            }

            return true;
        }
    }
}
