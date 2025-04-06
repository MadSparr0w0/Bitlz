using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PriceCategory
{
    public partial class MainForm: Form
    {
        public MainForm()
        {
            InitializeComponent();
            Fonts();
        }
        PrivateFontCollection font;

        private double[] input1D;
        private double[,] input2D;
        private int days;
        private string file;

        public static string analysis;

        private void SolveB_Click(object sender, EventArgs e)
        {
            MaxVTForm vt = new MaxVTForm(categoryTB, reasonTB, dataInfo1TB, dataInfo2TB, dataInfo3TB, data1TB, data2TB, data3TB, exportB, file);
            vt.Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetupVerticalInput();
            tooltipPattern.SetToolTip(ChooseFileB, "Пример заполнения для импорта");
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            HistoryForm history = new HistoryForm();
            history.Show();
        }

        private void RefreshTable()
        {
            monthGrid.Rows.Clear();
            if (verticalRB.Checked)
            {
                monthGrid.Rows.Add(days * 24);
                int k = 0;
                int g = 1;
                for (int i = 0; i < days * 24; i++)
                {
                    if (k < 24)
                    {
                        monthGrid.Rows[i].Cells[0].Value = g;
                        k++;
                    }
                    else
                    {
                        k = 0;
                        g++;
                        monthGrid.Rows[i].Cells[0].Value = g;
                        k++;
                    }
                    monthGrid.Rows[i].Cells[1].Value = $"{k - 1}:00";
                }
            }
            else
            {
                monthGrid.Rows.Add(days);
                for (int i = 0; i < days; i++)
                {
                    monthGrid.Rows[i].Cells[0].Value = i + 1;
                }
            }
        }

        private void AddRecord_Click(object sender, EventArgs e)
        {
            monthGrid.Visible = true;
            verticalRB.Visible = true;
            horizontalRB.Visible = true;
            AddRowsB.Visible = true;
            RowsCB.Visible = true;
        }

        private void AddRowsB_Click(object sender, EventArgs e)
        {
            if (RowsCB.SelectedIndex != -1)
            {
                RefreshTable();
            }
            else
            {
                MessageBox.Show("Выберите количество дней!", "Ошибка");
            }
        }

        private void SetupVerticalInput()
        {
            monthGrid.Columns.Clear();
            monthGrid.Rows.Clear();

            var dayColumn = new DataGridViewTextBoxColumn
            {
                HeaderText = "День",
                Name = "Day",
                Width = 35,
                ReadOnly = true
            };
            dayColumn.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dayColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            monthGrid.Columns.Add(dayColumn);

            var hourColumn = new DataGridViewTextBoxColumn
            {
                HeaderText = "Час (0:00-23:00)",
                Name = "Hour",
                Width = 80,
                ReadOnly = true
            };
            hourColumn.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            hourColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            monthGrid.Columns.Add(hourColumn);

            var volumeColumn = new DataGridViewTextBoxColumn
            {
                HeaderText = "Объем (кВт·ч)",
                Name = "Volume",
                Width = 80
            };
            volumeColumn.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            volumeColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            volumeColumn.DefaultCellStyle.Format = "N2";
            monthGrid.Columns.Add(volumeColumn);
        }

        private void SetupHorizontalInput()
        {
            monthGrid.Columns.Clear();
            monthGrid.Rows.Clear();

            var dayColumn = new DataGridViewTextBoxColumn
            {
                HeaderText = "Дни",
                Name = "Day",
                Width = 30,
                Frozen = true,
                ReadOnly = true
            };
            dayColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            monthGrid.Columns.Add(dayColumn);

            for (int hour = 0; hour < 24; hour++)
            {
                var column = new DataGridViewTextBoxColumn
                {
                    HeaderText = $"{hour}:00",
                    Name = $"Hour_{hour}",
                    Width = 60,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleCenter,
                        Format = "N2"
                    }
                };
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                monthGrid.Columns.Add(column);
            }
        }

        private void InputTypeChanged(object sender, EventArgs e)
        {
            if (verticalRB.Checked)
                SetupVerticalInput();
            else
                SetupHorizontalInput();
        }

        private void RowsCB_SelectedIndexChanged(object sender, EventArgs e)
        {
            days = int.Parse(RowsCB.SelectedItem.ToString());
        }

        private void Fonts()
        {
            font = new PrivateFontCollection();
            font.AddFontFile(@"Data\font\Inter-V.ttf");
            verticalRB.Font = new Font(font.Families[0], 11);
            horizontalRB.Font = new Font(font.Families[0], 11);
            AddRowsB.Font = new Font(font.Families[0], 14);
            RowsCB.Font = new Font(font.Families[0], 20);
            monthGrid.Font = new Font(font.Families[0], 8);
            categoryTB.Font = new Font(font.Families[0], 9.25f);
            reasonTB.Font = new Font(font.Families[0], 10);
            dataInfo1TB.Font = new Font(font.Families[0], 10.75f);
            dataInfo2TB.Font = new Font(font.Families[0], 10.75f);
            dataInfo3TB.Font = new Font(font.Families[0], 10.75f);
            data1TB.Font = new Font(font.Families[0], 7.25f);
            data2TB.Font = new Font(font.Families[0], 7.25f);
            data3TB.Font = new Font(font.Families[0], 7.25f);
            exportB.Font = new Font(font.Families[0], 14);
        }

        private void exportB_Click(object sender, EventArgs e)
        {
            ExportResults(analysis);
        }

        private void ExportResults(string analysisText)
        {
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Text files (*.txt)|*.txt|Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv";
                saveDialog.Title = "Экспорт результатов";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveDialog.FileName;
                    string extension = Path.GetExtension(filePath).ToLower();

                    try
                    {
                        switch (extension)
                        {
                            case ".txt":
                                File.WriteAllText(filePath, analysisText);
                                break;

                            case ".csv":
                                ExportToCsv(filePath, analysisText);
                                break;

                            case ".xlsx":
                                ExportToExcel(filePath, analysisText);
                                break;

                            default:
                                MessageBox.Show("Неизвестный формат файла", "Ошибка",
                                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                        }

                        MessageBox.Show($"Результаты успешно экспортированы в {filePath}",
                                      "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при экспорте: {ex.Message}",
                                      "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ExportToCsv(string filePath, string analysisText)
        {
            var lines = analysisText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var csvContent = string.Join(Environment.NewLine, lines.Select(line => $"\"{line.Replace("\"", "\"\"")}\""));
            File.WriteAllText(filePath, csvContent);
        }

        private void ExportToExcel(string filePath, string analysisText)
        {
            ExcelPackage.License.SetNonCommercialOrganization("Non");

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Результаты");

                var lines = analysisText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                for (int i = 0; i < lines.Length; i++)
                {
                    worksheet.Cells[i + 1, 1].Value = lines[i];
                }

                package.Save();
            }
        }

        private void ChooseFileB_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                string filePath = @"Pattern\Пример Заполнения.xlsx";
                if (!File.Exists(filePath))
                {
                    return;
                }
                string argument = "/select, \"" + filePath + "\"";
                Process.Start("explorer.exe", argument);
            }
            else
            {
                VerticalTableParser vert = new VerticalTableParser(monthGrid);
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
                    openFileDialog.Title = "Выберите файл Excel";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        file = openFileDialog.FileName;
                        monthGrid.Rows.Clear();
                        monthGrid.Visible = true;
                        if (verticalRB.Checked)
                        {
                            monthGrid.Rows.Add(31 * 24);
                        }
                        else
                        {
                            monthGrid.Rows.Add(31);
                        }
                        (input1D, input2D) = vert.ProcessExcelFile(openFileDialog.FileName, verticalRB.Checked);

                        if (input2D == null)
                        {
                            int k = 0;
                            int g = 1;
                            for (int i = 0; i < input1D.GetLength(0); i++)
                            {
                                if (k < 24)
                                {
                                    monthGrid.Rows[i].Cells[0].Value = g;
                                    k++;
                                }
                                else
                                {
                                    k = 0;
                                    g++;
                                    monthGrid.Rows[i].Cells[0].Value = g;
                                    k++;
                                }
                                monthGrid.Rows[i].Cells[1].Value = $"{k - 1}:00";
                                try
                                {
                                    monthGrid.Rows[i].Cells[2].Value = input1D[i];
                                }
                                catch
                                {
                                    int val = int.Parse(monthGrid.Rows[monthGrid.RowCount - 1].Cells[0].Value.ToString());
                                    monthGrid.Rows.Add(1);
                                    if (k < 24)
                                    {
                                        monthGrid.Rows[i].Cells[0].Value = g;
                                        k++;
                                    }
                                    else
                                    {
                                        k = 0;
                                        g++;
                                        monthGrid.Rows[i].Cells[0].Value = g;
                                        k++;
                                    }
                                    monthGrid.Rows[i].Cells[1].Value = $"{k - 1}:00";
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < input2D.GetLength(0); i++)
                            {
                                monthGrid.Rows[i].Cells[0].Value = i + 1;
                                for (int j = 0; j < input2D.GetLength(1); j++)
                                {
                                    try
                                    {
                                        monthGrid.Rows[i].Cells[j + 1].Value = input2D[i, j];
                                    }
                                    catch
                                    {
                                        int val = int.Parse(monthGrid.Rows[monthGrid.RowCount - 1].Cells[0].Value.ToString());
                                        monthGrid.Rows.Add(1);
                                        monthGrid.Rows[monthGrid.RowCount - 1].Cells[0].Value = val + 1;
                                        monthGrid.Rows[i].Cells[j + 1].Value = input2D[i, j];
                                    }
                                }
                            }
                            for (int i = 0; i < input2D.GetLength(0); i++)
                            {
                                if (input2D[i, 0] == 0)
                                {
                                    monthGrid.Rows.RemoveAt(i);
                                }
                            }
                        }

                    }
                }
            }
        }

        private void HistoryB_Click(object sender, EventArgs e)
        {
            HistoryForm history = new HistoryForm();
            history.Show();
        }

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            VerticalTableParser vert = new VerticalTableParser(monthGrid);
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0 && (files[0].EndsWith(".xlsx") || files[0].EndsWith(".xls")))
            {
                file = files[0];
                monthGrid.Rows.Clear();
                monthGrid.Visible = true;
                if (verticalRB.Checked)
                {
                    monthGrid.Rows.Add(31 * 24);
                }
                else
                {
                    monthGrid.Rows.Add(31);
                }
                (input1D, input2D) = vert.ProcessExcelFile(files[0], verticalRB.Checked);

                if (input2D == null)
                {
                    int k = 0;
                    int g = 1;
                    for (int i = 0; i < input1D.GetLength(0); i++)
                    {
                        if (k < 24)
                        {
                            monthGrid.Rows[i].Cells[0].Value = g;
                            k++;
                        }
                        else
                        {
                            k = 0;
                            g++;
                            monthGrid.Rows[i].Cells[0].Value = g;
                            k++;
                        }
                        monthGrid.Rows[i].Cells[1].Value = $"{k - 1}:00";
                        try
                        {
                            monthGrid.Rows[i].Cells[2].Value = input1D[i];
                        }
                        catch
                        {
                            int val = int.Parse(monthGrid.Rows[monthGrid.RowCount - 1].Cells[0].Value.ToString());
                            monthGrid.Rows.Add(1);
                            if (k < 24)
                            {
                                monthGrid.Rows[i].Cells[0].Value = g;
                                k++;
                            }
                            else
                            {
                                k = 0;
                                g++;
                                monthGrid.Rows[i].Cells[0].Value = g;
                                k++;
                            }
                            monthGrid.Rows[i].Cells[1].Value = $"{k - 1}:00";
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < input2D.GetLength(0); i++)
                    {
                        monthGrid.Rows[i].Cells[0].Value = i + 1;
                        for (int j = 0; j < input2D.GetLength(1); j++)
                        {
                            try
                            {
                                monthGrid.Rows[i].Cells[j + 1].Value = input2D[i, j];
                            }
                            catch
                            {
                                int val = int.Parse(monthGrid.Rows[monthGrid.RowCount - 1].Cells[0].Value.ToString());
                                monthGrid.Rows.Add(1);
                                monthGrid.Rows[monthGrid.RowCount - 1].Cells[0].Value = val + 1;
                                monthGrid.Rows[i].Cells[j + 1].Value = input2D[i, j];
                            }
                        }
                    }
                    for (int i = 0; i < input2D.GetLength(0); i++)
                    {
                        if (input2D[i, 0] == 0)
                        {
                            monthGrid.Rows.RemoveAt(i);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Пожалуйста, перетащите файл Excel (.xls или .xlsx)",
                               "Неверный формат", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
