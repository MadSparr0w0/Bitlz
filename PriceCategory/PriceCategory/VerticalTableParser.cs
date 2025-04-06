using ExcelDataReader;
using MathNet.Numerics;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PriceCategory
{
    class VerticalTableParser
    {
        private DataGridView grid;

        public VerticalTableParser(DataGridView grid)
        {
            this.grid = grid;
        }

        private DataTable FilterEmptyRows(DataTable sourceTable)
        {
            DataTable filteredTable = sourceTable.Clone();

            foreach (DataRow row in sourceTable.Rows)
            {
                if (!row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field.ToString())))
                {
                    filteredTable.ImportRow(row);
                }
            }

            return filteredTable;
        }

        public (double[], double[,]) ProcessExcelFile(string filePath, Boolean isVert)
        {
            //try
            //{
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        if (result.Tables.Count > 0)
                        {
                            var cleanTable = FilterEmptyRows(result.Tables[0]);
                            var processedTable = ProcessHourlyData(cleanTable, isVert);
                            return processedTable;
                        }
                    }
                }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"Ошибка при чтении файла: {ex.Message}",
            //                  "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            return (null, null);
        }

        private (double[], double[,]) ProcessHourlyData(DataTable sourceTable, Boolean isVert)
        {
            if (!sourceTable.Columns.Contains("День") || !sourceTable.Columns.Contains("Час") || !sourceTable.Columns.Contains("Объем"))
            {
                throw new Exception("Файл должен содержать столбцы 'День', 'Час' и 'Объем'");
            }

            double[] input1D = null;
            double[,] input2D = null;

            if (isVert)
            {
                int i = 0;
                input1D = new double[sourceTable.Rows.Count];
                foreach (DataRow row in sourceTable.Rows)
                {
                    input1D[i] = Double.Parse(row.ItemArray[2].ToString()).Round(2);
                    i++;
                }
            }
            else
            {
                int i = 0;
                int j = 0;
                input2D = new double[31, grid.ColumnCount - 1];

                foreach (DataRow row in sourceTable.Rows)
                {
                    input2D[j, i] = Double.Parse(row.ItemArray[2].ToString()).Round(2);
                    i++;

                    if (i >= 24)
                    {
                        i = 0;
                        j++;
                    }
                }
            }

            return (input1D, input2D);
        }
    }
}
