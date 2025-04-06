using ExcelDataReader;
using MathNet.Numerics;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PriceCategory
{
    class SolveCategory
    {
        private DataTable _originalData;

        private TextBox categoryTB;
        private TextBox reasonTB;
        private TextBox dataInfo1TB;
        private TextBox dataInfo2TB;
        private TextBox dataInfo3TB;
        private TextBox data1TB;
        private TextBox data2TB;
        private TextBox data3TB;
        public static string desc;

        public SolveCategory(TextBox categoryTB, TextBox reasonTB, TextBox dataInfo1TB, TextBox dataInfo2TB, TextBox dataInfo3TB, TextBox data1TB, TextBox data2TB, TextBox data3TB)
        {
            this.categoryTB = categoryTB;
            this.reasonTB = reasonTB;
            this.dataInfo1TB = dataInfo1TB;
            this.dataInfo2TB = dataInfo2TB;
            this.dataInfo3TB = dataInfo3TB;
            this.data1TB = data1TB;
            this.data2TB = data2TB;
            this.data3TB = data3TB;
        }

        private readonly TariffParameters tariffs = new TariffParameters
        {
            FirstCategoryRate = 4.2m,
            SecondCategoryDayRate = 6.0m,
            SecondCategoryNightRate = 3.0m,
            PowerRate = 700,
            TransmissionSingleRate = 2.0m,
            TransmissionDoubleRateLoss = 0.35m,
            TransmissionDoubleRateMaintenance = 600,
            WholesaleMarketRate = 1.5m
        };

        public void ProcessExcelFile(string filePath, double vt)
        {
            try
            {
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
                            _originalData = result.Tables[0].Copy();
                            Calculate(vt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении файла: {ex.Message}",
                              "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Calculate(double vt)
        {
            try
            {
                ConsumptionData consumptionData;
                List<decimal> peakLoads = null;

                if (vt < 670)
                {
                    var dailyValues = new List<decimal>();


                    foreach (DataRow row in _originalData.Rows)
                    {
                        if (decimal.TryParse(row[0].ToString(), out decimal value))
                            dailyValues.Add(value);
                    }

                    consumptionData = new DailyConsumption(dailyValues);
                    peakLoads = consumptionData.GetPeakLoads();
                }
                else
                {
                    var hourlyValues = new List<List<decimal>>();
                    var currentDay = new List<decimal>();
                    foreach (DataRow row in _originalData.Rows)
                    {
                        if (decimal.TryParse(row[0].ToString(), out decimal value))
                        {
                            currentDay.Add(value);
                            if (currentDay.Count == 24)
                            {
                                hourlyValues.Add(new List<decimal>(currentDay));
                                currentDay.Clear();
                            }
                        }
                    }


                    if (hourlyValues.Count == 0)
                    {
                        throw new Exception("Не удалось загрузить данные. Проверьте формат файла.");
                    }

                    consumptionData = new HourlyConsumption(hourlyValues);
                    peakLoads = consumptionData.GetPeakLoads();
                }

                var calculator = new TariffCalculator(consumptionData, tariffs, reasonTB);
                var result = calculator.CalculateBestTariff();


                DisplayResults(result, peakLoads);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при расчете: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DisplayResults((PriceCategory BestCategory, string Analysis) result, List<decimal> peakLoads)
        {
            var analysis = new System.Text.StringBuilder();
            categoryTB.Text = $"{GetCategoryName(result.BestCategory)}";
            analysis.AppendLine($"РЕКОМЕНДУЕМАЯ ТАРИФНАЯ КАТЕГОРИЯ: {GetCategoryName(result.BestCategory)}\n");
            analysis.AppendLine(result.Analysis);

            analysis.AppendLine("ОБОСНОВАНИЕ РЕКОМЕНДАЦИИ:");
            analysis.AppendLine(desc);

            if (peakLoads != null && peakLoads.Any())
            {
                var allLoadValues = GetAllLoadValuesFromOriginalData();

                if (allLoadValues.Any())
                {
                    decimal maxLoad = Math.Round(allLoadValues.Max(), 2);
                    decimal minLoad = Math.Round(allLoadValues.Min(), 2);
                    decimal avgLoad = Math.Round(allLoadValues.Average(), 2);
                    decimal nonUniformityCoefficient = Math.Round(maxLoad / avgLoad, 2);

                    analysis.AppendLine("\nДЕТАЛЬНЫЙ АНАЛИЗ НАГРУЗКИ:");
                    analysis.AppendLine($"- Максимальная нагрузка: {maxLoad:N2} кВт (наибольшее значение из столбца 'Объем')");
                    analysis.AppendLine($"- Минимальная нагрузка: {minLoad:N2} кВт (наименьшее значение из столбца 'Объем')");
                    analysis.AppendLine($"- Средняя нагрузка: {avgLoad:N2} кВт");
                    analysis.AppendLine($"- Коэффициент неравномерности: {nonUniformityCoefficient:N2} " +
                                       "(отношение максимальной нагрузки к средней, показывает неравномерность потребления)");

                    analysis.AppendLine("\nПОЯСНЕНИЕ:");
                    analysis.AppendLine("Коэффициент неравномерности показывает, насколько максимальная нагрузка ");
                    analysis.AppendLine("превышает среднюю. Чем выше значение, тем более неравномерное ");
                    analysis.AppendLine("потребление электроэнергии у предприятия.");

                    dataInfo1TB.Text = "Максимальная нагрузка:";
                    dataInfo2TB.Text = "Минимальная нагрузка:";
                    dataInfo3TB.Text = "Коэффициент неравномерности:\n(отношение макс/сред)";
                    data1TB.Text = $"{maxLoad} кВт";
                    data2TB.Text = $"{minLoad} кВт";
                    data3TB.Text = $"{nonUniformityCoefficient}";
                }
            }
            MainForm.analysis = analysis.ToString();
        }

        private string GetCategoryName(PriceCategory category)
        {
            switch (category)
            {
                case PriceCategory.First: return "Первая (низкое потребление)";
                case PriceCategory.Second: return "Вторая (дневной/ночной тариф)";
                case PriceCategory.Third: return "Третья (с учетом мощности)";
                case PriceCategory.Fourth: return "Четвертая (двухставочный тариф)";
                case PriceCategory.Fifth: return "Пятая (третья с прогнозированием)";
                case PriceCategory.Sixth: return "Шестая (четвертая с прогнозированием)";
                default: return "Не определена";
            }
        }

        private List<decimal> GetAllLoadValuesFromOriginalData()
        {
            var values = new List<decimal>();

            if (_originalData != null && _originalData.Columns.Contains("Объем"))
            {
                foreach (DataRow row in _originalData.Rows)
                {
                    if (decimal.TryParse(row["Объем"].ToString(), out decimal value))
                    {
                        values.Add(value);
                    }
                }
            }

            return values;
        }

        public enum PriceCategory
        {
            First,
            Second,
            Third,
            Fourth,
            Fifth,
            Sixth
        }

        public class TariffParameters
        {
            public decimal FirstCategoryRate { get; set; }
            public decimal SecondCategoryDayRate { get; set; }
            public decimal SecondCategoryNightRate { get; set; }
            public decimal PowerRate { get; set; }
            public decimal TransmissionSingleRate { get; set; }
            public decimal TransmissionDoubleRateLoss { get; set; }
            public decimal TransmissionDoubleRateMaintenance { get; set; }
            public decimal WholesaleMarketRate { get; set; }
        }

        public abstract class ConsumptionData
        {
            public decimal TotalConsumption { get; protected set; }
            public abstract List<decimal> GetPeakLoads();
        }

        public class DailyConsumption : ConsumptionData
        {
            public List<decimal> DailyValues { get; }

            public DailyConsumption(List<decimal> dailyValues)
            {
                DailyValues = dailyValues;
                TotalConsumption = dailyValues.Sum();
            }

            public override List<decimal> GetPeakLoads()
            {
                return new List<decimal> { DailyValues.Max() };
            }
        }

        public class HourlyConsumption : ConsumptionData
        {
            public List<List<decimal>> HourlyValues { get; }

            public HourlyConsumption(List<List<decimal>> hourlyValues)
            {
                HourlyValues = hourlyValues;
                TotalConsumption = hourlyValues.Sum(day => day.Sum());
            }

            public override List<decimal> GetPeakLoads()
            {
                return HourlyValues.Select(day => day.Max()).ToList();
            }
        }

        public class TariffCalculator
        {
            private readonly ConsumptionData _consumption;
            private readonly TariffParameters _tariffs;
            private TextBox reasonTB;

            public TariffCalculator(ConsumptionData consumption, TariffParameters tariffs, TextBox reasonTB)
            {
                _consumption = consumption;
                _tariffs = tariffs;
                this.reasonTB = reasonTB;
            }

            public (PriceCategory BestCategory, string Analysis) CalculateBestTariff()
            {
                var costs = new Dictionary<PriceCategory, decimal>();
                var analysis = new System.Text.StringBuilder();

                costs[PriceCategory.First] = _consumption.TotalConsumption * _tariffs.FirstCategoryRate;

                decimal dayConsumption = _consumption.TotalConsumption * 0.6m;
                decimal nightConsumption = _consumption.TotalConsumption * 0.4m;
                costs[PriceCategory.Second] = dayConsumption * _tariffs.SecondCategoryDayRate +
                                            nightConsumption * _tariffs.SecondCategoryNightRate;

                List<decimal> peakLoads = null;
                if (_consumption is HourlyConsumption hourlyConsumption)
                {
                    decimal avgPeakLoad = hourlyConsumption.GetPeakLoads().Average();
                    peakLoads = hourlyConsumption.GetPeakLoads();

                    costs[PriceCategory.Third] = _consumption.TotalConsumption * _tariffs.WholesaleMarketRate +
                                              avgPeakLoad * _tariffs.PowerRate +
                                              _consumption.TotalConsumption * _tariffs.TransmissionSingleRate;

                    costs[PriceCategory.Fourth] = _consumption.TotalConsumption * _tariffs.WholesaleMarketRate +
                                               avgPeakLoad * _tariffs.PowerRate +
                                               _consumption.TotalConsumption * _tariffs.TransmissionDoubleRateLoss +
                                               avgPeakLoad * _tariffs.TransmissionDoubleRateMaintenance;

                    costs[PriceCategory.Fifth] = costs[PriceCategory.Third] * 0.95m;
                    costs[PriceCategory.Sixth] = costs[PriceCategory.Fourth] * 0.95m;
                }

                var bestCategory = costs.OrderBy(x => x.Value).First().Key;
                reasonTB.Text = GetCategoryAnalysis(bestCategory, peakLoads);
                desc = GetCategoryAnalysis(bestCategory, peakLoads);

                return (bestCategory, analysis.ToString());
            }

            public string GetCategoryAnalysis(PriceCategory category, List<decimal> peakLoads)
            {
                var analysis = new StringBuilder();

                switch (category)
                {
                    case PriceCategory.First:
                        analysis.AppendLine("- Ваше предприятие имеет низкое и стабильное потребление электроэнергии");
                        analysis.AppendLine("- Отсутствуют выраженные пики нагрузки в течение суток");
                        analysis.AppendLine("- Простая схема расчета - оплата только за фактически потребленный объем");
                        analysis.AppendLine("- Минимальные требования к учету и отчетности");
                        break;

                    case PriceCategory.Second:
                        analysis.AppendLine("- Ваше предприятие имеет выраженный ночной график потребления");
                        analysis.AppendLine("- Ночное потребление составляет значительную долю от общего объема");
                        analysis.AppendLine("- Выгодно для предприятий с ночным циклом работы (холодильные установки, освещение и т.д.)");
                        analysis.AppendLine("- Требуется раздельный учет дневного и ночного потребления");
                        break;

                    case PriceCategory.Third:
                        analysis.AppendLine("- Ваше предприятие имеет значительное потребление с выраженными пиками нагрузки");
                        analysis.AppendLine($"- Средняя пиковая нагрузка: {peakLoads?.Average():N0} кВт");
                        analysis.AppendLine("- Оптимально для предприятий с постоянным графиком работы");
                        analysis.AppendLine("- Требуется учет как объема потребления, так и мощности");
                        break;

                    case PriceCategory.Fourth:
                        analysis.AppendLine("- Ваше предприятие имеет управляемые пики нагрузки");
                        analysis.AppendLine($"- Максимальная нагрузка: {peakLoads?.Max():N0} кВт");
                        analysis.AppendLine("- Выгодно при возможности смещения нагрузки с пиковых периодов");
                        analysis.AppendLine("- Сложный расчет с учетом мощности, потерь и обслуживания сетей");
                        break;

                    case PriceCategory.Fifth:
                    case PriceCategory.Sixth:
                        analysis.AppendLine("- Ваше предприятие может точно прогнозировать потребление на год вперед");
                        analysis.AppendLine("- Доступна экономия до 5% за счет точного планирования");
                        analysis.AppendLine("- Требуется ежемесячная сверка фактического и планового потребления");
                        analysis.AppendLine("- Необходимо компенсировать отклонения от плана");
                        break;
                }

                return analysis.ToString();
            }
        }
    }
}