using System;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;

namespace PriceCategory
{
    public partial class MaxVTForm: Form
    {
        private TextBox categoryTB;
        private TextBox reasonTB;
        private TextBox dataInfo1TB;
        private TextBox dataInfo2TB;
        private TextBox dataInfo3TB;
        private TextBox data1TB;
        private TextBox data2TB;
        private TextBox data3TB;
        private Button exportB;
        private string file;
        public double countVt;

        public MaxVTForm(TextBox categoryTB, TextBox reasonTB, TextBox dataInfo1TB, TextBox dataInfo2TB, TextBox dataInfo3TB, TextBox data1TB, TextBox data2TB, TextBox data3TB, Button exportB, string file)
        {
            InitializeComponent();
            this.categoryTB = categoryTB;
            this.reasonTB = reasonTB;
            this.dataInfo1TB = dataInfo1TB;
            this.dataInfo2TB = dataInfo2TB;
            this.dataInfo3TB = dataInfo3TB;
            this.data1TB = data1TB;
            this.data2TB = data2TB;
            this.data3TB = data3TB;
            this.exportB = exportB;
            this.file = file;
        }

        private void DoneB_Click(object sender, EventArgs e)
        {
            try
            {
                switch (vtCB.SelectedIndex)
                {
                    case -1:
                        countVt = Double.Parse(maxPowerTB.Text);
                        break;

                    case 0:
                        countVt = Double.Parse(maxPowerTB.Text);
                        break;

                    case 1:
                        countVt = Double.Parse(maxPowerTB.Text) * 1000;
                        break;
                }

                if (countVt < 0)
                {
                    throw new Exception();
                }
                SolveCategory solve = new SolveCategory(categoryTB, reasonTB, dataInfo1TB, dataInfo2TB, dataInfo3TB, data1TB, data2TB, data3TB);
                solve.ProcessExcelFile(file, countVt);
                exportB.Visible = true;

                string dbPath = @"Data Source=Data\history.sqlite;Version=3;";

                using (var connection = new SQLiteConnection(dbPath))
                {
                    connection.Open();

                    string createTableQuery = @"
                CREATE TABLE IF NOT EXISTS History (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Date TEXT NOT NULL,
                    Text TEXT NOT NULL,
                    Link TEXT NOT NULL
                );";

                    using (var command = new SQLiteCommand(createTableQuery, connection))
                    {
                        command.ExecuteNonQuery();
                    }

                    string folder = @"FilesHistory";

                    if (!Directory.Exists(folder))
                    {
                        Directory.CreateDirectory(folder);
                    }

                    string fileName = $"{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx";
                    string dest = Path.Combine(folder, fileName);

                    string insertQuery = "INSERT INTO History (Date, Text, Link) VALUES (@date, @text, @link)";
                    using (var command = new SQLiteCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        command.Parameters.AddWithValue("@text", MainForm.analysis);
                        command.Parameters.AddWithValue("@link", dest);

                        command.ExecuteNonQuery();
                    }
                }

                string destinationFolder = @"FilesHistory";

                if (!Directory.Exists(destinationFolder))
                {
                    Directory.CreateDirectory(destinationFolder);
                }

                string newFileName = $"{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx";
                string destination = Path.Combine(destinationFolder, newFileName);

                File.Copy(file, destination, overwrite: true);
                this.Close();
            }
            catch { MessageBox.Show("Неккоректный формат данных!", "Ошибка"); }
        }
    }
}
