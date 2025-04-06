using System;
using System.Data.SQLite;
using System.Drawing.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;

namespace PriceCategory
{
    public partial class HistoryForm: Form
    {
        public HistoryForm()
        {
            InitializeComponent();
        }

        PrivateFontCollection font;
        private string Path;

        private void HistoryForm_Load(object sender, EventArgs e)
        {
            font = new PrivateFontCollection();
            font.AddFontFile(@"Data\font\Inter-V.ttf");
            LoadData();
        }

        private void LoadData()
        {
            string dbPath = @"Data Source=Data\history.sqlite;Version=3;";

            using (var connection = new SQLiteConnection(dbPath))
            {
                connection.Open();

                string selectQuery = "SELECT Date, Text, Link FROM History";
                using (var command = new SQLiteCommand(selectQuery, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string dateT = reader.GetString(0);
                            string textT = reader.GetString(1);
                            string linkT = reader.GetString(2);

                            Label date = new Label
                            {
                                Text = $"{dateT}",
                                AutoSize = true,
                                MaximumSize = new Size(HistoryLayout.Width - 20, 0),
                                Margin = new Padding(5),
                                Font = new Font(font.Families[0], 12),
                                ForeColor = Color.FromArgb(21, 154, 53)
                            };
                            HistoryLayout.Controls.Add(date);

                            Label text = new Label
                            {
                                Text = $"{textT}",
                                AutoSize = true,
                                MaximumSize = new Size(HistoryLayout.Width - 20, 0),
                                Margin = new Padding(5),
                                Font = new Font(font.Families[0], 12),
                                ForeColor = Color.FromArgb(216, 205, 1)
                            };
                            HistoryLayout.Controls.Add(text);

                            string t = "";
                            if (linkT == "")
                            {
                                t = "Данные были введены вручную";
                            }
                            else
                            {
                                t = "Импортированный файл";
                                Path = reader.GetString(2);
                            }

                            LinkLabel link = new LinkLabel
                            {
                                Text = t,
                                AutoSize = true,
                                MaximumSize = new Size(HistoryLayout.Width - 20, 0),
                                Margin = new Padding(5),
                                Font = new Font(font.Families[0], 12),
                                ForeColor = Color.FromArgb(18, 185, 230)
                            };
                            link.LinkColor = Color.FromArgb(18, 185, 230);
                            link.VisitedLinkColor = Color.FromArgb(18, 185, 230);
                            link.ActiveLinkColor = Color.FromArgb(18, 185, 230);

                            if (!(linkT == ""))
                            {
                                link.LinkClicked += Link_LinkClicked;
                            }

                            HistoryLayout.Controls.Add(link);

                            Label space = new Label
                            {
                                Text = "---------------------------------------------------------------",
                                AutoSize = true,
                                MaximumSize = new Size(HistoryLayout.Width - 20, 0),
                                Margin = new Padding(5),
                                Font = new Font(font.Families[0], 12),
                                ForeColor = Color.FromArgb(18, 185, 230)
                            };
                            HistoryLayout.Controls.Add(space);
                        }
                    }
                }
            }
        }

        private void Link_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            Process.Start("explorer.exe", $"/select,\"{Path}\"");
        }
    }
}
