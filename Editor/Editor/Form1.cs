using System.Data;
using System.Globalization;
using CsvHelper;

namespace Editor
{
    public partial class Form1 : Form
    {
        DataTable dt = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filename = openFileDialog.FileName;

                using (var streamer = new StreamReader(filename))
                {
                    using (var csvReader = new CsvReader(streamer, CultureInfo.InvariantCulture))
                    {
                        csvReader.Read();
                        csvReader.ReadHeader();
                        string[] headerRow = csvReader.Context.Reader.HeaderRecord;


                        foreach (string header in headerRow)
                        {
                            dt.Columns.Add(header);
                        }

                        var records = csvReader.GetRecords<dynamic>().ToList();

                        foreach (var record in records)
                        {
                            DataRow dr = dt.NewRow();
                            int i = 0;
                            foreach (var recordInstance in record)
                            {
                                dr[i] = recordInstance.Value;
                                i++;
                            }
                            dt.Rows.Add(dr);
                        }
                    }
                }
            }

            csvDataGridView.DataSource = dt;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*
            saveFileDialog.InitialDirectory =@"C:\";
            saveFileDialog.Title = "Save csv Files";
            //saveFileDialog.Filter = "csv";
            saveFileDialog.CheckFileExists = true;
            saveFileDialog.CheckPathExists = true;
            saveFileDialog.DefaultExt = "csv";
            saveFileDialog.RestoreDirectory = true;
            */

            string filename = @"C:\Users\jands\Documents";
            using (var writer = new StreamWriter(filename))
            {
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    // Write columns
                    foreach (DataColumn column in dt.Columns)
                    {
                        csv.WriteField(column.ColumnName);
                    }
                    csv.NextRecord();

                    // Write row values
                    foreach (DataRow row in dt.Rows)
                    {
                        for (var i = 0; i < dt.Columns.Count; i++)
                        {
                            csv.WriteField(row[i]);
                        }
                        csv.NextRecord();
                    }

                }
            }
        }
    }
}