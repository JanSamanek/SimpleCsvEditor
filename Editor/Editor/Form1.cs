using System.Data;
using System.Globalization;
using CsvHelper;

namespace Editor
{
    public partial class Form1 : Form
    {
        DataSet dataSet = new DataSet();
        DataTable dtFinal = new DataTable();


        int recordNum = 0;
        public Form1()
        {
            InitializeComponent();
            this.csvDataGridView.DragDrop += new DragEventHandler(this.csvDataGridView_DragDrop);
            this.csvDataGridView.DragEnter += new DragEventHandler(this.csvDataGridView_Enter);
        }

        private void csvDataGridView_Enter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop);

                foreach(string file in fileList)
                {
                    loadCsvData(file);
                }
                for (int i = 0; i < dataSet.Tables.Count; i++)
                {
                    dtFinal.Merge(dataSet.Tables[i]);
                }

                csvDataGridView.DataSource = dtFinal;
            }
        }

        private void csvDataGridView_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void loadCsvData(string filename)
        {
            using (var streamer = new StreamReader(filename))
            {
                using (var csvReader = new CsvReader(streamer, CultureInfo.InvariantCulture))
                {
                    try
                    {
                        DataTable dt = new DataTable();
                        csvReader.Read();
                        csvReader.ReadHeader();
                        string[] headerRow = csvReader.Context.Reader.HeaderRecord;


                        foreach (string header in headerRow)
                        {
                            dt.Columns.Add(header);
                        }

                        var records = csvReader.GetRecords<dynamic>();

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
                       dataSet.Tables.Add(dt);
                    }
                    catch (CsvHelper.BadDataException)
                    {
                       MessageBox.Show("Please load a csv file.");
                    }
                }
            }
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filename = openFileDialog.FileName;
                loadCsvData(filename);
            }

            for(int i = 0; i < dataSet.Tables.Count; i++)
            {
                dtFinal.Merge(dataSet.Tables[i]);
            }
            csvDataGridView.DataSource = dtFinal;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            saveFileDialog.InitialDirectory =@"C:\";
            saveFileDialog.Title = "Save csv Files";
            saveFileDialog.Filter = "csv files (*.csv)|*.csv";
            saveFileDialog.CheckFileExists = false;
            saveFileDialog.CheckPathExists = true;
            saveFileDialog.DefaultExt = "csv";
            saveFileDialog.RestoreDirectory = true;
          
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filename = saveFileDialog.FileName;

                using (var writer = new StreamWriter(filename))
                {
                    using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                    {
                        // Write columns
                        foreach (DataColumn column in dtFinal.Columns)
                        {
                            csv.WriteField(column.ColumnName);
                        }
                        csv.NextRecord();

                        // Write row values
                        foreach (DataRow row in dtFinal.Rows)
                        {
                            for (var i = 0; i < dtFinal.Columns.Count; i++)
                            {
                                csv.WriteField(row[i]);
                            }
                            csv.NextRecord();
                        }

                    }
                }
            }
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataSet = new DataSet();
            dtFinal = new DataTable();
            csvDataGridView.DataSource = dtFinal;
        }
    }
}