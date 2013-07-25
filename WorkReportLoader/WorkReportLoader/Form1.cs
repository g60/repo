using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WorkReportLoader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();

            fdlg.Title = "Select file";
            fdlg.InitialDirectory = @"c:\";
            fdlg.FileName = txtFileName.Text;
            fdlg.Filter = "Text and CSV Files(*.txt, *.csv)|*.txt;*.csv|Text Files(*.txt)|*.txt|CSV Files(*.csv)|*.csv|All Files(*.*)|*.*";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtFileName.Text = fdlg.FileName;
                Import(fdlg.FileName);
                Application.DoEvents();
            } // end if

        }

        DataTable dt_work = new DataTable();

        private void Import(string filename)
        {
            int lineNo = 0;

            dt_work.Clear();
            dt_work.Columns.Clear();


            TextFieldParser parser = new TextFieldParser(filename);
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            while (!parser.EndOfData)
            {
                string[] fields = parser.ReadFields();

                if (lineNo == 0)
                {
                    
                    for (int i = 0; i < fields.Length; i++)
                    {
                        dt_work.Columns.Add(fields[i]);
                    } // end for
                }
                else
                {
                    DataRow rowToAdd = dt_work.NewRow();

                    rowToAdd.ItemArray = fields;

                    //for (int i = 0; i < fields.Length; i++)
                    //{
                    //    rowToAdd.ItemArray[i] = fields[i];
                    //} // end for
                    
                    dt_work.Rows.Add(rowToAdd);
                } // end if-then-else

                lineNo++;

            }



            parser.Close();

            dataGridView1.DataSource = dt_work;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable new_dt_work = dt_work.Clone();

            List<int> rowsToRemove = new List<int>();

            string curr_tech = "";

            // have to process the table backwards!!
            for (int i = dt_work.Rows.Count-1; i >= 0; i--)
            {

                DataRow curr_row = dt_work.Rows[i];

                // remove any blank rows
                if (curr_row["Prefix"].ToString() == "")
                {
                    if (!rowsToRemove.Contains(i))
                    {
                        rowsToRemove.Add(i);
                    } // end if
                } // end if

                // remove any total rows
                if (String.Compare(curr_row["Prefix"].ToString(), "Totals", true) == 0)
                {
                    if (!rowsToRemove.Contains(i))
                    {
                        rowsToRemove.Add(i);
                    } // end if
                } // end if

                // remove any grand total rows
                if (String.Compare(curr_row["Prefix"].ToString(), "Grand Totals", true) == 0)
                {
                    if (!rowsToRemove.Contains(i))
                    {
                        rowsToRemove.Add(i);
                    } // end if
                } // end if

                // fill in the technicians name
                if (curr_row["TechInit"].ToString() == "")
                {
                    curr_row["TechInit"] = curr_tech;
                }
                else
                {
                    curr_tech = curr_row["TechInit"].ToString();
                } // end if-then-else

            } // end for

            

            foreach (int rowToRemove in rowsToRemove)
            {
                dt_work.Rows.RemoveAt(rowToRemove);
            } // end foreach

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string CsvFpath = @"C:\scanner\CSV-EXPORT.csv";
            try
            {
                System.IO.StreamWriter csvFileWriter = new StreamWriter(CsvFpath, false);

                string columnHeaderText = "";

                int countColumn = dt_work.Columns.Count;

                columnHeaderText = dt_work.Columns[0].ColumnName;

                for (int i = 1; i < countColumn; i++)
                {
                    columnHeaderText = columnHeaderText + ',' + dt_work.Columns[i].ColumnName;
                }

                csvFileWriter.WriteLine(columnHeaderText);

                foreach (DataRow currRow in dt_work.Rows)
                {
                    for (int i = 0; i < currRow.ItemArray.Length; i++)
                    {
                        if (i > 0)
                        {
                            csvFileWriter.Write(",");
                        } // end if

                        csvFileWriter.Write("\"" + currRow.ItemArray[i] + "\"");
                    } // end for

                    csvFileWriter.WriteLine();

                } // end foreach

                //string columnHeaderText = "";

                //int countColumn = dataGridView1.ColumnCount - 1;

                //if (countColumn >= 0)
                //{
                //    columnHeaderText = dataGridView1.Columns[0].HeaderText;
                //}

                //for (int i = 1; i <= countColumn; i++)
                //{
                //    columnHeaderText = columnHeaderText + ',' + dataGridView1.Columns[i].HeaderText;
                //}


                //csvFileWriter.WriteLine(columnHeaderText);

                //foreach (DataGridViewRow dataRowObject in dataGridView1.Rows)
                //{
                //    if (!dataRowObject.IsNewRow)
                //    {
                //        string dataFromGrid = "";

                //        dataFromGrid = dataRowObject.Cells[0].Value.ToString();

                //        for (int i = 1; i <= countColumn; i++)
                //        {
                //            dataFromGrid = dataFromGrid + ',' + dataRowObject.Cells[i].Value.ToString();

                //            csvFileWriter.Write(dataFromGrid);
                //        }
                //        csvFileWriter.WriteLine();
                //    }
                //}


                csvFileWriter.Flush();
                csvFileWriter.Close();
            }
            catch (Exception exceptionObject)
            {
                MessageBox.Show(exceptionObject.ToString());
            }
        }
    }
}
