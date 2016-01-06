using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel;

namespace test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataSet result = new DataSet();

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = "in";

            if (fileName == "")
            {
                MessageBox.Show("Enter Valid file name");
                return;
            }

            converToCSV(comboBox1.SelectedIndex);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string Chosen_File = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Chosen_File = openFileDialog1.FileName;
            }
            if (Chosen_File == String.Empty)
            {
                return;
            }
            textBox1.Text = Chosen_File;

            getExcelData(textBox1.Text);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = this.folderBrowserDialog1.ShowDialog();
            string foldername = "";
            if (result == DialogResult.OK)
            {
                foldername = this.folderBrowserDialog1.SelectedPath;
            }
        }

        private void getExcelData(string file)
        {

            if (file.EndsWith(".xlsx"))
            {
                // Reading from a binary Excel file (format; *.xlsx)
                FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                result = excelReader.AsDataSet();
                excelReader.Close();
            }

            if (file.EndsWith(".xls"))
            {
                // Reading from a binary Excel file ('97-2003 format; *.xls)
                FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                result = excelReader.AsDataSet();
                excelReader.Close();
            }

            List<string> items = new List<string>();
            for (int i = 0; i < result.Tables.Count; i++)
                items.Add(result.Tables[i].TableName.ToString());
            comboBox1.DataSource = items;

        }

        private void converToCSV(int ind)
        {
            // sheets in excel file becomes tables in dataset
            //result.Tables[0].TableName.ToString(); // to get sheet name (table name)

            string a = "Year,Date,National Number\r\n";
            string b = "";
            string c = "";
            int row_no = 1;//not from 0 because it contains column-labels
            double d = 0.0;
            int date_col = 0;
            int nat_col = 0;

            for (int i = 0; i < result.Tables[ind].Columns.Count; i++)
            {
                if (result.Tables[ind].Rows[0][i].ToString().Contains("DATE") || result.Tables[ind].Rows[0][i].ToString().Contains("Date")
                    || result.Tables[ind].Rows[0][i].ToString().Contains("date"))
                    date_col = i;
                if (result.Tables[ind].Rows[0][i].ToString().Contains("NAT") || result.Tables[ind].Rows[0][i].ToString().Contains("Nat")
                    || result.Tables[ind].Rows[0][i].ToString().Contains("nat"))
                    nat_col = i;
            }

            for (int i1 = row_no; i1 < result.Tables[ind].Rows.Count; i1++)
            {
                b = "";

                for (int i2 = 0; i2 < result.Tables[ind].Columns.Count; i2++)
                {

                    if (result.Tables[ind].Rows[i1][i2].ToString() == "")
                        continue;

                    if (i2 == date_col)
                    {
                        if (double.TryParse(result.Tables[ind].Rows[i1][i2].ToString(), out d))
                        {
                            DateTime date = DateTime.FromOADate(d);
                            b += date.Year.ToString() + ",";

                            if (date.Day < 10)
                                b += "0";
                            b += date.Day.ToString() + "-";
                            if (date.Month < 10)
                                b += "0";
                            b += date.Month.ToString() + "-" + date.Year.ToString() + ",";
                            //b += result.Tables[ind].Rows[i1][i].ToString().Replace("/", "-") + ","; //this was for when input date was not formated(text only)
                        }
                    }
                    if (i2 == nat_col)
                    {
                        c = result.Tables[ind].Rows[i1][i2].ToString();
                    }
                    /*else
                        b += result.Tables[ind].Rows[i1][i2].ToString() + ",";
                     * */
                }
                if (b == "")
                    continue;
                /*
                char[] cA = new char[b.Length + 2];
                cA = b.ToCharArray();
                for (int i = 0; i < cA.Length; i++)
                {
                    int counter = cA.Length - 1;

                    if (cA[i] == '-')
                        if (cA[i - 2] != '0' && cA[i - 2] != '1' && cA[i - 2] != '2' && cA[i - 2] != '3' && cA[i - 2] != '4' && cA[i - 2] != '5' && cA[i - 2] != '6'
                            && cA[i - 2] != '7' && cA[i - 2] != '8' && cA[i - 2] != '9')
                        {
                            while (counter > i - 2)
                            {
                                cA[counter] = cA[counter - 1];
                                counter--;
                            }
                            cA[i - 1] = '0';
                            i++;
                        }
                }
                b = "";
                foreach (char c in cA)
                    b += c.ToString();
                 * */
                a += b + c + "\r\n";
            }

            string output = Program.strPath + "\\" + "in.csv";
            StreamWriter csv = new StreamWriter(@output, false, Encoding.UTF8);
            csv.Write(a);
            csv.Close();

            MessageBox.Show("File converted succussfully");

            textBox1.Text = "";
            comboBox1.DataSource = null;
            return;
        }

    }
}/*while (row_no < result.Tables[ind].Rows.Count)
            {
                string b = "";

                for (int i = 0; i < result.Tables[ind].Columns.Count; i++)
                {

                    if (result.Tables[ind].Rows[row_no][i].ToString() == "")
                        continue;

                    if (i == date_col)
                    {
                        if (double.TryParse(result.Tables[ind].Rows[row_no][i].ToString(), out d))
                        {
                            DateTime date = DateTime.FromOADate(d);
                            string s = "";
                            b += date.ToShortDateString() + ",";

                            //b += result.Tables[ind].Rows[row_no][i].ToString().Replace("/", "-") + ",";
                        }
                    }
                    else
                        b += result.Tables[ind].Rows[row_no][i].ToString() + ",";
                }
                row_no++;
                if (b == "")
                    continue;
                /*
                char[] cA = new char[b.Length + 2];
                cA = b.ToCharArray();
                for (int i = 0; i < cA.Length; i++)
                {
                    int counter = cA.Length - 1;

                    if (cA[i] == '-')
                        if (cA[i - 2] != '0' && cA[i - 2] != '1' && cA[i - 2] != '2' && cA[i - 2] != '3' && cA[i - 2] != '4' && cA[i - 2] != '5' && cA[i - 2] != '6'
                            && cA[i - 2] != '7' && cA[i - 2] != '8' && cA[i - 2] != '9')
                        {
                            while (counter > i - 2)
                            {
                                cA[counter] = cA[counter - 1];
                                counter--;
                            }
                            cA[i - 1] = '0';
                            i++;
                        }
                }
                b = "";
                foreach (char c in cA)
                    b += c.ToString();
                
                a += b + "\r\n";
            }*/