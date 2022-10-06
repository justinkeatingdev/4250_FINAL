using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Microsoft.VisualBasic;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.WinFormsUtilities;



namespace WindowsFormsApp1
{


    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        public void Form1_Load(object sender, EventArgs e)
        {
            MessageBox.Show("Welcome to 4250 Final_Project Demo, click ok to begin.");

            dataGridView1.Rows.Add(label6.Text);
            dataGridView1.Rows.Add(label7.Text);
            dataGridView1.Rows.Add(label8.Text);
            dataGridView1.Rows.Add(label10.Text);
            dataGridView1.Rows.Add(label13.Text);
            dataGridView1.Rows.Add();
            dataGridView1.Rows.Add("HW Avg.");
            dataGridView1.Rows.Add("TEST Avg.");
            dataGridView1.Rows.Add("FINAL GRADE");

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");


        }


        public string studentName;


        private void button2_Click(object sender, EventArgs e)
        {

            string message, title;
            
            message = "Enter in Student Name: ";
            title = "Student";


            studentName = Interaction.InputBox(message, title);
            label3.Text = studentName;
            label11.Text = studentName;

        }


        public void button3_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Cells[1].Value = "";
            }


            if (comboBox1.Text == "Math")
            {
                TextWriter txt = new StreamWriter("C:\\4250_FINAL\\studentDetailsMath.txt");
                txt.Write("Student: " + studentName + "\n" + "Class: " + comboBox1.Text
                    + "\n" + "Homework 1: " + textBox1.Text + "\n" + "Homework 2: " + textBox2.Text + "\n" + "Homework 3: " + textBox3.Text
                    + "\n" + "Test 1: " + textBox4.Text + "\n" + "Test 2: " + textBox5.Text);
                txt.Close();

                dataGridView1.Rows[0].Cells[1].Value = (textBox1.Text + " = " + textBox6.Text);
                dataGridView1.Rows[1].Cells[1].Value = (textBox2.Text + " = " + textBox7.Text) ;
                dataGridView1.Rows[2].Cells[1].Value = (textBox3.Text + " = " + textBox8.Text);
                dataGridView1.Rows[3].Cells[1].Value = (textBox4.Text + " = " + textBox9.Text);
                dataGridView1.Rows[4].Cells[1].Value = (textBox5.Text + " = " + textBox10.Text);

                double Mhw1 = Int32.Parse(textBox6.Text);
                double Mhw2 = Int32.Parse(textBox7.Text);
                double Mhw3 = Int32.Parse(textBox8.Text);
                double Mtst1 = Int32.Parse(textBox9.Text);
                double Mtst2 = Int32.Parse(textBox10.Text);



                double Mhwavg = (Mhw1 + Mhw2 + Mhw3) / 300 * 100;
                double Mtstavg = (Mtst1 + Mtst2) / 200 * 100;

                string mhwcalc = Mhwavg.ToString("0.00");
                string mtstcalc = Mtstavg.ToString("0.00");

                dataGridView1.Rows[6].Cells[1].Value = mhwcalc  + "%"  ;
                dataGridView1.Rows[7].Cells[1].Value = mtstcalc  + "%"  ;


                double Moverall = (Mhw1 + Mhw2 + Mhw3 + Mtst1 + Mtst2) / 500 * 100;


                string mFinal = Moverall.ToString("0.00");



                if (Moverall >= 90 && Moverall <= 100)
                {
                    dataGridView1.Rows[8].Cells[1].Value = mFinal + "%" + " = A";
                }
                else

                if (Moverall >= 80 && Moverall <= 89)
                {
                    dataGridView1.Rows[8].Cells[1].Value = mFinal + "%" + " = B";
                }
                else

                if (Moverall >= 70 && Moverall <= 79)
                {
                    dataGridView1.Rows[8].Cells[1].Value = mFinal + "%" + " = C";
                }

                else

                if (Moverall > 60 && Moverall <= 69)
                {
                    dataGridView1.Rows[8].Cells[1].Value = mFinal + "%" + " = D";
                }

                else

                if (Moverall >= 0 && Moverall <= 59)
                {
                    dataGridView1.Rows[8].Cells[1].Value = mFinal + "%" + " = F";
                }



            }


            var count = this.dataGridView1.Rows.Cast<DataGridViewRow>()
               .Count(row => row.Cells["Math"].Value.ToString().Equals(" = 100"));

            this.textBox11.Text = count.ToString();




            if (comboBox1.Text == "Science")
            {
                TextWriter txt = new StreamWriter("C:\\4250_FINAL\\studentDetailsScience.txt");
                txt.Write("Student: " + studentName + "\n" + "Class: " + comboBox1.Text
                    + "\n" + "Homework 1: " + textBox1.Text + "\n" + "Homework 2: " + textBox2.Text + "\n" + "Homework 3: " + textBox3.Text
                    + "\n" + "Test 1: " + textBox4.Text + "\n" + "Test 2: " + textBox5.Text);
                txt.Close();

                dataGridView1.Rows[0].Cells[2].Value = (textBox1.Text + " = " + textBox6.Text);
                dataGridView1.Rows[1].Cells[2].Value = (textBox2.Text + " = " + textBox7.Text);
                dataGridView1.Rows[2].Cells[2].Value = (textBox3.Text + " = " + textBox8.Text);
                dataGridView1.Rows[3].Cells[2].Value = (textBox4.Text + " = " + textBox9.Text);
                dataGridView1.Rows[4].Cells[2].Value = (textBox5.Text + " = " + textBox10.Text);

                double Shw1 = Int32.Parse(textBox6.Text);
                double Shw2 = Int32.Parse(textBox7.Text);
                double Shw3 = Int32.Parse(textBox8.Text);
                double Stst1 = Int32.Parse(textBox9.Text);
                double Stst2 = Int32.Parse(textBox10.Text);

                double Shwavg = (Shw1 + Shw2 + Shw3) / 300 * 100;
                double Ststavg = (Stst1 + Stst2) / 200 * 100;

                string Shwcalc = Shwavg.ToString("0.00");
                string Ststcalc = Ststavg.ToString("0.00");


                dataGridView1.Rows[6].Cells[2].Value = Shwcalc + "%";
                dataGridView1.Rows[7].Cells[2].Value = Ststcalc + "%";


                double Soverall = (Shw1 + Shw2 + Shw3 + Stst1 + Stst2) / 500 * 100;



                string SFinal = Soverall.ToString("0.00");



                if (Soverall >= 90 && Soverall <= 100)
                {
                    dataGridView1.Rows[8].Cells[2].Value = SFinal + "%" + " = A";
                }
                else

                if (Soverall >= 80 && Soverall <= 89)
                {
                    dataGridView1.Rows[8].Cells[2].Value = SFinal + "%" + " = B";
                }
                else

                if (Soverall >= 70 && Soverall <= 79)
                {
                    dataGridView1.Rows[8].Cells[2].Value = SFinal + "%" + " = C";
                }

                else

                if (Soverall > 60 && Soverall <= 69)
                {
                    dataGridView1.Rows[8].Cells[2].Value = SFinal + "%" + " = D";
                }

                else

                if (Soverall >= 0 && Soverall <= 59)
                {
                    dataGridView1.Rows[8].Cells[2].Value = SFinal + "%" + " = F";
                }



            }



            if (comboBox1.Text == "History")
            {
                TextWriter txt = new StreamWriter("C:\\4250_FINAL\\studentDetailsHistory.txt");
                txt.Write("Student: " + studentName + "\n" + "Class: " + comboBox1.Text
                    + "\n" + "Homework 1: " + textBox1.Text + "\n" + "Homework 2: " + textBox2.Text + "\n" + "Homework 3: " + textBox3.Text
                    + "\n" + "Test 1: " + textBox4.Text + "\n" + "Test 2: " + textBox5.Text);
                txt.Close();

                dataGridView1.Rows[0].Cells[3].Value = (textBox1.Text + " = " + textBox6.Text);
                dataGridView1.Rows[1].Cells[3].Value = (textBox2.Text + " = " + textBox7.Text);
                dataGridView1.Rows[2].Cells[3].Value = (textBox3.Text + " = " + textBox8.Text);
                dataGridView1.Rows[3].Cells[3].Value = (textBox4.Text + " = " + textBox9.Text);
                dataGridView1.Rows[4].Cells[3].Value = (textBox5.Text + " = " + textBox10.Text);

                double Hhw1 = Int32.Parse(textBox6.Text);
                double Hhw2 = Int32.Parse(textBox7.Text);
                double Hhw3 = Int32.Parse(textBox8.Text);
                double Htst1 = Int32.Parse(textBox9.Text);
                double Htst2 = Int32.Parse(textBox10.Text);

                double Hhwavg = (Hhw1 + Hhw2 + Hhw3) / 300 * 100;
                double Htstavg = (Htst1 + Htst2) / 200 * 100;

                string Hhwcalc = Hhwavg.ToString("0.00");
                string Htstcalc = Htstavg.ToString("0.00");


                dataGridView1.Rows[6].Cells[3].Value = Hhwcalc + "%";
                dataGridView1.Rows[7].Cells[3].Value = Htstcalc + "%";


                double Hoverall = (Hhw1 + Hhw2 + Hhw3 + Htst1 + Htst2) / 500 * 100;



                string HFinal = Hoverall.ToString("0.00");



                if (Hoverall >= 90 && Hoverall <= 100)
                {
                    dataGridView1.Rows[8].Cells[3].Value = HFinal + "%" + " = A";
                }
                else

                if (Hoverall >= 80 && Hoverall <= 89)
                {
                    dataGridView1.Rows[8].Cells[3].Value = HFinal + "%" + " = B";
                }
                else

                if (Hoverall >= 70 && Hoverall <= 79)
                {
                    dataGridView1.Rows[8].Cells[3].Value = HFinal + "%" + " = C";
                }

                else

                if (Hoverall > 60 && Hoverall <= 69)
                {
                    dataGridView1.Rows[8].Cells[3].Value = HFinal + "%" + " = D";
                }

                else

                if (Hoverall >= 0 && Hoverall <= 59)
                {
                    dataGridView1.Rows[8].Cells[3].Value = HFinal + "%" + " = F";
                }

            }


        }



        public void OPEN_Click(object sender, EventArgs e)
        {

            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
            openFileDialog.FilterIndex = 2;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var workbook = ExcelFile.Load(openFileDialog.FileName);

                // From ExcelFile to DataGridView.
                DataGridViewConverter.ExportToDataGridView(workbook.Worksheets.ActiveWorksheet, this.dataGridView1, new ExportToDataGridViewOptions() { ColumnHeaders = true });
            }


        }



        private void button1_Click(object sender, EventArgs e)
        {

            string message, title;
            object numberGrade; message = "Enter in Grade Value: ";
            title = "Grade Calculator";


            numberGrade = Interaction.InputBox(message, title);
            
            if ((string)numberGrade == "")
            { 
            
            numberGrade = 0; //was created if nothign was inserted into the number input
            }
                int numberGradeint = int.Parse(string.Format("{0}", numberGrade)); //casting object to intenger

            if(numberGradeint >= 90 && numberGradeint <= 100)
            {
                Interaction.MsgBox("Great job, you got an A");
            }
            else

            if(numberGradeint >= 80 && numberGradeint <= 89)
            {
                Interaction.MsgBox("Good Work, you got a B ");
            }
            else

                if (numberGradeint >= 70 && numberGradeint <= 79)
            { 
                Interaction.MsgBox("Ok, you got a C "); 
            }

                else

                if(numberGradeint > 0 && numberGradeint <= 69)
            { 
                Interaction.MsgBox("Better Luck Next Time ");
            }
                else

                if(numberGradeint == 0)
            {
                Interaction.MsgBox("Nothing was inserted ");
            }

            }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            label1.Text = comboBox1.Text;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {
           
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string studentFile = studentName;

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var workbook = new ExcelFile();
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // From DataGridView to ExcelFile.
                DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridView1, new ImportFromDataGridViewOptions() { ColumnHeaders = true });

                workbook.Save(saveFileDialog.FileName);
            }
        }

    }
 }

    
