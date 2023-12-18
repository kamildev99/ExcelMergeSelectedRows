using CsvHelper;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Range = Microsoft.Office.Interop.Excel.Range;
using Label = System.Windows.Forms.Label;
using TextBox = System.Windows.Forms.TextBox;

namespace ExcelMergeSelectedRows
{
    public partial class Form2 : Form
    {

        public String FileName { get; set; } = "";
        public Form2()
        {

            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            try
            {

                var reader = new StreamReader(File.OpenRead(FileName));
                var lineFirst = reader.ReadLine();


                //first line which is not null
                while (lineFirst == null)
                { 
                    lineFirst = reader.ReadLine();
                }

                char delimiter = CsvSeperatorDetector.DetectSeparator(FileName);

                //labels
                //textboxes
                List<TextBox> textBoxes = new List<TextBox>();
               
                int top = 20;
                //handle this, can be null values
                foreach (string value in lineFirst.Split(delimiter))
                {
                     
                    Label labelTmp = new Label();
                    labelTmp.Text = value;
                    labelTmp.Location = new System.Drawing.Point(50, top);        
                    this.Controls.Add(labelTmp);
                    top = top + 30;
                }

                

                List<String> valuesLineFirst = lineFirst.Split(CsvSeperatorDetector.DetectSeparator(FileName)).ToList();
                this.textBox1.Text += String.Join(",", valuesLineFirst.ToArray()) + "\n \n \n new line \n \n ";
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();



                    this.textBox1.Text += line.ToString() + "\n";


                }

                this.textBox1.Text += "\n\n\n Speparator to: ";
                this.textBox1.Text += "\n\n\n" + CsvSeperatorDetector.DetectSeparator(FileName);


                //MessageBox.Show("Udało się jestem w form2 !!! i przekazana wartość: " + FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Wystąpił błąd podczas otwierania pliku Excel: " + ex.Message);
                this.Close();
            }
        }
    }
}
