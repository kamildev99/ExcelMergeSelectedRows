using CsvHelper;
using CsvHelper.Configuration;
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
using CheckBox = System.Windows.Forms.CheckBox;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelMergeSelectedRows
{
    public partial class Form2 : Form
    {

        //zmienne
       // public var newSetsRecords = null;
        List<Tuple<Label, CheckBox>> tupleControls = new List<Tuple<Label, CheckBox>>();
        List<Dictionary<string, object>> newSetsRecords =  new List<Dictionary<string, object>>();
        Dictionary<string, string> columnValues = new Dictionary<string, string>();


        public String FileName { get; set; } = "";
        public Form2()
        {

            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            try
            {
                

                ////////////// TU NOWE
                ///
                int top = 20;
                var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = true, // Ustawienie, czy plik CSV ma nagłówek
                    MissingFieldFound = null, 
                    Delimiter = CsvSeperatorDetector.DetectSeparator(FileName).ToString(),
                };

                
                this.textBox1.Text += "\n UWAGA NOWE  ---------------- \n ";
                using (var readerN = new StreamReader(File.OpenRead(FileName)))
                using (var csv = new CsvReader(readerN, csvConfig))
                {
                    var records = new List<Dictionary<string, object>>();
                    int howMany = 0;

                    csv.Read();
                    csv.ReadHeader();

                    var headersNew = csv.HeaderRecord;
                    //var recordsNew = csv.g;

                    foreach (var header in headersNew)
                    {
                      //  record[header] = csv.GetField(header);

                        //this.textBox1.AppendText(header);
                        Label labelTmp = new Label();
                        labelTmp.Text = header;
                        labelTmp.Width = 200;
                        labelTmp.Location = new System.Drawing.Point(50, top);

                        CheckBox checkBoxTmp = new CheckBox();
                        checkBoxTmp.Text = "Grupuj";
                        checkBoxTmp.Location = new System.Drawing.Point(300, top);

                        
                        tupleControls.Add(new Tuple<Label, CheckBox>(labelTmp, checkBoxTmp));
                        

                        this.Controls.Add(labelTmp);
                        this.Controls.Add(checkBoxTmp);
                        top = top + 30;

                        howMany++;
                    }

                    while (csv.Read())
                    {
                        var record = new Dictionary<string, object>();


                        foreach (var header in headersNew)
                        {
                            var recordTmp = csv.GetField(header);
                            columnValues.Add(header, recordTmp.ToString());
                        }



                            records.Add(record);
                        this.newSetsRecords.Add(record);
                    }
                    this.textBox1.AppendText(howMany.ToString());

                   
                    //foreach(column in records.)
                }
/////////////////////////////////
//////////////////////////////
            }
            catch (Exception ex)
            {
                MessageBox.Show("Wystąpił błąd podczas otwierania pliku Excel: " + ex.Message);
                this.Close();
            }
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {


           // var resultExcel = from records in newSetsRecords;
            
        }
    }
}
