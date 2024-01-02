using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelMergeSelectedRows
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

           // Stream streamCsv = null;

            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Title = "Wybierz Plik";
            openFileDialog1.Filter = "CSV files *.csv|*csv";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {


                Form2 window2 = new Form2();
                window2.FileName = openFileDialog1.FileName;
                window2.Show();


            }
        }
    }
}