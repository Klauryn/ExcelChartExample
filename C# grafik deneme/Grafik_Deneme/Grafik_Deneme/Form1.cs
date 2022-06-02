using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms.DataVisualization.Charting;

namespace Grafik_Deneme
{
    public partial class Form1 : Form
    {
        public static Form1 instance;
        public static TextBox textbox;
        public string path = Directory.GetParent(System.Reflection.Assembly.GetExecutingAssembly().Location).FullName; // return the application.exe current folder
        public Form1()
        {
            textbox = textBox1;
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //Excel_CreateFile();
            //Excel_WriteData();
            Excel_OpenFile("Reference\\Açık Tip 7_5 Ton_3.csv", "B1");
            Excel_OpenFile("Reference\\Açık Tip 7_5 Ton_2.csv", "B2");
            //Excel_RowAndColumn();

            chart1.Titles.Add("Line Chart Example");
            /*
            chart1.Series["B1"].Points.AddXY("1", "20");
            chart1.Series["B1"].Points.AddXY("1", "25");
            chart1.Series["B1"].Points.AddXY("1", "30");

            chart1.Series["B2"].Points.AddXY("2", "44");
            chart1.Series["B2"].Points.AddXY("2", "57");
            chart1.Series["B2"].Points.AddXY("2", "23");

            chart1.Series["B3"].Points.AddXY("3", "16");
            chart1.Series["B3"].Points.AddXY("3", "77");
            chart1.Series["B3"].Points.AddXY("3", "96");
            */
        }
        public void Excel_RowAndColumn()
        {
            string fileName = Path.Combine(path, "Reference\\Açık Tip 7_5 Ton_3.csv"); //
            Excel _excel = new Excel(fileName, 1);
            Tuple<int, int> _tuple = _excel.RowsAndColumns();
            textBox1.Text = _tuple.Item2.ToString();
            _excel.Save();
            _excel.Close();
        }
        public void Excel_OpenFile(string data, string seriesName)
        {
            chart1.Series.Add(seriesName).ChartType = SeriesChartType.Point;
            chart1.Series.FindByName(seriesName).MarkerSize = 3;

            //chart1.ChartAreas[0].AxisY.Minimum = 0;
            chart1.ChartAreas[0].AxisY.Maximum = 6500;
            //chart1.ChartAreas[0].AxisX.Maximum = 500;
            chart1.ChartAreas[0].AxisX.Minimum = 230;


            string fileName = Path.Combine(path, data);
            string[] csvLines = File.ReadAllLines(fileName);
            var firstNames = new List<double>();
            var secondNames = new List<double>();

            for (int i = 1; i < csvLines.Length; i++)
            {
                string[] rowData = csvLines[i].Split(';');
                firstNames.Add(double.Parse(rowData[0]));
                secondNames.Add(double.Parse(rowData[1]));
            }

            string[,] rowAndCol = new string[firstNames.Count, secondNames.Count];

            //Bir Sheet(sayfa)' lik olacak şekilde dosya konumu verilerek içindeki veriler
            //çağırılmak üzere "_excel" adlı bir excel nesnesi oluşturuldu.
            Excel _excel = new Excel(fileName, 1);
            

            //_tuple değişkenine return edilen rows ve columns değerlerini atıyorum.
            Tuple<int, int> _tuple = _excel.RowsAndColumns();
            object[,] _string = new object[_tuple.Item1, _tuple.Item2];
            //exceldeki tüm satır ve sütunlardaki verileri çift boyutlu
            //_string[,] adlı değişkenimde saklıyorum.
            
            //combobox' a exceldeki tüm satır ve sütundaki verileri yazdırıyorum.
            for (int i = 1; i < csvLines.Length/*_tuple.Item1*/; i++)
            {
                chart1.Series[seriesName].Points.AddXY(firstNames[i-1], secondNames[i-1]);
            }
            _excel.Save();
            _excel.Close();
        }
        public void Excel_WriteData()
        {
            string fileName = Path.Combine(path, "Excel\\Created.xlsx");
            Excel _excel = new Excel(fileName, 1);
            _excel.WritetoCell(0, 0, "kuçu");
            _excel.Save();

            /*fileName = Path.Combine(path, "Excel\\WriteDataTest.xlsx");
            _excel.SaveAs(fileName);*/
            _excel.Close();
        }
        public void Excel_CreateFile()
        {
            Excel _excel = new Excel();
            _excel.CreateNewFile();
            string fileName = Path.Combine(path, "Excel\\Created.xlsx");
            _excel.SaveAs(fileName);
            _excel.Close();
        }

    }
}
