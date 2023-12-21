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
using ExcelDataReader;

namespace MNK {
  public partial class Form1 : Form {
    private String fileName = string.Empty;
    private DataTableCollection tableCollection = null;

    public Form1() {
      InitializeComponent();
      MaximizeBox = false;
      FormBorderStyle = FormBorderStyle.Fixed3D;
      this.Text = "МНК";
    }

    public struct Dots {
      public double x;
      public double y;
      public Dots(double myX, double myY) {
        x = myX;
        y = myY;
      }
    }
    List<Dots> dots = new List<Dots>();


    void DrowDots() {
      for (int turn = 0; turn < dots.Count; ++turn) {
        this.chart1.Series[1].Points.AddXY(dots[turn].x, dots[turn].y);
      }
    }

    double FindMin() {
      double min = double.MaxValue;
      for (int turn = 0; turn < dots.Count; ++turn) {
        if (dots[turn].x < min) min = dots[turn].x;
      }
      return min;
    }
    double FindMax() {
      double max = double.MinValue;
      for (int turn = 0; turn < dots.Count; ++turn) {
        if (dots[turn].x > max) max = dots[turn].x;
      }
      return max;
    }

    string  Func2() {
      double sumXY = 0;
      double sumX = 0;
      double sumY = 0;
      double sumPowX = 0;
      for (int turn = 0; turn < dots.Count; ++turn) {
        sumXY += dots[turn].x * dots[turn].y;
        sumX += dots[turn].x;
        sumY += dots[turn].y;
        sumPowX += dots[turn].x * dots[turn].x;
      }
      double a = (dots.Count * sumXY - sumX * sumY) / (dots.Count * sumPowX - sumX * sumX);
      double b = (sumY - a * sumX) / dots.Count;
      for (int turn = Convert.ToInt32(FindMin()); turn <= Convert.ToInt32(FindMax()); ++turn) {
        double thisY = a * turn + b;
        this.chart1.Series[0].Points.AddXY(turn, thisY);
      }
      String answer = "";
      if (b >= 0) {
        answer = "y = " + Convert.ToString(Math.Round(a,3)) + " * x + " + Convert.ToString(Math.Round(b, 3)) + "\n";
      } else {
        answer = "y = " + Convert.ToString(Math.Round(a, 3)) + " * x " + Convert.ToString(Math.Round(b, 3)) + "\n";
      }
      return answer;
    }

    string Func3() {
      double sumPow4X = 0;
      double sumPow3X = 0;
      double sumPow2X = 0;
      double sumX = 0;
      double sumPow2XY = 0;
      double sumXY = 0;
      double sumY = 0;
      for (int turn = 0; turn < dots.Count; ++turn) {
        sumPow4X += Math.Pow(dots[turn].x, 4);
        sumPow3X += Math.Pow(dots[turn].x, 3);
        sumPow2X += Math.Pow(dots[turn].x, 2);
        sumX += dots[turn].x;
        sumPow2XY += Math.Pow(dots[turn].x, 2) * dots[turn].y;
        sumXY += dots[turn].x * dots[turn].y;
        sumY += dots[turn].y;
      }
      double del = sumPow4X * sumPow2X * dots.Count + sumPow3X * sumX * sumPow2X + sumPow3X * sumX * sumPow2X - Math.Pow(sumPow2X, 3) - sumPow3X * sumPow3X * dots.Count - sumX * sumX * sumPow4X;
      double del1 = sumPow2XY * sumPow2X * dots.Count + sumPow3X * sumX * sumY + sumXY * sumX * sumPow2X - sumY * sumPow2X * sumPow2X - sumXY * sumPow3X * dots.Count - sumX * sumX * sumPow2XY;
      double del2 = sumPow4X * sumXY * dots.Count + sumPow3X * sumY * sumPow2X + sumPow2XY * sumX * sumPow2X - sumPow2X * sumPow2X * sumXY - sumPow2XY * sumPow3X * dots.Count - sumY * sumX * sumPow4X;
      double del3 = sumPow4X * sumPow2X * sumY + sumPow3X * sumX * sumPow2XY + sumPow3X * sumXY * sumPow2X - sumPow2X * sumPow2X * sumPow2XY - sumPow3X * sumPow3X * sumY - sumX * sumXY * sumPow4X;
      double a = del1 / del;
      double b = del2 / del;
      double c = del3 / del;
      for (int turn = Convert.ToInt32(FindMin()); turn <= Convert.ToInt32(FindMax()); ++turn) {
        double thisY = a * turn * turn + b * turn + c;
        this.chart1.Series[2].Points.AddXY(turn, thisY);
      }
      String myB = "";
      String myC = "";
      string myA = Convert.ToString(Math.Round(a, 3)) + " * x^2 ";
      if (b < 0) {
        myB = Convert.ToString(Math.Round(b, 3)) + " * x ";
      } else {
        myB = "+ " + Convert.ToString(Math.Round(b, 3)) + " * x ";
      }
      if (c < 0) {
        myC = Convert.ToString(Math.Round(c, 3));
      } else {
        myC = "+ " + Convert.ToString(Math.Round(c, 3));
      }
      String answer = "y = " + myA + myB + myC;
      return answer;
    }

    private void Form1_Load(object sender, EventArgs e) {

    }

    private void рассчитатьToolStripMenuItem_Click(object sender, EventArgs e) {
      try {
        this.chart1.Series[0].Points.Clear();
        this.chart1.Series[1].Points.Clear();
        this.chart1.Series[2].Points.Clear();
        for (int turn = 0; turn < dataGridView1.RowCount - 1; ++turn) {
           dots.Add(new Dots(Convert.ToDouble(dataGridView1[0, turn].Value), Convert.ToDouble(dataGridView1[1, turn].Value)));
          }
        DrowDots();
        label1.Text = "Найденные уравнения:\n";
        label1.Text += Func2();
        label1.Text += Func3();
      } catch {
        Mistake();
      }
    }

    private void заполнитьАвтоматическиToolStripMenuItem_Click(object sender, EventArgs e) {
      var rand = new Random();
      for (int turn = 0; turn < 10; ++turn) {
        DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
        row.Cells[0].Value = Convert.ToDouble(rand.Next(-100, 100));
        row.Cells[1].Value = Convert.ToDouble(rand.Next(-100, 100));
        dataGridView1.Rows.Add(row);
      }
    }

    void Mistake() {
      label1.Text = "Ошибка";
    }

    private void очиститьToolStripMenuItem_Click(object sender, EventArgs e) {
      label1.Text = "";
      this.chart1.Series[0].Points.Clear();
      this.chart1.Series[1].Points.Clear();
      this.chart1.Series[2].Points.Clear();
      dataGridView1.DataSource = null;
      dataGridView1.Rows.Clear();
      dots.Clear();
    }

    private void заToolStripMenuItem_Click(object sender, EventArgs e) {
      заполнитьАвтоматическиToolStripMenuItem.Visible = false;
      dataGridView1.Columns.Clear();
      try {
        DialogResult res = openFileDialog1.ShowDialog();
        if (res == DialogResult.OK) {
          fileName = openFileDialog1.FileName;
          Text = fileName;
          OpenExcelFile(fileName);
        } else {
          throw new Exception("Файл не выбран");
        }
      } catch(Exception ex) {
        MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    private void OpenExcelFile(string path) {
      FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
      IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
      DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration() {
        ConfigureDataTable = (x) => new ExcelDataTableConfiguration() {
          UseHeaderRow = true
        }
      });
      tableCollection = db.Tables;
      toolStripComboBox1.Items.Clear();
      foreach(DataTable table in tableCollection) {
        toolStripComboBox1.Items.Add(table.TableName);
      }
      toolStripComboBox1.SelectedIndex = 0;
    }

    private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e) {
      DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];
      dataGridView1.DataSource = table;
    }
  }
}