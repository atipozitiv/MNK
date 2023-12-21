namespace MNK {
  partial class Form1 {
    /// <summary>
    /// Обязательная переменная конструктора.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Освободить все используемые ресурсы.
    /// </summary>
    /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
    protected override void Dispose(bool disposing) {
      if (disposing && (components != null)) {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Код, автоматически созданный конструктором форм Windows

    /// <summary>
    /// Требуемый метод для поддержки конструктора — не изменяйте 
    /// содержимое этого метода с помощью редактора кода.
    /// </summary>
    private void InitializeComponent() {
      System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
      System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
      System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
      System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
      this.dataGridView1 = new System.Windows.Forms.DataGridView();
      this.x = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.y = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.menuStrip1 = new System.Windows.Forms.MenuStrip();
      this.действияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.рассчитатьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.заполнитьАвтоматическиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.очиститьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.заToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
      this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
      this.label1 = new System.Windows.Forms.Label();
      this.toolStrip1 = new System.Windows.Forms.ToolStrip();
      this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
      this.toolStripComboBox1 = new System.Windows.Forms.ToolStripComboBox();
      this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
      this.menuStrip1.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
      this.toolStrip1.SuspendLayout();
      this.SuspendLayout();
      // 
      // dataGridView1
      // 
      this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.x,
            this.y});
      this.dataGridView1.Location = new System.Drawing.Point(12, 28);
      this.dataGridView1.Name = "dataGridView1";
      this.dataGridView1.Size = new System.Drawing.Size(245, 435);
      this.dataGridView1.TabIndex = 0;
      // 
      // x
      // 
      this.x.HeaderText = "x";
      this.x.Name = "x";
      // 
      // y
      // 
      this.y.HeaderText = "y";
      this.y.Name = "y";
      // 
      // menuStrip1
      // 
      this.menuStrip1.Dock = System.Windows.Forms.DockStyle.Bottom;
      this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.действияToolStripMenuItem});
      this.menuStrip1.Location = new System.Drawing.Point(0, 504);
      this.menuStrip1.Name = "menuStrip1";
      this.menuStrip1.Size = new System.Drawing.Size(862, 24);
      this.menuStrip1.TabIndex = 1;
      this.menuStrip1.Text = "menuStrip1";
      // 
      // действияToolStripMenuItem
      // 
      this.действияToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.рассчитатьToolStripMenuItem,
            this.заполнитьАвтоматическиToolStripMenuItem,
            this.очиститьToolStripMenuItem,
            this.заToolStripMenuItem});
      this.действияToolStripMenuItem.Name = "действияToolStripMenuItem";
      this.действияToolStripMenuItem.Size = new System.Drawing.Size(70, 20);
      this.действияToolStripMenuItem.Text = "Действия";
      // 
      // рассчитатьToolStripMenuItem
      // 
      this.рассчитатьToolStripMenuItem.Name = "рассчитатьToolStripMenuItem";
      this.рассчитатьToolStripMenuItem.Size = new System.Drawing.Size(239, 22);
      this.рассчитатьToolStripMenuItem.Text = "Рассчитать";
      this.рассчитатьToolStripMenuItem.Click += new System.EventHandler(this.рассчитатьToolStripMenuItem_Click);
      // 
      // заполнитьАвтоматическиToolStripMenuItem
      // 
      this.заполнитьАвтоматическиToolStripMenuItem.Name = "заполнитьАвтоматическиToolStripMenuItem";
      this.заполнитьАвтоматическиToolStripMenuItem.Size = new System.Drawing.Size(239, 22);
      this.заполнитьАвтоматическиToolStripMenuItem.Text = "Добавить 10 случайных точек";
      this.заполнитьАвтоматическиToolStripMenuItem.Click += new System.EventHandler(this.заполнитьАвтоматическиToolStripMenuItem_Click);
      // 
      // очиститьToolStripMenuItem
      // 
      this.очиститьToolStripMenuItem.Name = "очиститьToolStripMenuItem";
      this.очиститьToolStripMenuItem.Size = new System.Drawing.Size(239, 22);
      this.очиститьToolStripMenuItem.Text = "Очистить";
      this.очиститьToolStripMenuItem.Click += new System.EventHandler(this.очиститьToolStripMenuItem_Click);
      // 
      // заToolStripMenuItem
      // 
      this.заToolStripMenuItem.Name = "заToolStripMenuItem";
      this.заToolStripMenuItem.Size = new System.Drawing.Size(239, 22);
      this.заToolStripMenuItem.Text = "Загрузить из EXCEL";
      this.заToolStripMenuItem.Click += new System.EventHandler(this.заToolStripMenuItem_Click);
      // 
      // chart1
      // 
      chartArea1.Name = "ChartArea1";
      this.chart1.ChartAreas.Add(chartArea1);
      this.chart1.Location = new System.Drawing.Point(295, 80);
      this.chart1.Name = "chart1";
      series1.BorderWidth = 2;
      series1.ChartArea = "ChartArea1";
      series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
      series1.Color = System.Drawing.Color.Fuchsia;
      series1.CustomProperties = "LabelStyle=Top";
      series1.EmptyPointStyle.MarkerSize = 10;
      series1.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      series1.MarkerColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
      series1.MarkerSize = 10;
      series1.Name = "function2";
      series2.BorderWidth = 10;
      series2.ChartArea = "ChartArea1";
      series2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
      series2.Color = System.Drawing.Color.Blue;
      series2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
      series2.MarkerBorderWidth = 5;
      series2.MarkerSize = 6;
      series2.Name = "dots";
      series3.BorderWidth = 2;
      series3.ChartArea = "ChartArea1";
      series3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
      series3.Color = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
      series3.Name = "function3";
      this.chart1.Series.Add(series1);
      this.chart1.Series.Add(series2);
      this.chart1.Series.Add(series3);
      this.chart1.Size = new System.Drawing.Size(555, 421);
      this.chart1.TabIndex = 2;
      this.chart1.Text = "chart1";
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(292, 28);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(0, 13);
      this.label1.TabIndex = 3;
      // 
      // toolStrip1
      // 
      this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel1,
            this.toolStripComboBox1});
      this.toolStrip1.Location = new System.Drawing.Point(0, 0);
      this.toolStrip1.Name = "toolStrip1";
      this.toolStrip1.Size = new System.Drawing.Size(862, 25);
      this.toolStrip1.TabIndex = 4;
      this.toolStrip1.Text = "toolStrip1";
      // 
      // toolStripLabel1
      // 
      this.toolStripLabel1.Name = "toolStripLabel1";
      this.toolStripLabel1.Size = new System.Drawing.Size(33, 22);
      this.toolStripLabel1.Text = "Лист";
      // 
      // toolStripComboBox1
      // 
      this.toolStripComboBox1.Name = "toolStripComboBox1";
      this.toolStripComboBox1.Size = new System.Drawing.Size(121, 25);
      this.toolStripComboBox1.SelectedIndexChanged += new System.EventHandler(this.toolStripComboBox1_SelectedIndexChanged);
      // 
      // openFileDialog1
      // 
      this.openFileDialog1.FileName = "openFileDialog1";
      this.openFileDialog1.Filter = "Excel|*.xlsx";
      // 
      // Form1
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(862, 528);
      this.Controls.Add(this.toolStrip1);
      this.Controls.Add(this.label1);
      this.Controls.Add(this.chart1);
      this.Controls.Add(this.dataGridView1);
      this.Controls.Add(this.menuStrip1);
      this.MainMenuStrip = this.menuStrip1;
      this.Name = "Form1";
      this.Text = "Form1";
      this.Load += new System.EventHandler(this.Form1_Load);
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
      this.menuStrip1.ResumeLayout(false);
      this.menuStrip1.PerformLayout();
      ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
      this.toolStrip1.ResumeLayout(false);
      this.toolStrip1.PerformLayout();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.DataGridView dataGridView1;
    private System.Windows.Forms.DataGridViewTextBoxColumn x;
    private System.Windows.Forms.DataGridViewTextBoxColumn y;
    private System.Windows.Forms.MenuStrip menuStrip1;
    private System.Windows.Forms.ToolStripMenuItem действияToolStripMenuItem;
    private System.Windows.Forms.ToolStripMenuItem рассчитатьToolStripMenuItem;
    private System.Windows.Forms.ToolStripMenuItem заполнитьАвтоматическиToolStripMenuItem;
    private System.ComponentModel.BackgroundWorker backgroundWorker1;
    private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.ToolStripMenuItem очиститьToolStripMenuItem;
    private System.Windows.Forms.ToolStripMenuItem заToolStripMenuItem;
    private System.Windows.Forms.ToolStrip toolStrip1;
    private System.Windows.Forms.ToolStripLabel toolStripLabel1;
    private System.Windows.Forms.ToolStripComboBox toolStripComboBox1;
    private System.Windows.Forms.OpenFileDialog openFileDialog1;
  }
}

