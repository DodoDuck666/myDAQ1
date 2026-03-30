using System;
using System.IO;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using NationalInstruments.DAQmx;

namespace myDAQ1
{
    public partial class Form1 : Form
    {
        // UI Controls
        private IContainer components = null;
        private ComboBox cmbDevice;
        private CheckBox chkSaveData;
        private NumericUpDown numAo1Min;
        private NumericUpDown numAo1Max;
        private NumericUpDown numAo1Steps;
        private NumericUpDown numInterval;
        private NumericUpDown numAo0Min;
        private NumericUpDown numAo0Max;
        private NumericUpDown numAo0Step;
        private Button btnStart;
        private Button btnStop;
        private Label lblAi0;
        private Label lblAi1;
        private Label lblIb;
        private Label lblIc;
        private Chart chartOutput;
        private Chart chartIbUbe;
        private Chart chartIcIb;
        private Chart chartIcUbe;
        private Timer timer1;

        // DAQ Tasks and Readers/Writers
        private Task taskAO0;
        private Task taskAO1;
        private Task taskAI0;
        private Task taskAI1;
        private AnalogSingleChannelWriter writerAO0;
        private AnalogSingleChannelWriter writerAO1;
        private AnalogSingleChannelReader readerAI0;
        private AnalogSingleChannelReader readerAI1;

        // Measurement State Variables
        private int currentStepAO1 = 0;
        private double currentVoltageAO0 = 0;
        private StreamWriter fileWriter;

        // Circuit Constants
        private const double R_B = 1000.0; // Base resistor (1k Ohm)
        private const double R_C = 1000.0; // Collector resistor (R0)

        public Form1()
        {
            InitializeComponent();
            LoadConnectedDevices();

            btnStart.Click += BtnStart_Click;
            btnStop.Click += BtnStop_Click;
            timer1.Tick += Timer1_Tick;
        }

        private void InitializeComponent()
        {
            this.components = new Container();

            // Initialize Controls
            this.cmbDevice = new ComboBox();
            this.chkSaveData = new CheckBox();
            this.numAo1Min = new NumericUpDown();
            this.numAo1Max = new NumericUpDown();
            this.numAo1Steps = new NumericUpDown();
            this.numInterval = new NumericUpDown();
            this.numAo0Min = new NumericUpDown();
            this.numAo0Max = new NumericUpDown();
            this.numAo0Step = new NumericUpDown();
            this.btnStart = new Button();
            this.btnStop = new Button();
            this.lblAi0 = new Label();
            this.lblAi1 = new Label();
            this.lblIb = new Label();
            this.lblIc = new Label();
            this.chartOutput = new Chart();
            this.chartIbUbe = new Chart();
            this.chartIcIb = new Chart();
            this.chartIcUbe = new Chart();
            this.timer1 = new Timer(this.components);

            Label lblInput1 = new Label() { Text = "AO1 Min (V)", Location = new Point(20, 20), AutoSize = true };
            Label lblInput2 = new Label() { Text = "AO1 Max (V)", Location = new Point(20, 60), AutoSize = true };
            Label lblInput3 = new Label() { Text = "AO1 Steps", Location = new Point(20, 100), AutoSize = true };
            Label lblInput4 = new Label() { Text = "Interval (ms)", Location = new Point(20, 140), AutoSize = true };
            Label lblInput5 = new Label() { Text = "AO0 Min (V)", Location = new Point(20, 180), AutoSize = true };
            Label lblInput6 = new Label() { Text = "AO0 Max (V)", Location = new Point(20, 220), AutoSize = true };
            Label lblInput7 = new Label() { Text = "AO0 Step (V)", Location = new Point(20, 260), AutoSize = true };
            Label lblDevice = new Label() { Text = "DAQ Device", Location = new Point(20, 300), AutoSize = true };

            ((ISupportInitialize)(this.numAo1Min)).BeginInit();
            ((ISupportInitialize)(this.numAo1Max)).BeginInit();
            ((ISupportInitialize)(this.numAo1Steps)).BeginInit();
            ((ISupportInitialize)(this.numInterval)).BeginInit();
            ((ISupportInitialize)(this.numAo0Min)).BeginInit();
            ((ISupportInitialize)(this.numAo0Max)).BeginInit();
            ((ISupportInitialize)(this.numAo0Step)).BeginInit();
            ((ISupportInitialize)(this.chartOutput)).BeginInit();
            ((ISupportInitialize)(this.chartIbUbe)).BeginInit();
            ((ISupportInitialize)(this.chartIcIb)).BeginInit();
            ((ISupportInitialize)(this.chartIcUbe)).BeginInit();
            this.SuspendLayout();

            // NumericUpDown Configurations
            this.numAo1Min.Location = new Point(120, 18);
            this.numAo1Min.Value = 0;
            this.numAo1Max.Location = new Point(120, 58);
            this.numAo1Max.Value = 8;
            this.numAo1Steps.Location = new Point(120, 98);
            this.numAo1Steps.Maximum = 10000;
            this.numAo1Steps.Value = 400;
            this.numInterval.Location = new Point(120, 138);
            this.numInterval.Maximum = 1000;
            this.numInterval.Value = 10;
            this.numAo0Min.Location = new Point(120, 178);
            this.numAo0Min.Value = 2;
            this.numAo0Max.Location = new Point(120, 218);
            this.numAo0Max.Value = 10;
            this.numAo0Step.DecimalPlaces = 1;
            this.numAo0Step.Location = new Point(120, 258);
            this.numAo0Step.Value = 1;

            // Device Combo and Checkbox shifted lower
            this.cmbDevice.Location = new Point(120, 298);
            this.cmbDevice.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbDevice.Size = new Size(120, 24);

            this.chkSaveData.Location = new Point(20, 335);
            this.chkSaveData.Text = "Save Data to CSV";
            this.chkSaveData.AutoSize = true;
            this.chkSaveData.Checked = false;

            // Buttons & Labels shifted lower
            this.btnStart.Location = new Point(20, 365);
            this.btnStart.Size = new Size(100, 40);
            this.btnStart.Text = "Start";
            this.btnStop.Location = new Point(130, 365);
            this.btnStop.Size = new Size(100, 40);
            this.btnStop.Text = "Stop";

            this.lblAi0.Location = new Point(20, 420);
            this.lblAi0.Text = "AI0: 0 V";
            this.lblAi0.AutoSize = true;
            this.lblAi1.Location = new Point(20, 450);
            this.lblAi1.Text = "AI1: 0 V";
            this.lblAi1.AutoSize = true;
            this.lblIb.Location = new Point(140, 420);
            this.lblIb.Text = "IB: 0 A";
            this.lblIb.AutoSize = true;
            this.lblIc.Location = new Point(140, 450);
            this.lblIc.Text = "IC: 0 A";
            this.lblIc.AutoSize = true;

            // Setup Charts with Titles and Formatting
            SetupChart(this.chartOutput, "AO1 Sawtooth", SeriesChartType.Line, 300, 20, "Time (Steps)", "Voltage (V)", "0", "0.00");
            SetupChart(this.chartIbUbe, "IB - UBE Curve", SeriesChartType.Point, 680, 20, "UBE (V)", "IB (A)", "0.00", "0.00E0");
            SetupChart(this.chartIcIb, "IC - IB Curve", SeriesChartType.Point, 300, 350, "IB (A)", "IC (A)", "0.00E0", "0.00E0");
            SetupChart(this.chartIcUbe, "IC - UBE Curve", SeriesChartType.Point, 680, 350, "UBE (V)", "IC (A)", "0.00", "0.00E0");

            // Form Configurations
            this.ClientSize = new Size(1060, 680);
            this.Controls.AddRange(new Control[] { this.cmbDevice, this.chkSaveData, this.numAo1Min, this.numAo1Max,
                                                   this.numAo1Steps, this.numInterval, this.numAo0Min, this.numAo0Max,
                                                   this.numAo0Step, this.btnStart, this.btnStop, this.lblAi0, this.lblAi1,
                                                   this.lblIb, this.lblIc, this.chartOutput, this.chartIbUbe,
                                                   this.chartIcIb, this.chartIcUbe, lblDevice, lblInput1, lblInput2,
                                                   lblInput3, lblInput4, lblInput5, lblInput6, lblInput7 });
            this.Name = "Form1";
            this.Text = "Transistor DAQ Measurement";

            ((ISupportInitialize)(this.numAo1Min)).EndInit();
            ((ISupportInitialize)(this.numAo1Max)).EndInit();
            ((ISupportInitialize)(this.numAo1Steps)).EndInit();
            ((ISupportInitialize)(this.numInterval)).EndInit();
            ((ISupportInitialize)(this.numAo0Min)).EndInit();
            ((ISupportInitialize)(this.numAo0Max)).EndInit();
            ((ISupportInitialize)(this.numAo0Step)).EndInit();
            ((ISupportInitialize)(this.chartOutput)).EndInit();
            ((ISupportInitialize)(this.chartIbUbe)).EndInit();
            ((ISupportInitialize)(this.chartIcIb)).EndInit();
            ((ISupportInitialize)(this.chartIcUbe)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void SetupChart(Chart chart, string title, SeriesChartType type, int x, int y, string xTitle, string yTitle, string xFormat, string yFormat)
        {
            ChartArea area = new ChartArea("ChartArea1");

            // Apply Annotations and Formatting
            area.AxisX.Title = xTitle;
            area.AxisY.Title = yTitle;
            area.AxisX.LabelStyle.Format = xFormat;
            area.AxisY.LabelStyle.Format = yFormat;

            Series series = new Series("Series1") { ChartArea = "ChartArea1", ChartType = type };
            chart.ChartAreas.Add(area);
            chart.Series.Add(series);
            chart.Titles.Add(title);
            chart.Location = new Point(x, y);
            chart.Size = new Size(350, 300);
        }

        private void LoadConnectedDevices()
        {
            try
            {
                string[] devices = DaqSystem.Local.Devices;
                if (devices.Length > 0)
                {
                    cmbDevice.Items.AddRange(devices);
                    cmbDevice.SelectedIndex = 0;
                }
                else
                {
                    cmbDevice.Items.Add("No devices found");
                    cmbDevice.SelectedIndex = 0;
                    btnStart.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not load DAQ devices: " + ex.Message);
            }
        }

        private void InitializeDAQ()
        {
            try
            {
                string dev = cmbDevice.SelectedItem.ToString();

                taskAO0 = new Task();
                taskAO0.AOChannels.CreateVoltageChannel($"{dev}/ao0", "", 0, 10, AOVoltageUnits.Volts);
                writerAO0 = new AnalogSingleChannelWriter(taskAO0.Stream);

                taskAO1 = new Task();
                taskAO1.AOChannels.CreateVoltageChannel($"{dev}/ao1", "", 0, 10, AOVoltageUnits.Volts);
                writerAO1 = new AnalogSingleChannelWriter(taskAO1.Stream);

                taskAI0 = new Task();
                taskAI0.AIChannels.CreateVoltageChannel($"{dev}/ai0", "", AITerminalConfiguration.Differential, -10.0, 10.0, AIVoltageUnits.Volts);
                readerAI0 = new AnalogSingleChannelReader(taskAI0.Stream);

                taskAI1 = new Task();
                taskAI1.AIChannels.CreateVoltageChannel($"{dev}/ai1", "", AITerminalConfiguration.Differential, -10.0, 10.0, AIVoltageUnits.Volts);
                readerAI1 = new AnalogSingleChannelReader(taskAI1.Stream);
            }
            catch (DaqException ex)
            {
                MessageBox.Show("DAQ Initialization Error: " + ex.Message);
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (cmbDevice.SelectedItem.ToString() == "No devices found")
            {
                MessageBox.Show("Please connect a DAQ device before starting.");
                return;
            }

            if (chkSaveData.Checked)
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                    sfd.Title = "Save Measurement Data";
                    sfd.FileName = "measurement_data.csv";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        fileWriter = new StreamWriter(sfd.FileName);
                        fileWriter.WriteLine("AO0_V,AO1_V,AI0_V,AI1_V,IB_A,IC_A");
                    }
                    else
                    {
                        return;
                    }
                }
            }
            else
            {
                fileWriter = null;
            }

            InitializeDAQ();

            currentStepAO1 = 0;
            currentVoltageAO0 = (double)numAo0Min.Value;

            chartOutput.Series[0].Points.Clear();
            chartIbUbe.Series[0].Points.Clear();
            chartIcIb.Series[0].Points.Clear();
            chartIcUbe.Series[0].Points.Clear();

            timer1.Interval = (int)numInterval.Value;
            timer1.Start();
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                double ao1Min = (double)numAo1Min.Value;
                double ao1Max = (double)numAo1Max.Value;
                int ao1Steps = (int)numAo1Steps.Value;

                double stepSizeAO1 = (ao1Max - ao1Min) / ao1Steps;
                double voltageAO1 = ao1Min + (currentStepAO1 * stepSizeAO1);

                writerAO0.WriteSingleSample(true, currentVoltageAO0);
                writerAO1.WriteSingleSample(true, voltageAO1);

                double voltageAI0 = readerAI0.ReadSingleSample();
                double voltageAI1 = readerAI1.ReadSingleSample();

                double ib = (voltageAO1 - voltageAI1) / R_B;
                double ic = (currentVoltageAO0 - voltageAI0) / R_C;

                lblAi0.Text = $"AI0: {voltageAI0:F3} V";
                lblAi1.Text = $"AI1: {voltageAI1:F3} V";
                lblIb.Text = $"IB: {ib:E3} A";
                lblIc.Text = $"IC: {ic:E3} A";

                chartOutput.Series[0].Points.AddY(voltageAO1);
                chartIbUbe.Series[0].Points.AddXY(voltageAI1, ib);
                chartIcIb.Series[0].Points.AddXY(ib, ic);
                chartIcUbe.Series[0].Points.AddXY(voltageAI1, ic);

                if (fileWriter != null)
                {
                    fileWriter.WriteLine($"{currentVoltageAO0:F4},{voltageAO1:F4},{voltageAI0:F4},{voltageAI1:F4},{ib:E6},{ic:E6}");
                }

                currentStepAO1++;
                if (currentStepAO1 > ao1Steps)
                {
                    currentStepAO1 = 0;
                    currentVoltageAO0 += (double)numAo0Step.Value;

                    if (currentVoltageAO0 > (double)numAo0Max.Value)
                    {
                        BtnStop_Click(this, EventArgs.Empty);
                    }
                }
            }
            catch (DaqException ex)
            {
                timer1.Stop();
                MessageBox.Show("DAQ Execution Error: " + ex.Message);
            }
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            timer1.Stop();

            if (fileWriter != null)
            {
                fileWriter.Close();
                fileWriter.Dispose();
            }

            try { writerAO0?.WriteSingleSample(true, 0); } catch { }
            try { writerAO1?.WriteSingleSample(true, 0); } catch { }

            DisposeDAQ();
        }

        private void DisposeDAQ()
        {
            taskAO0?.Dispose();
            taskAO1?.Dispose();
            taskAI0?.Dispose();
            taskAI1?.Dispose();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            BtnStop_Click(this, EventArgs.Empty);
            base.OnFormClosing(e);
        }
    }
}