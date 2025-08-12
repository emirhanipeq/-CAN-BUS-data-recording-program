using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Windows.Forms;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection.Emit;

namespace CAN_Excel_Logger
{
    public partial class Form1 : Form
    {
        Excel.Application excelApp;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        int currentRow = 2;
        bool excelKayitAktif = false;
        string excelPath = "C:\\veri_kaydi.xlsx";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Items.AddRange(SerialPort.GetPortNames());
            if (comboBox1.Items.Count > 0)
                comboBox1.SelectedIndex = 0;

            int[] baudRates = { 9600, 19200, 38400, 57600, 115200 };
            foreach (int rate in baudRates)
                comboBox2.Items.Add(rate.ToString());
            comboBox2.SelectedIndex = 0; // 115200

            label1.Text = "CAN ID:";
            label2.Text = "Data:";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen)
            {
                serialPort1.PortName = comboBox1.Text;
                serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text);
                serialPort1.Open();
                timer1.Start();
                MessageBox.Show("Bağlantı açıldı.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            if (serialPort1.IsOpen)
                serialPort1.Close();
            MessageBox.Show("Bağlantı kapatıldı.");
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!serialPort1.IsOpen)
                    return;

                string veri = serialPort1.ReadLine().Trim();
                // Örnek veri: "ID:0F6 DATA:8E 87 32 FA 26 8E BE 86"

                if (veri.StartsWith("ID:"))
                {
                    string[] parts = veri.Split(new[] { "ID:", "DATA:" }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length == 2)
                    {
                        string canId = parts[0].Trim();
                        string data = parts[1].Trim();

                        label1.Text = "CAN ID: " + canId;
                        label2.Text = "Data: " + data;

                        if (excelKayitAktif && worksheet != null)
                        {
                            worksheet.Cells[currentRow, 1] = DateTime.Now.ToString("HH:mm:ss");
                            worksheet.Cells[currentRow, 2] = canId;
                            worksheet.Cells[currentRow, 3] = data;
                            currentRow++;
                        }
                    }
                }
                else
                {
                    label1.Text = "Geçersiz veri: " + veri;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Okuma hatası: " + ex.Message);
                timer1.Stop();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!excelKayitAktif)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "Excel dosyasını kaydet";
                saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                saveFileDialog.FileName = "veri_kaydi.xlsx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelPath = saveFileDialog.FileName;
                    excelApp = new Excel.Application();
                    workbook = excelApp.Workbooks.Add();
                    worksheet = workbook.Worksheets[1];
                    worksheet.Cells[1, 1] = "Zaman";
                    worksheet.Cells[1, 2] = "CAN ID";
                    worksheet.Cells[1, 3] = "Data";
                    currentRow = 2;
                    excelKayitAktif = true;
                    MessageBox.Show("Excel kaydı başlatıldı.");
                }
            }
            else
            {
                workbook.SaveAs(excelPath);
                workbook.Close(false);
                excelApp.Quit();
                excelApp = null;
                workbook = null;
                worksheet = null;
                excelKayitAktif = false;
                MessageBox.Show("Excel'e veri kaydı durduruldu ve dosya kaydedildi.");
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (serialPort1.IsOpen)
                serialPort1.Close();
            if (excelKayitAktif)
            {
                workbook.SaveAs(excelPath);
                workbook.Close(false);
                excelApp.Quit();
            }
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {

        }
    }
}
