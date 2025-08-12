using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;



namespace canexceldeneme
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


        private void SerialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string veri = serialPort1.ReadLine().Trim();

                //ekrana yaz
                this.Invoke(new Action(() =>
                {
                    if (veri.StartsWith("ID:"))
                    {
                        string[] parts = veri.Split(new[] { "ID:", "DATA:" }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length == 2)
                        {
                            string canId = parts[0].Trim();
                            string data = parts[1].Trim();

                          label1.Text = "CAN ID: " + canId;
                            label2.Text = "Data: " + data;

                            string[] dataBytes = data.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                            if (excelKayitAktif && worksheet != null)
                            {
                                worksheet.Cells[currentRow, 1] = DateTime.Now.ToString("HH:mm:ss");
                                worksheet.Cells[currentRow, 2] = canId;
                                for (int i = 0; i < 8; i++)
                                {
                                    worksheet.Cells[currentRow, 3 + i] = (i < dataBytes.Length) ? dataBytes[i] : "";
                                }
                                currentRow++;
                            }
                        }
                    }
                    else
                    {
                        label1.Text = "Geçersiz veri: " + veri;
                    }
                }));
            }
            catch (Exception ex)
            {
                this.Invoke(new Action(() =>
                {
                    MessageBox.Show("Okuma hatası: " + ex.Message);
                }));
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen)
            {
                serialPort1.PortName = comboBox1.Text;
                serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text);
                serialPort1.NewLine = "\n";
                serialPort1.ReadTimeout = 1000;
                serialPort1.Open();
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

        private void Form1_Load_1(object sender, EventArgs e)
        {
            serialPort1.DataReceived += SerialPort1_DataReceived;

            // Seri portları yükle
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(SerialPort.GetPortNames());
            if (comboBox1.Items.Count > 0)
                comboBox1.SelectedIndex = 0;

            // Baudrate’leri yükle
            comboBox2.Items.Clear();
            int[] baudRates = { 9600, 19200, 38400, 57600, 115200 };
            foreach (int rate in baudRates)
                comboBox2.Items.Add(rate.ToString());
            comboBox2.SelectedIndex = 4; // 115200

            // Elle giriş engelle
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;

            label1.Text = "CAN ID:";
            label2.Text = "Data:";
        }

        private void button3_Click_1(object sender, EventArgs e)
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
                    worksheet.Cells[1, 3] = "Byte 1";
                    worksheet.Cells[1, 4] = "Byte 2";
                    worksheet.Cells[1, 5] = "Byte 3";
                    worksheet.Cells[1, 6] = "Byte 4";
                    worksheet.Cells[1, 7] = "Byte 5";
                    worksheet.Cells[1, 8] = "Byte 6";
                    worksheet.Cells[1, 9] = "Byte 7";
                    worksheet.Cells[1, 10] = "Byte 8";
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
    }
}
