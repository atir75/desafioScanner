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
using WIA;

namespace ScannerDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ListScanners();

            textBox1.Text = Path.GetTempPath();
            comboBox1.SelectedIndex = 1;

        }

        private void ListScanners()
        {
            listBox1.Items.Clear();

            var deviceManager = new DeviceManager();
            for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
            {
                if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                {
                    continue;
                }

                listBox1.Items.Add(
                    new Scanner(deviceManager.DeviceInfos[i])
                );
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Task.Factory.StartNew(StartScanning).ContinueWith(result => TriggerScan());
        }

        private void TriggerScan()
        {
            Console.WriteLine("Deu certo!!");
        }

        public void StartScanning()
        {
            Scanner device = null;

            this.Invoke(new MethodInvoker(delegate ()
            {
                device = listBox1.SelectedItem as Scanner;
            }));

            if (device == null)
            {
                MessageBox.Show("Selecione um scanner",
                                "Aviso",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }else if(String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Digite um nome",
                                "Aviso",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ImageFile image = new ImageFile();
            string imageExtension = "";

            this.Invoke(new MethodInvoker(delegate ()
            {
                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        image = device.ScanImage(WIA.FormatID.wiaFormatPNG);
                        imageExtension = ".png";
                        break;
                    case 1:
                        image = device.ScanImage(WIA.FormatID.wiaFormatJPEG);
                        imageExtension = ".jpeg";
                        break;
                    case 2:
                        image = device.ScanImage(WIA.FormatID.wiaFormatBMP);
                        imageExtension = ".bmp";
                        break;
                    case 3:
                        image = device.ScanImage(WIA.FormatID.wiaFormatGIF);
                        imageExtension = ".gif";
                        break;
                    case 4:
                        image = device.ScanImage(WIA.FormatID.wiaFormatTIFF);
                        imageExtension = ".tiff";
                        break;
                }
            }));
            
            
            // Save the image
            var path = Path.Combine(textBox1.Text, textBox2.Text + imageExtension);

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            image.SaveFile(path);

            pictureBox1.Image = new Bitmap(path);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            DialogResult result = folderDlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                textBox1.Text = folderDlg.SelectedPath;
            }
        }

    }
}
