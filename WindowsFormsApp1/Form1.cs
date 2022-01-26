using System;
using System.Drawing;
using System.Windows.Forms;
using MessagingToolkit.QRCode.Codec;
using System.Text.RegularExpressions;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            const string PREFIX = "&range=";
            const string pattern1 = "[a-zA-Z][0-9]+-[0-9]+";
            const string pattern2 = "[a-zA-Z][0-9]+-[a-zA-Z][0-9]+";
            string PrettyLink = "";
            string PrettyRange = "";
            string RawLink = textBox1.Text;
            string RawRange = textBox2.Text;
            var regex1 = Regex.Match(RawRange, pattern1);
            var regex2 = Regex.Match(RawRange, pattern2);
            bool succsessFlag = false;
            //Regex1 - A11-12
            if (regex1.Success)
            {
                string range = regex1.Value;
                const string ColumnPattern = "[a-zA-Z]";
                const string RowPattern = "[0-9]+";
                string column = Regex.Match(range, ColumnPattern).Value;
                var rows = Regex.Matches(range, RowPattern);
                foreach (var row in rows)
                {
                    PrettyRange += column + row + ":";
                }
                succsessFlag = true;
            }
            //Regex2 - A11-A12
            else if (regex2.Success)
            {
                const string CellPattern = "[a-zA-Z][0-9]+";
                var cells = Regex.Matches(RawRange, CellPattern);
                foreach (var cell in cells)
                {
                    PrettyRange += cell + ":";
                }
                succsessFlag = true;
            }
            else
            {
                MessageBox.Show("Неверный диапазон ячеек!");
                textBox2.Clear();
            }
            if (succsessFlag == true)
            {
                PrettyRange = PrettyRange.TrimEnd(':');
                PrettyLink += RawLink + PREFIX + PrettyRange;
                //Generate QR-code
                QRCodeEncoder encoder = new QRCodeEncoder();
                encoder.QRCodeScale = 4;
                Bitmap bitmap = encoder.Encode(PrettyLink);
                pictureBox1.Image = bitmap as Image;

                button2.Enabled = true;
                pictureBox1.Enabled = true;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog(); 
            save.Filter = "PNG|*.png|JPEG|*.jpg|GIF|*.gif|BMP|*.bmp";
            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                pictureBox1.Image.Save(save.FileName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            {
                SaveFileDialog save = new SaveFileDialog(); 
                save.Filter = "PNG|*.png|JPEG|*.jpg|GIF|*.gif|BMP|*.bmp"; 
                if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
                {
                    pictureBox1.Image.Save(save.FileName);
                }
            }
        }
    }
}
