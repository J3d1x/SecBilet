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
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace _3bilet
{
    public partial class Form1 : Form
    {
        public double resultWater, resultLight, result;
        public string heat = "Выкл";
        public void ReceiptCount()
        {
            string date = DateTime.Now.ToString();
            int number;
            Random rnd = new Random();
            number = rnd.Next(10000, 100000);

            string path = @"C:\Users\Роман\Desktop\2 билет\3bilet\Квитанция\Квитанция-шаблон.docx";

            if (!File.Exists(path))
            {
                MessageBox.Show("Ошибка на правильность путя");
                return;
            }

            Word.Application app = new Word.Application();
            Word.Document doc =  app.Documents.Open(path);


            string newFileName = $"Квитанция_{DateTime.Now.ToString("yyyyMMddHHmmss")}.docx";
            string newFilePath = Path.Combine(Path.GetDirectoryName(path), newFileName);

            ReplaceText(doc, "[date]", date);
            ReplaceText(doc, "[number]", number.ToString());
            ReplaceText(doc, "[cold]", textBox1.Text);
            ReplaceText(doc, "[hot]", textBox1.Text);
            ReplaceText(doc, "[price]", textBox3.Text);
            ReplaceText(doc, "[present]", textBox4.Text);
            ReplaceText(doc, "[past]", textBox5.Text);
            ReplaceText(doc, "[overall]", result.ToString());
            ReplaceText(doc, "[heat]", heat.ToString());

            doc.SaveAs(newFilePath);
            app.Visible = true;

        }
        public void ReplaceText(Word.Document doc, string find, string Replace)
        {
            Word.Range range = doc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: find, ReplaceWith:Replace);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ReceiptCount();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public Form1()
        {
            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            try
            {
                if (textBox1.Text.Length == 0 || textBox2.Text.Length == 0 || textBox3.Text.Length == 0 || textBox4.Text.Length == 0 || textBox5.Text.Length == 0)
                {
                    MessageBox.Show("Заполните все поля!");
                    return;
                }
                if (textBox1.Text.Length > 16 || textBox2.Text.Length > 16 || textBox3.Text.Length > 16 || textBox4.Text.Length > 16 || textBox5.Text.Length > 16)
                {
                    MessageBox.Show("Число слишком огромное!");
                    return;
                }
                double hotWater = Convert.ToDouble(textBox1.Text);
                double coldWater = Convert.ToDouble(textBox2.Text);
                double priceLight = Convert.ToDouble(textBox3.Text);
                double presentLight = Convert.ToDouble(textBox4.Text);
                double pastLight = Convert.ToDouble(textBox5.Text);

                if (hotWater < 0 || coldWater < 0 || priceLight < 0 || presentLight < 0 || pastLight <0 )
                {
                    MessageBox.Show("Значение не может быть меньше нуля!");
                    return;
                }

                resultWater = hotWater * 12.76 + coldWater * 9.32;
                resultLight = priceLight * (presentLight - pastLight);
                result = resultLight + resultWater;

                if (checkBox1.Checked == true)
                {
                    heat = "3450.50";
                    result += 3450.50;
                }

                MessageBox.Show("Успех!");
            }
            catch (FormatException ex)
            {
              MessageBox.Show(ex.Message);
            }
            listBox1.Items.Add("Счет по воде: " + resultWater + "Р.");
            listBox1.Items.Add("Счет по свету: " + resultLight + "Р.");
            listBox1.Items.Add("Итог: " + result + "Р.");

        }
    }
}
