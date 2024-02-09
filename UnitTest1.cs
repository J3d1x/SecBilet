using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using _3bilet;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Form1 form = new Form1();

            form.textBox1.Text = "12354,42403034600475E+1310 ";
            form.textBox2.Text = "12354,42403034600475E+1310 ";
            form.textBox3.Text = "12354,42403034600475E+1310 ";
            form.textBox4.Text = "12354,42403034600475E+1310 ";
            form.textBox5.Text = "12354,42403034600475E+1310 ";
            form.button1_Click(null, null);


        }
        [TestMethod]
        public void TestMethod2()
        {
            Form1 form = new Form1();

            form.textBox1.Text = "-1";
            form.textBox2.Text = "-5";
            form.textBox3.Text = "1";
            form.textBox4.Text = "2";
            form.textBox5.Text = "3";
            form.button1_Click(null, null);

        }

        [TestMethod]
        public void TestMethod3()
        {
            Form1 form = new Form1();

            form.textBox1.Text = null;
            form.textBox2.Text = "1";
            form.textBox3.Text = "2";
            form.textBox4.Text = "3";
            form.textBox5.Text = "3";
            form.button1_Click(null, null);

        }
    }

}
