using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel_test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            //Книга.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            //Значения [y - строка,x - столбец]
            ObjWorkSheet.Cells[3, 1] = textBox1.Text;  // 3-строка, 1-ый столбец
            ObjWorkSheet.Cells[3, 2] = textBox2.Text;  // 3-строка, 2-ой столбец
            ObjWorkSheet.Cells[3, 3] = textBox3.Text;  // 3-строка, 3-ий столбец

            //Вызываем нашу созданную эксельку.
            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;
        }
    }
}
