using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelExamle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
         //Открываем файл Экселя
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                //Очищаем от старого текста окно вывода.
                richTextBox1.Clear();

                //Выбираем первые сто записей из столбца.
                for (int i = 1; i < 101; i++)
                {
                    //Выбираем область таблицы. (в нашем случае просто ячейку)
                    Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range(textBox1.Text + i.ToString(), textBox1.Text + i.ToString());
                    //Добавляем полученный из ячейки текст.
                    richTextBox1.Text = richTextBox1.Text + range.Text.ToString()+"\n";
                    Application.DoEvents();
                }

                //Удаляем приложение (выходим из экселя) - ато будет висеть в процессах!
                ObjExcel.Quit();
            }
        }
        
    }
}
