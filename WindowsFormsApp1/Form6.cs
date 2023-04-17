using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
namespace WindowsFormsApp1
{
    public partial class Form6 : Form
    {
        public Form6(SqlConnection con)
        {
            InitializeComponent();
            this.con = con;
        }
        public SqlConnection con;
        private void Form6_Load(object sender, EventArgs e)
        {
            // Подсчет количества категорий
            String quertString = @"select count(kod_prodavca) from prodavec;";
            con.Open();
            SqlCommand table = new SqlCommand(quertString, con);
            SqlDataReader reader = table.ExecuteReader();
            reader.Read();
            Int16 k = Convert.ToInt16(reader[0]);
            con.Close();
            //Отправляется запрос в таблицу Категории
            quertString = @"select * from prodavec;";
            con.Open();
            table = new SqlCommand(quertString, con);
            reader = table.ExecuteReader();
            int[] id = new int[k];//Создается массив размером к, в который будут записываться Id категорий
            dataGridView1.ColumnCount = k;//Устанавливается количество столбцов dataGridView
            k = 0;//С этого момента переменная к определяет номер столбца для записи
            dataGridView1.Rows.Add();
            while (reader.Read())
            {
                dataGridView1[k, 0].Style.BackColor = Color.Red; //Устанавливаем заливку ячейки для наглядности
                id[k] = Convert.ToInt32(reader[0]);//Запоминаем ID в массив
                dataGridView1[k, 0].Value = reader[1].ToString();
                k++;
            }
            reader.Close();
            con.Close();
            for (int i = 0; i < id.Length; i++)//Для каждой категории
            {
                //Запрос номеров лотов с соответствующими продавцами
                String quertString2 = @"select nomer_lota from aukcuonnye_veshi where aukcuonnye_veshi.kod_prodavca='" + id[i] + "';";
                con.Open();
                SqlCommand table2 = new SqlCommand(quertString2, con);
                SqlDataReader reader2 = table2.ExecuteReader();
                int kn = 1;
                while (reader2.Read())
                {
                    if (i != 0)//Осуществляется проверка, необходимо ли добавлять строчки, т.к. количество книг у каждой категории разное
                    {
                        if (dataGridView1.RowCount < kn)//Если у данной категории количество книг больше, чем строк в таблице
                        { dataGridView1.Rows.Add(); }
                    }
                    dataGridView1[i, kn].Value = reader2[0].ToString();
                    kn++;
                }
                reader2.Close();
                con.Close();
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Установка фильтра для диалога сохранения файла
            saveFileDialog1.Filter = "Файлы Excel (*.xls; *.xlsx) | *.xls; *.xlsx";


            if (saveFileDialog1.ShowDialog() == DialogResult.OK)//Если пользователь сохранил документ
            {
                //Создание Excel документа
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                // Создание новой рабочей книги в этом документе
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // Создание нового листа в вышесозданной книге
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                // Устанавливает свойство видимости документа за программой. Можно установить false
                app.Visible = true;
                worksheet = workbook.ActiveSheet;// Определение значения объекта
                worksheet.Name = "Продавцы по названию товаров"; // Изменение имени рабочего листа
                // Заполнение Excel документа
                worksheet.Cells[1, 1] = "Продавцы по названию товаров:";
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[2, i] = dataGridView1[i - 1, 0].Value;
                    worksheet.Columns[i].ColumnWidth = 30;//Установление ширины столбцов
                    worksheet.Cells[2, i].Font.Color = Color.Red;//Установление цвета шрифта столбцов

                }
                for (int i = 1; i < dataGridView1.RowCount; i++)
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    { worksheet.Cells[i + 2, j + 1] = dataGridView1[j, i].Value; }
                // Сохраняет документ
                workbook.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 f1 = new Form1();
            f1.ShowDialog();
        }
    }
}
