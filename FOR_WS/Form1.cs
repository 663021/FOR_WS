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

namespace FOR_WS
{
    public partial class Form1 : Form
    {
        SqlConnection SqlConnection;

        public Form1()
        {
            InitializeComponent();
        }

        public string connectionString = @"Data Source=" + SystemInformation.ComputerName + @"\SQLEXPRESS;Initial Catalog=PODGOTOVKA;Integrated Security=True";

        private async void Form1_Load(object sender, EventArgs e)
        {
            SqlConnection = new SqlConnection(connectionString);
            await SqlConnection.OpenAsync();

            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("SELECT * FROM [Ингредиент]", SqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    //comboBox2.Items.Add(Convert.ToString(sqlReader["Артикул"]));
                }

                await command.ExecuteNonQueryAsync();
            }
            catch
            {

            }

            // Изменение

            /* SqlCommand command = new SqlCommand("UPDATE [Ингредиент] SET [Наименование]=@name,[Еденица измерения]=@ed,[Количество]=@kol,[Основной поставщик]=@Osn,[Тип Ингредиента]=@type,[Закупочная цена]=@price,[ГОСТ]=@GOST,[Фасовка]=@fas,[Характеристика]=@har WHERE [Артикул]=@art", SqlConnection);
            command.Parameters.AddWithValue("@name", textBox8.Text);
            command.Parameters.AddWithValue("@ed", textBox7.Text);
            command.Parameters.AddWithValue("@kol", textBox12.Text);
            command.Parameters.AddWithValue("@Osn", textBox11.Text);
            command.Parameters.AddWithValue("@type", textBox10.Text);
            command.Parameters.AddWithValue("@price", textBox4.Text);
            command.Parameters.AddWithValue("@GOST", textBox5.Text);
            command.Parameters.AddWithValue("@har", textBox9.Text);
            command.Parameters.AddWithValue("@fas", textBox6.Text);
            command.Parameters.AddWithValue("@art", comboBox4.Text);

            await command.ExecuteNonQueryAsync(); */

            // Удаление и диалогове окно

            /*SqlCommand command = new SqlCommand("DELETE [Ингредиент] WHERE [Артикул]=@art", SqlConnection);
            command.Parameters.AddWithValue("@art", comboBox4.Text);

            DialogResult result = MessageBox.Show("Вы действительно хотите удалить запись?", "Сообщение", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);

            if (result == DialogResult.Yes)
            {
                await command.ExecuteNonQueryAsync();
                MessageBox.Show("Удаление успешно произведено");
                Director o = new Director();
                o.Show();
                this.Hide();
            }*/

            // Добавление

            /*SqlCommand commad = new SqlCommand("INSERT INTO [Оборудование] (Маркировка,[Тип оборудования],[Степень износа],Поставщик,[Дата приобретения],Количество)VALUES(@m,@t,@s,@p,@d,@k)", SqlConnection);
            commad.Parameters.AddWithValue("@m", textBox1.Text);
            commad.Parameters.AddWithValue("@t", comboBox1.Text);
            commad.Parameters.AddWithValue("@s", comboBox2.Text);
            commad.Parameters.AddWithValue("@p", comboBox3.Text);
            commad.Parameters.AddWithValue("@d", textBox2.Text);
            commad.Parameters.AddWithValue("@k", textBox3.Text);

            await commad.ExecuteNonQueryAsync();*/

            // Относительный путь

            /*pictureBox2.ImageLocation = Environment.CurrentDirectory + @"\" + comboBox5.Text + ".png";
            pictureBox2.SizeMode = PictureBoxSizeMode.Zoom;*/

            // Отчеты

            /*Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Columns.ColumnWidth = 15;

            ExcelApp.Cells[1, 1] = "Остатки Фурнитуры";
            ExcelApp.Cells[3, 1] = "Партия";
            ExcelApp.Cells[3, 2] = "Артикул фурнитуры";
            ExcelApp.Cells[3, 3] = "Количество";

            for (int i = 0; i < dataGridView3.RowCount - 1; i++)
            {
                ExcelApp.Cells[i + 4, 1] = (dataGridView3[0, i].Value).ToString();
                ExcelApp.Cells[i + 4, 2] = (dataGridView3[1, i].Value).ToString();
                ExcelApp.Cells[i + 4, 3] = (dataGridView3[2, i].Value).ToString();
            }

            ExcelApp.Cells[dataGridView3.RowCount + 4, 1] = "Остатки Ткани";
            ExcelApp.Cells[dataGridView3.RowCount + 6, 1] = "Артикул ткани";
            ExcelApp.Cells[dataGridView3.RowCount + 6, 2] = "Рулон";

            for (int i = 0; i < dataGridView4.RowCount - 1; i++)
            {
                ExcelApp.Cells[i + dataGridView3.RowCount + 7, 1] = (dataGridView4[0, i].Value).ToString();
                ExcelApp.Cells[i + dataGridView3.RowCount + 7, 2] = (dataGridView4[1, i].Value).ToString();
            }
            ExcelApp.Visible = true;*/

            //ДАТА

            /*string s = DateTime.Now.ToString("dd.MM.yyyy");
            MessageBox.Show(s);*/

            //ПРОВЕРКА НА РЕГИСТР

            /*for (int i = 0; i < textBox2.TextLength; i++)
            {
                if (char.IsUpper(for_pass[i]))
                {
                    await command.ExecuteNonQueryAsync();

                    MessageBox.Show("Регистрация успешно завершена");

                    Form1 open_form = new Form1();
                    open_form.Show();
                    this.Hide();

                    break;
                }
            }*/
        }
    }
}
