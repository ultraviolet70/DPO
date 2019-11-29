using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DPO
{
    public partial class Form1 : Form
    {
        Excel.Application winex = new Excel.Application();

        SqlConnection sql = new SqlConnection("Data Source = DESKTOP-N7ITL14\\KATE;" +
                    "Initial Catalog = DPO3;" +
                    "Integrated Security = True;");

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                sql.Open();
                SqlCommand Add = new SqlCommand("Specialty_Insert", sql);
                Add.CommandType = CommandType.StoredProcedure;
                Add.Parameters.AddWithValue("@Name_Specialty", cueTextBox1.Text);
                Add.Parameters.AddWithValue("@ID_Specialty", cueTextBox4.Text);
                Add.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sql.Close();
            }
            cueTextBox1.Clear();
            cueTextBox4.Clear();
        }

        public SqlCommand command = new SqlCommand("Select [ID_Specialty], [Name_Specialty] as 'Название специальности', [ID_Specialty] as 'Код специальности' from [dbo].[Specialty]");
        public void GetData()
        {
            Action act = () =>
            {
                command.Connection = sql;
                command.Notification = null;
                SqlDependency dependency = new SqlDependency(command);
                SqlDependency.Start(sql.ConnectionString);
                dependency.OnChange += new OnChangeEventHandler(OnDataChanget);
                sql.Open();
                DataTable data = new DataTable();
                data.Load(command.ExecuteReader());
                dataGridView1.DataSource = data;
                sql.Close();
            };
            Invoke(act);
        }

        public void OnDataChanget(object sender, SqlNotificationEventArgs e)
        {
            if (e.Info != SqlNotificationInfo.Invalid)
                GetData();
        }

        public SqlCommand command2 = new SqlCommand("Select [ID_Group], [Name_group] as 'Название группы', [Specialty_ID] as 'Код специальности' from [dbo].[Groups]");
        public void GetData2()
        {
            Action act = () =>
            {
                command2.Connection = sql;
                command2.Notification = null;
                SqlDependency dependency = new SqlDependency(command2);
                SqlDependency.Start(sql.ConnectionString);
                dependency.OnChange += new OnChangeEventHandler(OnDataChanget2);
                sql.Open();
                DataTable data = new DataTable();
                data.Load(command2.ExecuteReader());
                dataGridView2.DataSource = data;
                sql.Close();
            };
            Invoke(act);
        }

        public void OnDataChanget2(object sender, SqlNotificationEventArgs e)
        {
            if (e.Info != SqlNotificationInfo.Invalid)
                GetData2();
        }

        public SqlCommand command5 = new SqlCommand("Select [ID_English_levels], [Name_English_levels] as 'Название уровня' from [dbo].[English_levels]");
        public void GetData5()
        {
            Action act = () =>
            {
                command5.Connection = sql;
                command5.Notification = null;
                SqlDependency dependency = new SqlDependency(command5);
                SqlDependency.Start(sql.ConnectionString);
                dependency.OnChange += new OnChangeEventHandler(OnDataChanget5);
                sql.Open();
                DataTable data = new DataTable();
                data.Load(command5.ExecuteReader());
                dataGridView6.DataSource = data;
                sql.Close();
            };
            Invoke(act);
        }

        public void OnDataChanget5(object sender, SqlNotificationEventArgs e)
        {
            if (e.Info != SqlNotificationInfo.Invalid)
                GetData5();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            GetData();
            GetData2();
            GetData3();
            GetData4();
            GetData5();
            dataGridView1.Columns["ID_Specialty"].Visible = false;
            dataGridView2.Columns["ID_Group"].Visible = false;
            dataGridView6.Columns["ID_English_levels"].Visible = false;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                sql.Open();
                SqlCommand Add = new SqlCommand("Group_Insert", sql);
                Add.CommandType = CommandType.StoredProcedure;
                Add.Parameters.AddWithValue("@Name_group", cueTextBox2.Text);
                Add.Parameters.AddWithValue("@Specialty_ID", cueTextBox3.Text);
                Add.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sql.Close();
            }
            cueTextBox2.Clear();
            cueTextBox3.Clear();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.Show();
        }

        public SqlCommand command3 = new SqlCommand("Select [Номер отделения], [Номер группы], [Номер договора], [Дата оплаты], [Стоимость], [Фамилия студента], [Имя студента], [Отчество студента], [Номер телефона студента], [Почта студента], [Дата рождения студента],  [Наличие паспорта студента], [Наличие фотографии студента], [Фамилия родителя], [Имя родителя], [Отчество родителя], [Телефон родителя], [Почта родителя], [Наличие паспорта родителя] from [dbo].[MetodickiView]");
        public void GetData3()
        {
            Action act = () =>
            {
                command3.Connection = sql;
                command3.Notification = null;
                SqlDependency dependency = new SqlDependency(command3);
                SqlDependency.Start(sql.ConnectionString);
                dependency.OnChange += new OnChangeEventHandler(OnDataChanget3);
                sql.Open();
                DataTable data = new DataTable();
                data.Load(command3.ExecuteReader());
                dataGridView3.DataSource = data;
                sql.Close();
            };
            Invoke(act);
        }

        public void OnDataChanget3(object sender, SqlNotificationEventArgs e)
        {
            if (e.Info != SqlNotificationInfo.Invalid)
                GetData3();
        }

        public SqlCommand command4 = new SqlCommand("Select [Номер отделения], [Номер группы], [Номер договора], [Дата оплаты], [Стоимость], [Со школой], [Фамилия студента], [Имя студента], [Отчество студента], [Номер телефона студента], [Почта студента], [Дата рождения студента],  [Наличие паспорта студента], [Наличие фотографии студента], [Фамилия родителя], [Имя родителя], [Отчество родителя], [Телефон родителя], [Почта родителя], [Наличие паспорта родителя] from [dbo].[DOVYZView]");
        public void GetData4()
        {
            Action act = () =>
            {
                command4.Connection = sql;
                command4.Notification = null;
                SqlDependency dependency = new SqlDependency(command4);
                SqlDependency.Start(sql.ConnectionString);
                dependency.OnChange += new OnChangeEventHandler(OnDataChanget4);
                sql.Open();
                DataTable data = new DataTable();
                data.Load(command4.ExecuteReader());
                dataGridView4.DataSource = data;
                sql.Close();
            };
            Invoke(act);
        }

        public void OnDataChanget4(object sender, SqlNotificationEventArgs e)
        {
            if (e.Info != SqlNotificationInfo.Invalid)
                GetData4();
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            DataRowView id = (DataRowView)dataGridView1.CurrentRow.DataBoundItem;
            try
            {
                sql.Open();
                SqlCommand del = new SqlCommand("Specialty_Delete", sql);
                if (MessageBox.Show("Вы действительно хотите удалить данную специальность?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    del.CommandType = CommandType.StoredProcedure;
                    del.Parameters.AddWithValue("@ID_Specialty", (string)id["ID_Specialty"]);
                    del.ExecuteNonQuery();
                }
                else
                {

                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sql.Close();
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DataRowView id1 = (DataRowView)dataGridView2.CurrentRow.DataBoundItem;
            try
            {
                sql.Open();
                SqlCommand del = new SqlCommand("Group_Delete", sql);
                if (MessageBox.Show("Вы действительно хотите удалить данную группу?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    del.CommandType = CommandType.StoredProcedure;
                    del.Parameters.AddWithValue("@ID_Group", (int)id1["ID_Group"]);
                    del.ExecuteNonQuery();
                }
                else
                {

                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sql.Close();
            }
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            Form4 form4 = new Form4();
            form4.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application ex = new Excel.Application();//Объявляем приложение
                ex.Visible = false;
                ex.SheetsInNewWorkbook = 1;//Количество листов в рабочей книге
                Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);//Добавить рабочую книгу
                ex.DisplayAlerts = false;//Отключить отображение окон с сообщениями
                Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);//Получаем первый лист документа
                sheet.Name = "Рабочие тетради";//Название листа
                sheet.Tab.Color = Color.LightBlue;

                ex.Cells[1, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; //Устанавливаем выравнивание ячеек 
                ex.Cells[1, 1].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter; //Устанавливаем выравнивание ячеек 

                Excel.Range range2 = sheet.get_Range("A1").Cells;
                range2.Merge(Type.Missing);
                range2.Cells[1, 1] = "Номер отделения";
                range2.ColumnWidth = 18.18;
                range2.Cells.Font.Name = "Calibri";
                range2.Cells.Font.Size = 10;
                range2.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range2.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range2.Interior.Color = Color.LightYellow;
                range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range3 = sheet.get_Range("A2").Cells;
                range3.Merge(Type.Missing);
                range3.ColumnWidth = 18.18;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range3.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[0].Value;
                }

                Excel.Range range4 = sheet.get_Range("B1").Cells;
                range4.Merge(Type.Missing);
                range4.Cells[1, 1] = "Номер группы";
                range4.ColumnWidth = 19.73;
                range4.Cells.Font.Name = "Calibri";
                range4.Cells.Font.Size = 10;
                range4.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range4.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range4.Interior.Color = Color.LightYellow;
                range4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range5 = sheet.get_Range("B2").Cells;
                range5.Merge(Type.Missing);
                range5.ColumnWidth = 19.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range5.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[1].Value;
                }

                Excel.Range range6 = sheet.get_Range("C1").Cells;
                range6.Merge(Type.Missing);
                range6.Cells[1, 1] = "Номер договора";
                range6.ColumnWidth = 20;
                range6.Cells.Font.Name = "Calibri";
                range6.Cells.Font.Size = 10;
                range6.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range6.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range6.Interior.Color = Color.LightYellow;
                range6.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range6.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range6.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range7 = sheet.get_Range("C2").Cells;
                range7.Merge(Type.Missing);
                range7.ColumnWidth = 20;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range7.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[2].Value;
                }

                Excel.Range range8 = sheet.get_Range("D1").Cells;
                range8.Merge(Type.Missing);
                range8.Cells[1, 1] = "Дата оплаты";
                range8.ColumnWidth = 18.91;
                range8.Cells.Font.Name = "Calibri";
                range8.Cells.Font.Size = 10;
                range8.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range8.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range8.Interior.Color = Color.LightYellow;
                range8.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range8.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range8.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range9 = sheet.get_Range("D2").Cells;
                range9.Merge(Type.Missing);
                range9.ColumnWidth = 18.91;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range9.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[3].Value;
                }

                Excel.Range range10 = sheet.get_Range("E1").Cells;
                range10.Merge(Type.Missing);
                range10.Cells[1, 1] = "Стоимость";
                range10.ColumnWidth = 15.64;
                range10.Cells.Font.Name = "Calibri";
                range10.Cells.Font.Size = 10;
                range10.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range10.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range10.Interior.Color = Color.LightYellow; range10.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range10.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range10.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range11 = sheet.get_Range("E2").Cells;
                range11.Merge(Type.Missing);
                range11.ColumnWidth = 15.64;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range11.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[4].Value;
                }

                Excel.Range range12 = sheet.get_Range("F1").Cells;
                range12.Merge(Type.Missing);
                range12.Cells[1, 1] = "Фамилия студента";
                range12.ColumnWidth = 14.91;
                range12.Cells.Font.Name = "Calibri";
                range12.Cells.Font.Size = 10;
                range12.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range12.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range12.Interior.Color = Color.LightYellow;
                range12.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range12.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range12.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range13 = sheet.get_Range("F2").Cells;
                range13.Merge(Type.Missing);
                range13.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range13.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[5].Value;
                }

                Excel.Range range14 = sheet.get_Range("G1").Cells;
                range14.Merge(Type.Missing);
                range14.Cells[1, 1] = "Имя студента";
                range14.ColumnWidth = 24.73;
                range14.Cells.Font.Name = "Calibri";
                range14.Cells.Font.Size = 10;
                range14.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range14.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range14.Interior.Color = Color.LightYellow;
                range14.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range14.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range14.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range15 = sheet.get_Range("G2").Cells;
                range15.Merge(Type.Missing);
                range15.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range15.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[6].Value;
                }

                Excel.Range range16 = sheet.get_Range("H1").Cells;
                range16.Merge(Type.Missing);
                range16.Cells[1, 1] = "Отчество студента";
                range16.ColumnWidth = 24.73;
                range16.Cells.Font.Name = "Calibri";
                range16.Cells.Font.Size = 10;
                range16.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range16.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range16.Interior.Color = Color.LightYellow;
                range16.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range16.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range16.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range17 = sheet.get_Range("H2").Cells;
                range17.Merge(Type.Missing);
                range17.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range17.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[7].Value;
                }

                Excel.Range range18 = sheet.get_Range("I1").Cells;
                range18.Merge(Type.Missing);
                range18.Cells[1, 1] = "Номер телефона студента";
                range18.ColumnWidth = 24.73;
                range18.Cells.Font.Name = "Calibri";
                range18.Cells.Font.Size = 10;
                range18.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range18.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range18.Interior.Color = Color.LightYellow;
                range18.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range18.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range18.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range19 = sheet.get_Range("I2").Cells;
                range19.Merge(Type.Missing);
                range19.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range19.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[8].Value;
                }

                Excel.Range range20 = sheet.get_Range("J1").Cells;
                range20.Merge(Type.Missing);
                range20.Cells[1, 1] = "Почта студента";
                range20.ColumnWidth = 24.73;
                range20.Cells.Font.Name = "Calibri";
                range20.Cells.Font.Size = 10;
                range20.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range20.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range20.Interior.Color = Color.LightYellow;
                range20.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range20.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range20.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range21 = sheet.get_Range("J2").Cells;
                range21.Merge(Type.Missing);
                range21.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range21.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[9].Value;
                }

                Excel.Range range22 = sheet.get_Range("K1").Cells;
                range22.Merge(Type.Missing);
                range22.Cells[1, 1] = "Дата рождения студента";
                range22.ColumnWidth = 24.73;
                range22.Cells.Font.Name = "Calibri";
                range22.Cells.Font.Size = 10;
                range22.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range22.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range22.Interior.Color = Color.LightYellow;
                range22.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range22.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range22.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range23 = sheet.get_Range("K2").Cells;
                range23.Merge(Type.Missing);
                range23.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range23.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[10].Value;
                }

                Excel.Range range24 = sheet.get_Range("L1").Cells;
                range24.Merge(Type.Missing);
                range24.Cells[1, 1] = "Наличие паспорта студента";
                range24.ColumnWidth = 24.73;
                range24.Cells.Font.Name = "Calibri";
                range24.Cells.Font.Size = 10;
                range24.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range24.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range24.Interior.Color = Color.LightYellow;
                range24.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range24.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range24.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range25 = sheet.get_Range("L2").Cells;
                range25.Merge(Type.Missing);
                range25.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range25.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[11].Value;
                }

                Excel.Range range26 = sheet.get_Range("M1").Cells;
                range26.Merge(Type.Missing);
                range26.Cells[1, 1] = "Наличие фото студента";
                range26.ColumnWidth = 24.73;
                range26.Cells.Font.Name = "Calibri";
                range26.Cells.Font.Size = 10;
                range26.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range26.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range26.Interior.Color = Color.LightYellow;
                range26.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range26.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range26.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range27 = sheet.get_Range("M2").Cells;
                range27.Merge(Type.Missing);
                range27.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range27.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[12].Value;
                }

                Excel.Range range28 = sheet.get_Range("N1").Cells;
                range28.Merge(Type.Missing);
                range28.Cells[1, 1] = "Фамилия родителя";
                range28.ColumnWidth = 24.73;
                range28.Cells.Font.Name = "Calibri";
                range28.Cells.Font.Size = 10;
                range28.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range28.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range28.Interior.Color = Color.LightYellow;
                range28.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range28.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range28.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range29 = sheet.get_Range("N2").Cells;
                range29.Merge(Type.Missing);
                range29.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range29.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[13].Value;
                }

                Excel.Range range30 = sheet.get_Range("O1").Cells;
                range30.Merge(Type.Missing);
                range30.Cells[1, 1] = "Имя родителя";
                range30.ColumnWidth = 24.73;
                range30.Cells.Font.Name = "Calibri";
                range30.Cells.Font.Size = 10;
                range30.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range30.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range30.Interior.Color = Color.LightYellow;
                range30.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range30.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range30.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range31 = sheet.get_Range("O2").Cells;
                range31.Merge(Type.Missing);
                range31.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range31.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[14].Value;
                }

                Excel.Range range32 = sheet.get_Range("P1").Cells;
                range32.Merge(Type.Missing);
                range32.Cells[1, 1] = "Отчество родителя";
                range32.ColumnWidth = 24.73;
                range32.Cells.Font.Name = "Calibri";
                range32.Cells.Font.Size = 10;
                range32.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range32.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range32.Interior.Color = Color.LightYellow;
                range32.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range32.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range32.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range33 = sheet.get_Range("P2").Cells;
                range33.Merge(Type.Missing);
                range33.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range33.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[15].Value;
                }

                Excel.Range range34 = sheet.get_Range("Q1").Cells;
                range34.Merge(Type.Missing);
                range34.Cells[1, 1] = "Номер телефона родителя";
                range34.ColumnWidth = 24.73;
                range34.Cells.Font.Name = "Calibri";
                range34.Cells.Font.Size = 10;
                range34.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range34.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range34.Interior.Color = Color.LightYellow;
                range34.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range34.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range34.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range35 = sheet.get_Range("Q2").Cells;
                range35.Merge(Type.Missing);
                range35.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range35.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[16].Value;
                }

                Excel.Range range36 = sheet.get_Range("R1").Cells;
                range36.Merge(Type.Missing);
                range36.Cells[1, 1] = "Почта родителя";
                range36.ColumnWidth = 24.73;
                range36.Cells.Font.Name = "Calibri";
                range36.Cells.Font.Size = 10;
                range36.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range36.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range36.Interior.Color = Color.LightYellow;
                range36.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range36.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range36.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range37 = sheet.get_Range("R2").Cells;
                range37.Merge(Type.Missing);
                range37.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range37.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[17].Value;
                }

                Excel.Range range38 = sheet.get_Range("S1").Cells;
                range38.Merge(Type.Missing);
                range38.Cells[1, 1] = "Наличие паспорта родителя";
                range38.ColumnWidth = 24.73;
                range38.Cells.Font.Name = "Calibri";
                range38.Cells.Font.Size = 10;
                range38.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range38.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range38.Interior.Color = Color.LightYellow;
                range38.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;//Границы
                range38.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThick;
                range38.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThick;

                Excel.Range range39 = sheet.get_Range("S2").Cells;
                range39.Merge(Type.Missing);
                range39.ColumnWidth = 24.73;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    range39.Cells[i + 1, 1] = dataGridView3.Rows[i].Cells[18].Value;
                }

                ex.Visible = true;//Отобразить Excel

                ex.Application.ActiveWorkbook.SaveAs("Рабочие тетради.xlsx", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Form5 form5 = new Form5();
            form5.Show();
        }
    }
}
