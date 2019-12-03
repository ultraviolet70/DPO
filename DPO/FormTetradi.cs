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

namespace DPO
{
    public partial class FormTetradi : Form
    {

        SqlConnection sql = new SqlConnection("Data Source = DESKTOP-N7ITL14\\KATE;" +
                        "Initial Catalog = DPO3;" +
                        "Integrated Security = True;");

        string ch1ch, ch2ch, ch3ch;

        public FormTetradi()
        {
            InitializeComponent();

            ch1ch = "Нет";
            ch2ch = "Нет";
            ch3ch = "Нет";

            sql.Open();
            SqlCommand command1 = new SqlCommand("SELECT ID_Specialty FROM Specialty", sql);
            SqlDataReader read1 = command1.ExecuteReader();
            while (read1.Read())
            {
                comboBox1.Items.Add(read1.GetValue(0).ToString());
            }
            read1.Close();
        }

        private void materialFlatButton1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dAdapter = new SqlDataAdapter("SELECT ID_Group, Name_group FROM Groups WHERE Specialty_ID='" + comboBox1.SelectedItem.ToString() + "'", sql);
            DataTable dt = new DataTable();
            dAdapter.Fill(dt);
            comboBox2.ValueMember = "ID_Group";
            comboBox2.DisplayMember = "Name_group";
            comboBox2.DataSource = dt;
            sql.Close();

            if (comboBox1.Text != "")
            {
                comboBox2.Enabled = true;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                ch1ch = "Да";
            }
            else
            {
                ch1ch = "Нет";
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                ch2ch = "Да";
            }
            else
            {
                ch2ch = "Нет";
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                ch3ch = "Да";
            }
            else
            {
                ch3ch = "Нет";
            }
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                sql.Open();
                SqlCommand Add = new SqlCommand("Metodichki_Insert", sql);
                Add.CommandType = CommandType.StoredProcedure;
                Add.Parameters.AddWithValue("@Cost", int.Parse(cueTextBox9.Text));
                Add.Parameters.AddWithValue("@Date_of_payment", maskedTextBox4.Text);
                Add.Parameters.AddWithValue("@Сontract_number", cueTextBox10.Text);

                Add.Parameters.AddWithValue("@Surname_student_Metodichki", cueTextBox1.Text);
                Add.Parameters.AddWithValue("@Name_student_Metodichki", cueTextBox2.Text);
                Add.Parameters.AddWithValue("@Middlename_student_Metodichki", cueTextBox3.Text);
                Add.Parameters.AddWithValue("@Phone_student_Metodichki", maskedTextBox2.Text);
                Add.Parameters.AddWithValue("@Email_student_Metodichki", cueTextBox4.Text);
                Add.Parameters.AddWithValue("@Date_of_birth_student_Metodichki", maskedTextBox4.Text);
                Add.Parameters.AddWithValue("@Pass_student_Metodichki", ch1ch);
                Add.Parameters.AddWithValue("@Photo_student_Metodichki", ch2ch);

                Add.Parameters.AddWithValue("@Surname_Parents_Metodichki", cueTextBox8.Text);
                Add.Parameters.AddWithValue("@Name_Parents_Metodichki", cueTextBox7.Text);
                Add.Parameters.AddWithValue("@Middlename_Parents_Metodichki", cueTextBox6.Text);
                Add.Parameters.AddWithValue("@Phone_Parents_Metodichki", maskedTextBox3.Text);
                Add.Parameters.AddWithValue("@Email_Parents_Metodichki", cueTextBox5.Text);
                Add.Parameters.AddWithValue("@Pass_Parents_Metodichki", ch3ch);
                Add.Parameters.AddWithValue("@Group_ID", comboBox2.SelectedValue.ToString());
                Add.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sql.Close();
                this.Close();
            }
        }
    }
}
