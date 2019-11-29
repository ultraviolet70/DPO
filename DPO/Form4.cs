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
    public partial class Form4 : Form
    {

        SqlConnection sql = new SqlConnection("Data Source = DESKTOP-N7ITL14\\KATE;" +
                            "Initial Catalog = DPO3;" +
                            "Integrated Security = True;");

        string ch4ch, ch5ch, ch6ch, ch7ch;

        private void checkBox28_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox28.Checked)
            {
                ch4ch = "Да";
            }
            else
            {
                ch4ch = "Нет";
            }
        }

        private void checkBox29_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox29.Checked)
            {
                ch5ch = "Да";
            }
            else
            {
                ch5ch = "Нет";
            }
        }

        private void checkBox35_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox35.Checked)
            {
                ch6ch = "Да";
            }
            else
            {
                ch6ch = "Нет";
            }
        }

        private void checkBox41_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox41.Checked)
            {
                ch7ch = "Да";
            }
            else
            {
                ch7ch = "Нет";
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                sql.Open();
                SqlCommand Add = new SqlCommand("DOVYZ_Insert", sql);
                Add.CommandType = CommandType.StoredProcedure;
                Add.Parameters.AddWithValue("@Cost1", int.Parse(cueTextBox36.Text));
                Add.Parameters.AddWithValue("@Date_of_payment1", maskedTextBox27.Text);
                Add.Parameters.AddWithValue("@Сontract_number", cueTextBox40.Text);
                Add.Parameters.AddWithValue("@School_dovyz", ch7ch);

                Add.Parameters.AddWithValue("@Surname_student_dovyz", cueTextBox20.Text);
                Add.Parameters.AddWithValue("@Name_student_dovyz", cueTextBox21.Text);
                Add.Parameters.AddWithValue("@Middlename_student_dovyz", cueTextBox22.Text);
                Add.Parameters.AddWithValue("@Phone_student_dovyz", maskedTextBox24.Text);
                Add.Parameters.AddWithValue("@Email_student_dovyz", cueTextBox25.Text);
                Add.Parameters.AddWithValue("@Date_of_birth_student_dovyz", maskedTextBox23.Text);
                Add.Parameters.AddWithValue("@Pass_student_dovyz", ch4ch);
                Add.Parameters.AddWithValue("@Photo_student_dovyz", ch5ch);

                Add.Parameters.AddWithValue("@Surname_Parents_dovyz", cueTextBox30.Text);
                Add.Parameters.AddWithValue("@Name_Parents_dovyz", cueTextBox31.Text);
                Add.Parameters.AddWithValue("@Middlename_Parents_dovyz", cueTextBox32.Text);
                Add.Parameters.AddWithValue("@Phone_Parents_dovyz", maskedTextBox33.Text);
                Add.Parameters.AddWithValue("@Email_Parents_dovyz", cueTextBox34.Text);
                Add.Parameters.AddWithValue("@Pass_Parents_dovyz", ch6ch);
                Add.Parameters.AddWithValue("@Group_ID", comboBox27.SelectedValue.ToString());
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

        private void comboBox26_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dAdapter = new SqlDataAdapter("SELECT ID_Group, Name_group FROM Groups WHERE Specialty_ID='" + comboBox26.SelectedItem.ToString() + "'", sql);
            DataTable dt = new DataTable();
            dAdapter.Fill(dt);
            comboBox27.ValueMember = "ID_Group";
            comboBox27.DisplayMember = "Name_group";
            comboBox27.DataSource = dt;
            sql.Close();

            if (comboBox26.Text != "")
            {
                comboBox27.Enabled = true;
            }
        }

        public Form4()
        {
            InitializeComponent();

            ch4ch = "Нет";
            ch5ch = "Нет";
            ch6ch = "Нет";
            ch7ch = "Нет";

            sql.Open();
            SqlCommand command5 = new SqlCommand("SELECT ID_Specialty FROM Specialty", sql);
            SqlDataReader read5 = command5.ExecuteReader();
            while (read5.Read())
            {
                comboBox26.Items.Add(read5.GetValue(0).ToString());
            }
            read5.Close();
        }
    }
}
