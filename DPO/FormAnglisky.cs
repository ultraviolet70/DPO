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
    public partial class FormAnglisky : Form
    {

        SqlConnection sql = new SqlConnection("Data Source = DESKTOP-N7ITL14\\KATE;" +
                            "Initial Catalog = DPO3;" +
                            "Integrated Security = True;");

        string ch1ch, ch2ch, ch3ch;

        public FormAnglisky()
        {
            InitializeComponent();

            ch1ch = "Нет";
            ch2ch = "Нет";
            ch3ch = "Нет";

            sql.Open();
            SqlCommand command1 = new SqlCommand("SELECT Name_English_levels FROM English_levels", sql);
            SqlDataReader read1 = command1.ExecuteReader();
            while (read1.Read())
            {
                comboBox52.Items.Add(read1.GetValue(0).ToString());
            }
            read1.Close();

            SqlCommand command2 = new SqlCommand("SELECT English_levels_time FROM English_levels_time", sql);
            SqlDataReader read2 = command2.ExecuteReader();
            while (read2.Read())
            {
                comboBox53.Items.Add(read2.GetValue(0).ToString());
            }
            read2.Close();

            SqlCommand command3 = new SqlCommand("SELECT ID_Specialty FROM Specialty", sql);
            SqlDataReader read3 = command3.ExecuteReader();
            while (read3.Read())
            {
                comboBox66.Items.Add(read3.GetValue(0).ToString());
            }
            read3.Close();
        }

        private void checkBox68_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox68.Checked)
            {
                ch1ch = "Да";
            }
            else
            {
                ch1ch = "Нет";
            }
        }

        private void checkBox69_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox69.Checked)
            {
                ch2ch = "Да";
            }
            else
            {
                ch2ch = "Нет";
            }
        }

        private void checkBox75_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox75.Checked)
            {
                ch3ch = "Да";
            }
            else
            {
                ch3ch = "Нет";
            }
        }

        private void comboBox66_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter dAdapter = new SqlDataAdapter("SELECT ID_Group, Name_group FROM Groups WHERE Specialty_ID='" + comboBox66.SelectedItem.ToString() + "'", sql);
            DataTable dt = new DataTable();
            dAdapter.Fill(dt);
            comboBox67.ValueMember = "ID_Group";
            comboBox67.DisplayMember = "Name_group";
            comboBox67.DataSource = dt;
            sql.Close();

            if (comboBox66.Text != "")
            {
                comboBox67.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                sql.Open();
                SqlCommand Add = new SqlCommand("Englishst_Insert", sql);
                Add.CommandType = CommandType.StoredProcedure;
                Add.Parameters.AddWithValue("@Cost2", int.Parse(cueTextBox76.Text));
                Add.Parameters.AddWithValue("@Cost3", int.Parse(cueTextBox77.Text));
                Add.Parameters.AddWithValue("@Cost4", int.Parse(cueTextBox78.Text));
                Add.Parameters.AddWithValue("@Date_of_payment2", maskedTextBox79.Text);
                Add.Parameters.AddWithValue("@Сontract_number", cueTextBox80.Text);

                Add.Parameters.AddWithValue("@Surname_student_englishst", cueTextBox60.Text);
                Add.Parameters.AddWithValue("@Name_student_englishst", cueTextBox61.Text);
                Add.Parameters.AddWithValue("@Middlename_student_englishst", cueTextBox62.Text);
                Add.Parameters.AddWithValue("@Phone_student_englishst", maskedTextBox64.Text);
                Add.Parameters.AddWithValue("@Email_student_englishst", cueTextBox65.Text);
                Add.Parameters.AddWithValue("@Date_of_birth_student_englishst", maskedTextBox63.Text);
                Add.Parameters.AddWithValue("@Pass_student_englishst", ch1ch);
                Add.Parameters.AddWithValue("@Photo_student_englishst", ch2ch);

                Add.Parameters.AddWithValue("@Surname_Parents_englishst", cueTextBox70.Text);
                Add.Parameters.AddWithValue("@Name_Parents_englishst", cueTextBox71.Text);
                Add.Parameters.AddWithValue("@Middlename_Parents_englishst", cueTextBox72.Text);
                Add.Parameters.AddWithValue("@Phone_Parents_englishst", maskedTextBox73.Text);
                Add.Parameters.AddWithValue("@Email_Parents_englishst", cueTextBox74.Text);
                Add.Parameters.AddWithValue("@Pass_Parents_englishst", ch3ch);
                Add.Parameters.AddWithValue("@Group_ID", comboBox67.SelectedValue.ToString());
                Add.Parameters.AddWithValue("@English_levels_ID", comboBox52.SelectedItem.ToString());
                Add.Parameters.AddWithValue("@English_levels_time_ID", comboBox53.SelectedItem.ToString());
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
