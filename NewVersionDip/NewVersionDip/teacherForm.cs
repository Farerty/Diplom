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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace NewVersionDip
{
    public partial class teacherForm : Form
    {
        private Point MouseHook;
        private int idTeacher = 10;
        private int idCafedra = 10;
        public static SqlConnection connection = new SqlConnection(Properties.Settings.Default.connString);
        public int idOfteacher = 0;
        public int idOfCafedra = 0;
        int[] indexSave = new int[2];
        public teacherForm()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
        }

        public Central Form1
        {
            get => default;
            set
            {
            }
        }

        private void teacherBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.teacherBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.kursRabDataSet);

        }
        
        private void teacherForm_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursRabDataSet.Cafedra". При необходимости она может быть перемещена или удалена.
            this.cafedraTableAdapter.Fill(this.kursRabDataSet.Cafedra);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kursRabDataSet.Teacher". При необходимости она может быть перемещена или удалена.
            this.teacherTableAdapter.Fill(this.kursRabDataSet.Teacher);

        }

        private void ChangeButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(Properties.Settings.Default.connString))
                {

                    string sql = @"SELECT * FROM Cafedra WHERE Caf_ID = " + caf_IDTextBox1.Text + "AND name = " + "'" + nameTextBox.Text + "'";
                    SqlDataAdapter sqlQuery = new SqlDataAdapter(sql, connection);
                    System.Data.DataTable dtblDob = new System.Data.DataTable();
                    sqlQuery.Fill(dtblDob);

                    if (!(dtblDob.Rows.Count == 1))
                    {
                        conn.Open();
                        using (SqlCommand cmd = new SqlCommand("UPDATE Cafedra SET name=@name " + "WHERE Caf_Id = " + Convert.ToInt32(caf_IDTextBox1.Text), conn))
                        {
                            //cmd.Parameters.AddWithValue("@Caf_ID", caf_IDTextBox1.Text);Caf_ID=@Caf_ID, 

                            cmd.Parameters.AddWithValue("@name", nameTextBox.Text);
                            //add whatever parameters you required to update here
                            int rows = cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("UPDATE Teacher SET  Caf_ID=@Caf_ID, T_Family=@T_Family, T_FirstName = @T_FirstName" +
                        ", T_SecondNamme = @T_SecondName, dateOfBirth = @dateOfBirth, ScienceStepen = @ScienceStepen, ScienceZvanie = @ScienceZvanie, PassportDate = @PassportDate" +
                        ", SNILS = @SNILS, toun = @toun, strit = @strit, house = @house, flow = @flow, ind = @ind, telephone = @telephone, diplomDate = @diplomDate, placeWorking = @placeWorking, Dolgnost = @Dolgnost" + " WHERE T_Id=" + Convert.ToInt32(t_IDTextBox.Text), conn))
                    {
                        //cmd.Parameters.AddWithValue("@T_ID", t_IDTextBox.Text);T_ID=@T_ID,
                        cmd.Parameters.AddWithValue("@Caf_ID", caf_IDTextBox.Text);
                        cmd.Parameters.AddWithValue("@T_Family", t_FamilyTextBox.Text);
                        cmd.Parameters.AddWithValue("@T_FirstName", t_FirstNameTextBox.Text);
                        cmd.Parameters.AddWithValue("@T_SecondName", t_SecondNammeTextBox.Text);
                        cmd.Parameters.AddWithValue("@dateOfBirth", dateOfBirthDateTimePicker.Value);
                        cmd.Parameters.AddWithValue("@ScienceStepen", scienceStepenTextBox.Text);
                        cmd.Parameters.AddWithValue("@ScienceZvanie", scienceZvanieTextBox.Text);
                        cmd.Parameters.AddWithValue("@PassportDate", passportDateTextBox.Text);
                        cmd.Parameters.AddWithValue("@SNILS", sNILSTextBox.Text);
                        cmd.Parameters.AddWithValue("@toun", tounTextBox.Text);
                        cmd.Parameters.AddWithValue("@strit", stritTextBox.Text);
                        cmd.Parameters.AddWithValue("@house", houseTextBox.Text);
                        cmd.Parameters.AddWithValue("@flow", flowTextBox.Text);
                        cmd.Parameters.AddWithValue("@ind", indTextBox.Text);
                        cmd.Parameters.AddWithValue("@telephone", telephoneTextBox.Text);
                        cmd.Parameters.AddWithValue("@diplomDate", diplomDateDateTimePicker.Value);
                        cmd.Parameters.AddWithValue("@placeWorking", placeWorkingTextBox.Text);
                        cmd.Parameters.AddWithValue("@Dolgnost", dolgnostTextBox.Text);

                        //add whatever parameters you required to update here
                        int rows = cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }
                refresh();
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
            finally { connection.Close(); }
            
        }

        private void DeleteBytton_Click(object sender, EventArgs e)
        {
            try
            {
                string sql = "";
                idOfteacher = Convert.ToInt32(t_IDTextBox.Text);
                sql = @"DELETE FROM Teacher WHERE T_ID = " + "@ID";
                connection.Open();
                using (SqlCommand sqlCommand = new SqlCommand(sql, connection))
                {
                    sqlCommand.Parameters.Add(new SqlParameter("@ID", SqlDbType.Int));
                    sqlCommand.Parameters["@ID"].Value = Convert.ToInt32(idOfteacher);

                    sqlCommand.ExecuteNonQuery();
                }
                connection.Close();

                refresh();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            
                
            
        }
        private void teacherPlus_Click(object sender, EventArgs e)
        {
            try
            {
                if (idTeacher < 0)
                {
                    idTeacher = 0;
                }
                idTeacher += 10;
                string sql = "SELECT * FROM Teacher WHERE T_ID > " + (idTeacher - 10) + "and Caf_ID " + " < " + idTeacher;
                using (SqlCommand sqlCommand = new SqlCommand(sql, connection))
                {
                    connection.Open();
                    using (SqlDataReader dataReader = sqlCommand.ExecuteReader())
                    {
                        DataTable dataTable = new DataTable();
                        dataTable.Load(dataReader);
                        teacherDataGridView.DataSource = dataTable;
                        dataReader.Close();
                    }
                    connection.Close();
                }
                teacherDataGridView.FirstDisplayedScrollingRowIndex = teacherDataGridView.Rows[teacherDataGridView.Rows.Count - 1].Index;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            

        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (idCafedra < 0)
                {
                    idCafedra = 0;
                }
                idCafedra += 10;
                string sql = @"SELECT * FROM Cafedra WHERE Caf_ID > " + (idCafedra - 10) + "and Caf_ID " + " < " + idCafedra;
                using (SqlCommand sqlCommand = new SqlCommand(sql, connection))
                {
                    connection.Open();
                    using (SqlDataReader dataReader = sqlCommand.ExecuteReader())
                    {
                        DataTable dataTable = new DataTable();
                        dataTable.Load(dataReader);
                        cafedraDataGridView.DataSource = dataTable;
                        dataReader.Close();
                    }
                    connection.Close();
                }
                cafedraDataGridView.FirstDisplayedScrollingRowIndex = cafedraDataGridView.Rows[cafedraDataGridView.Rows.Count - 1].Index;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            
        }

        private void CafedraMinus_Click(object sender, EventArgs e)
        {
            try
            {
                idCafedra -= 10;
                if (idCafedra < 0)
                {
                    idCafedra = 0;
                }

                string sql = @"SELECT * FROM Cafedra WHERE Caf_ID <" + idCafedra;
                using (SqlCommand sqlCommand = new SqlCommand(sql, connection))
                {
                    connection.Open();
                    using (SqlDataReader dataReader = sqlCommand.ExecuteReader())
                    {
                        DataTable dataTable = new DataTable();
                        dataTable.Load(dataReader);
                        cafedraDataGridView.DataSource = dataTable;
                        dataReader.Close();
                    }
                    connection.Close();
                }
                cafedraDataGridView.FirstDisplayedScrollingRowIndex = cafedraDataGridView.Rows[cafedraDataGridView.Rows.Count - 1].Index;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
           
        }

        private void teacherMinus_Click(object sender, EventArgs e)
        {
            try
            {
                idTeacher -= 10;
                if (idTeacher < 0)
                {
                    idTeacher = 0;
                }

                string sql = "SELECT * FROM Teacher WHERE T_ID > " + (idTeacher - 10) + "and T_ID " + " < " + idTeacher;
                using (SqlCommand sqlCommand = new SqlCommand(sql, connection))
                {
                    connection.Open();
                    using (SqlDataReader dataReader = sqlCommand.ExecuteReader())
                    {
                        DataTable dataTable = new DataTable();
                        dataTable.Load(dataReader);
                        teacherDataGridView.DataSource = dataTable;
                        dataReader.Close();
                    }
                    connection.Close();
                }
                teacherDataGridView.FirstDisplayedScrollingRowIndex = teacherDataGridView.Rows[teacherDataGridView.Rows.Count - 1].Index;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            
        }
        public void refresh()
        {
            try
            {

            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            string sql = "SELECT * FROM Teacher";
            using (SqlCommand sqlCommand = new SqlCommand(sql, connection))
            {
                connection.Open();
                using (SqlDataReader dataReader = sqlCommand.ExecuteReader())
                {
                    DataTable dataTable = new DataTable();
                    dataTable.Load(dataReader);
                    teacherDataGridView.DataSource = dataTable;
                    dataReader.Close();
                }
                connection.Close();
            }
            teacherDataGridView.FirstDisplayedScrollingRowIndex = teacherDataGridView.Rows[teacherDataGridView.Rows.Count - 1].Index;
            string sql1 = "SELECT * FROM Cafedra";
            using (SqlCommand sqlCommand = new SqlCommand(sql1, connection))
            {
                connection.Open();
                using (SqlDataReader dataReader = sqlCommand.ExecuteReader())
                {
                    DataTable dataTable = new DataTable();
                    dataTable.Load(dataReader);
                    cafedraDataGridView.DataSource = dataTable;
                    dataReader.Close();
                }
                connection.Close();
            }
            cafedraDataGridView.FirstDisplayedScrollingRowIndex = cafedraDataGridView.Rows[cafedraDataGridView.Rows.Count - 1].Index;
        }

        private void teacherForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel6_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
        }

        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
        }

        private void teacherDataGridView_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
        }

        private void panel4_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
        }

        private void cafedraDataGridView_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
        }
    }
}
