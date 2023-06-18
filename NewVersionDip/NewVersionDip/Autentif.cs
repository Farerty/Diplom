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

namespace NewVersionDip
{
    public partial class Autentif : Form
    {
        public bool flag = false; 
        public static SqlConnection connection = new SqlConnection(Properties.Settings.Default.connString);
        OpenFileDialog openFileDialogRes = new OpenFileDialog();
        private Point MouseHook;
        SaveFileDialog saveFileDialogBack = new SaveFileDialog();
        public static bool SQLStat = true;
        public Autentif()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            MessageBox.Show("Чтобы продолжить работу введите логин и пароль для базы данных!");
        }

        public Central Form1
        {
            get => default;
            set
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            try
            {
                string query = @"DECLARE @cookie varbinary(8000); EXEC sys.sp_setapprole " + "'" + textBox1.Text + "'" + ", " + "'" + textBox2.Text + "'" + " , @fCreateCookie = true, @cookie = @cookie OUTPUT;";

                connection.Open();
                SqlCommand sql1 = new SqlCommand(query, connection);
                flag = true;
                sql1.ExecuteNonQuery();
                this.Hide();
                Central form1 = new Central();
                MessageBox.Show("Успешно!");
                form1.Show();
            }
            catch(Exception ex) { textBox1.Text = ""; textBox2.Text = "";    }
            
            
            //sql1 = new SqlCommand(query, connection);
            //sql1.ExecuteNonQuery();
            //connection.Close();
            
            
        }

        

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
        }
    }
}
