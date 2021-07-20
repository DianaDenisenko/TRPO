using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Data.Sql;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Media.Effects;
namespace TRPO
{
    /// <summary>
    /// Логика взаимодействия для Window2.xaml
    /// </summary>
    public partial class Window2 : Window
    {
        public Window2()
        {
            InitializeComponent();
        }

        private void ToolBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        private void closer_MouseDown(object sender, MouseButtonEventArgs e)
        {
        
            this.Close();
        }

        private void remover_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.WindowState = WindowState.Minimized;

        }
        DataTable dt_user;
        public DataTable Select(string selectSQL) // функция подключения к базе данных и обработка запросов
        {

            DataTable dataTable = new DataTable("dataBase");                // создаём таблицу в приложении   // подключаемся к базе данных
            SqlConnection sqlConnection = new SqlConnection("server=DESKTOP-JRUISRI\\SQLEXPRESS;Trusted_Connection=Yes;DataBase=TRPO;");
            sqlConnection.Open();                                           // открываем базу данных
            SqlCommand sqlCommand = sqlConnection.CreateCommand();          // создаём команду

            sqlCommand.CommandText = selectSQL;// присваиваем команде текст
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand); // создаём обработчик
            sqlDataAdapter.Fill(dataTable);                                 // возращаем таблицу с результатом
            sqlConnection.Close();
            return dataTable;

        }
        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                dt_user = Select("SELECT id_Медсестры,Медсестра.Пароль, Медсестра.Логин From Медсестра WHERE Уволен is NULL");
                for (int i = 0; i < dt_user.Rows.Count; i++)
                {
                  
                        if (dt_user.Rows[i][2].ToString() == TB4.Text)
                        {

                        MessageBox.Show("Логин занят");
                        TB4.Text = "";
                        }
                      
                }
                if (TB1.Text != "" && TB2.Text != "" && TB3.Text != "" && TB4.Text != "" && TB5.Text != "")
                {
                    SqlConnection sqlConnection1 = new SqlConnection("server=DESKTOP-JRUISRI\\SQLEXPRESS;Trusted_Connection=Yes;DataBase=TRPO;");

                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "Insert into Медсестра (Имя, Фамилия, Отчество,Логин,Пароль)values(@nm,@fm,@ot,@log,@Par)";
                    cmd.Parameters.AddWithValue("@nm", TB1.Text.Trim());
                    cmd.Parameters.AddWithValue("@fm", TB2.Text.Trim());
                    cmd.Parameters.AddWithValue("@ot", TB3.Text.Trim());
                    cmd.Parameters.AddWithValue("@log", TB4.Text.Trim());
                    cmd.Parameters.AddWithValue("@Par", TB5.Text.Trim());
                cmd.Connection = sqlConnection1;
                sqlConnection1.Open();
                    cmd.ExecuteNonQuery();
                    sqlConnection1.Close();
                    MessageBox.Show("Регистрация прошла успешно");
                    this.Close();
                   
                }
                else MessageBox.Show("Введите данные");
        }
            catch
            {
                MessageBox.Show("Недопустимый символ");
                TB4.Text = "";
                TB5.Text = "";
            }
}

        private void TB1_TextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if ((sender as TextBox).Text.Length == 1)
            {
                (sender as TextBox).Text = (sender as TextBox).Text.ToUpper();
                (sender as TextBox).Select((sender as TextBox).Text.Length, 0);
            }
        }

        private void TB1_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp) || inp == (char)Key.Space)
                e.Handled = true;
            if ((sender as TextBox).Text.Length == 1)
            {
                (sender as TextBox).Text = (sender as TextBox).Text.ToUpper();
                (sender as TextBox).Select((sender as TextBox).Text.Length, 0);
            }
        }
    }
}
