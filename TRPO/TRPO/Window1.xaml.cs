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
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        public Window1()
        {
            InitializeComponent();
        }
        private void closer_MouseUp(object sender, MouseButtonEventArgs e)
        {
            this.Close();
        }

        private void remover_MouseUp(object sender, MouseButtonEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
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
            System.Windows.Application.Current.Shutdown();
            this.Close();
        }

        private void remover_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {

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

        private void Button1_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                bool f =true;
                if (TB1.Text != "" && TB2.Password != "")
                {
                    dt_user = Select("SELECT id_Медсестры,Медсестра.Пароль, Медсестра.Логин From Медсестра");
                    for (int i = 0; i < dt_user.Rows.Count; i++)
                    {
                        if (dt_user.Rows[i][1].ToString() == TB2.Password)
                        {
                            if (dt_user.Rows[i][2].ToString() == TB1.Text)
                            {
                                f = false;
                                int id = (int)dt_user.Rows[i][0];
                                MainWindow mainWindow = new MainWindow(id);
                                Hide();
                                mainWindow.ShowDialog();
                                ShowDialog();
                            }
                            
                        }
                    }
                    
                }
                if (f)
                {
                    MessageBox.Show("Проверьте правильность введенных данных");
                }
            }
            catch
            {
                MessageBox.Show("Недопустимый символ");
            }
        }

        private void TextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Window2 w2 = new Window2();

            Hide();
            w2.ShowDialog();
            ShowDialog();
        }
    }
}
