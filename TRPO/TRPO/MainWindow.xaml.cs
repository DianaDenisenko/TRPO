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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.Sql;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Media.Effects;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using Microsoft.Win32;
namespace TRPO
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    /// 
    public partial class MainWindow : Window
    {

        private void ToolBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        private void closer_MouseDown(object sender, MouseButtonEventArgs e)
        {
            DialogResult = false;
        }

        private void remover_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }
        string whatTable = "";
        DataTable dt_user;
        DataTable dt2;
        int nowid;
        string cher = "";
        public MainWindow(int d)
        {
            InitializeComponent();
            Vrach5.CommandBindings.Add(new CommandBinding(ApplicationCommands.Paste, OnPasteCommand));
            Pacient4.CommandBindings.Add(new CommandBinding(ApplicationCommands.Paste, OnPasteCommand));
            TB1.DisplayDateStart = DateTime.Now;
            var now = DateTime.Now;
            nowid = d;
            id3 = nowid;
            TB1.DisplayDateStart = now.AddDays(1);
            TB1.DisplayDateEnd = now.AddYears(1);
            Pacient6.DisplayDateStart = now.AddYears(-100);
            Pacient6.DisplayDateEnd = now;
            // Pacient6.Style.Resources.IsReadOnly = true;
            DataGrid1.CanUserAddRows = false;
            DataGrid1.CanUserDeleteRows = false;
            DataGrid1.CanUserResizeColumns = false;
            DataGrid1.CanUserResizeRows = false;
            DataGrid1.CanUserReorderColumns = false;
            DataGrid1.IsReadOnly = true;
            dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра  WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");
            DataGrid1.ItemsSource = dt_user.DefaultView;
            whatTable = "Талон";
            cher = "L";
            oper = "Plus";


            //dt_user = Select("SELECT Id_Специальности, Наименование FROM [dbo].[Специальность]"); // получаем данные из таблицы
            DropShadowEffect shd = new DropShadowEffect();
            //DataGrid1.ItemsSource = dt_user.DefaultView;
            shd.ShadowDepth = 0;
            shd.Opacity = 1;
            shd.BlurRadius = 6;
            shd.Color = Colors.DarkOrange;
            SelAdd.Effect = shd;
            SelAdd.Foreground = Brushes.DarkOrange;
        }
        public void load()
        {
            if (cher == "L")
            {
                (DataGrid1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

                (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "HH:mm";
                (DataGrid1.Columns[5] as DataGridTextColumn).Width = 130;
                (DataGrid1.Columns[7] as DataGridTextColumn).Width = 95;
                (DataGrid1.Columns[6] as DataGridTextColumn).Width = 130;
                DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                DataGrid1.Columns[1].Visibility = Visibility.Hidden;
                vt.Visibility = Visibility.Hidden;
                whatTable = "Талон";
                GridVrach.Visibility = Visibility.Hidden;
                Spec.Visibility = Visibility.Hidden;
                GridPac.Visibility = Visibility.Hidden;
                GridMed.Visibility = Visibility.Hidden;
                GridTalon.Visibility = Visibility.Visible;

            }
            cher = "";
        }
        public void OnPasteCommand(object sender, ExecutedRoutedEventArgs e)
        {
            //здесь можно задать условия запрета вставки, если ничего не писать, то просто запрещается вставка
        }
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

        public DataTable Selectt(string selectSQL, int id) // функция подключения к базе данных и обработка запросов
        {

            DataTable dataTable = new DataTable("dataBase");                // создаём таблицу в приложении   // подключаемся к базе данных
            SqlConnection sqlConnection = new SqlConnection("server=DESKTOP-JRUISRI\\SQLEXPRESS;Trusted_Connection=Yes;DataBase=TRPO;");
            sqlConnection.Open();                                           // открываем базу данных
            SqlCommand sqlCommand = sqlConnection.CreateCommand();          // создаём команду

            sqlCommand.CommandText = selectSQL;// присваиваем команде текст
            sqlCommand.Parameters.AddWithValue("@am", id);
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand); // создаём обработчик
            sqlDataAdapter.Fill(dataTable);                                 // возращаем таблицу с результатом
            sqlConnection.Close();
            return dataTable;

        }
        public void insert()
        { //Добавление метод
            SqlConnection sqlConnection1 =
            new SqlConnection("server=DESKTOP-JRUISRI\\SQLEXPRESS;Trusted_Connection=Yes;DataBase=TRPO;");

            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            if (whatTable == "Специальность")
            {
                cmd.CommandText = "Insert into Специальность (Наименование)values(@nm)";
                cmd.Parameters.AddWithValue("@nm", TBS.Text.Trim());
            }

            if (whatTable == "Медсестра")
            {
                cmd.CommandText = "Insert into Медсестра (Имя, Фамилия, Отчество)values(@nm,@fm,@ot)";
                cmd.Parameters.AddWithValue("@nm", Med1.Text);
                cmd.Parameters.AddWithValue("@fm", Med2.Text);
                cmd.Parameters.AddWithValue("@ot", Med3.Text);
            }
            if (whatTable == "Участок")
            {
                cmd.CommandText = "Insert into Участок (Наименование)values(@nm)";
                cmd.Parameters.AddWithValue("@nm", TBS.Text.Trim());
            }
            if (whatTable == "Талон")
            {
                cmd.CommandText = "Insert into Талон (Id_Медсестры, Id_Пациента, Id_Врача, [Время приёма], [Дата приема])values(@id3,@id1,@id2,@vr,@dt)";
                cmd.Parameters.AddWithValue("@id3", id3);
                cmd.Parameters.AddWithValue("@id1", id1);
                cmd.Parameters.AddWithValue("@id2", id2);
                DateTime dt = new DateTime(2021, 1, 21, h, m, 0);

                cmd.Parameters.AddWithValue("@vr", dt);
                cmd.Parameters.AddWithValue("@dt", TB1.Text);
            }
            if (whatTable == "Врач")
            {
                cmd.CommandText = "Insert into Врач (Фамилия,Имя, Отчество, Id_Специальности, телефон, кабинет)values(@id3,@id1,@id2,@vr,@dt,@dy)";
                cmd.Parameters.AddWithValue("@id3", Vrach1.Text.Trim());
                cmd.Parameters.AddWithValue("@id1", Vrach2.Text.Trim());
                cmd.Parameters.AddWithValue("@id2", Vrach3.Text.Trim());
                cmd.Parameters.AddWithValue("@vr", id1);
                cmd.Parameters.AddWithValue("@dt", Vrach5.Text.Trim());
                cmd.Parameters.AddWithValue("@dy", int.Parse(Vrach6.Text.Trim()));
            }
            if (whatTable == "Пациент")
            {
                cmd.CommandText = "Insert into Пациент (Фамилия,Имя, Отчество, Телефон, Адрес, [Дата Рождения], Id_Участка)values(@id3,@id1,@id2,@vr,@dt,@dr,@dy)";
                cmd.Parameters.AddWithValue("@id3", Pacient1.Text.Trim());
                cmd.Parameters.AddWithValue("@id1", Pacient2.Text.Trim());
                cmd.Parameters.AddWithValue("@id2", Pacient3.Text.Trim());
                cmd.Parameters.AddWithValue("@vr", Pacient4.Text);
                cmd.Parameters.AddWithValue("@dt", Pacient5.Text);
                cmd.Parameters.AddWithValue("@dr", Pacient6.Text);
                cmd.Parameters.AddWithValue("@dy", id1);
            }
            cmd.Connection = sqlConnection1;

            sqlConnection1.Open();
            cmd.ExecuteNonQuery();
            sqlConnection1.Close();


        }
        string s;
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Добавление
            try
            {
                if (whatTable == "Медсестра")
                {
                    if (whatTable == "Медсестра" && Med1.Text != "" && Med2.Text != "" && Med3.Text != "")
                    {
                        insert();
                        dt_user = Select("SELECT Имя, Фамилия, Отчество FROM [dbo].[Медсестра]"); // получаем данные из таблицы
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                        Med1.Clear();
                        Med2.Clear();
                        Med3.Clear();
                    }
                    else
                        MessageBox.Show("Данные введены некорректно либо не введены");
                }
                if (whatTable == "Участок")
                {
                    if (whatTable == "Участок" && TBS.Text != "")
                    {
                        insert();
                        dt_user = Select("SELECT Id_Участка, Наименование FROM [dbo].[Участок]  WHERE Исключен is NULL"); // получаем данные из таблицы

                        DataGrid1.ItemsSource = dt_user.DefaultView;
                        DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                        TBS.Clear();
                    }
                    else
                        MessageBox.Show("Данные введены некорректно либо не введены");
                }
                if (whatTable == "Специальность")
                {
                    if (whatTable == "Специальность" && TBS.Text != "")
                    {
                        insert();
                        dt_user = Select("SELECT Id_Специальности, Наименование FROM [dbo].[Специальность] WHERE Состояние is NULL"); // получаем данные из таблицы

                        DataGrid1.ItemsSource = dt_user.DefaultView;
                        DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                        TBS.Clear();
                    }
                    else MessageBox.Show("Данные введены некорректно либо не введены");
                }

                if (whatTable == "Талон")
                {
                    if (TB1.Text != "" && TB2.Text != "" && TB3.Text != "" && TB4.Text != "")
                    {
                        Boolean ch = true;
                        for (int i = 0; i < DataGrid1.Items.Count; i++)
                        {

                            DataRowView row1 = (DataRowView)DataGrid1.Items[i];
                            string s = TB2.Text; string s1 = TB1.Text + " 0:00:00";
                            string r = row1["Время приёма"].ToString();
                            string r1 = row1["Дата приема"].ToString();                       //проверка по времени
                            if (r.IndexOf(s) != -1 && r1.IndexOf(s1) != -1)
                            {


                                int pac = (int)row1["Id_Пациента"];
                                if (pac == id1) { MessageBox.Show("На это время вы уже заказали талон"); ch = false; break; }
                            }
                        }
                        if (ch)
                        {
                            insert();
                            dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра  WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");

                            DataGrid1.ItemsSource = dt_user.DefaultView;
                            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                            DataGrid1.Columns[1].Visibility = Visibility.Hidden;
                            (DataGrid1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

                            (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "HH:mm";
                            (DataGrid1.Columns[5] as DataGridTextColumn).Width = 130;
                            (DataGrid1.Columns[7] as DataGridTextColumn).Width = 95;
                            (DataGrid1.Columns[6] as DataGridTextColumn).Width = 130;
                            TB1.Text = "";
                            TB2.Clear();
                            TB3.Clear();
                            TB4.Clear();
                           
                        }

                    }
                    else MessageBox.Show("Введите все данные");
                }
                bool have = false;
                if (whatTable == "Врач")
                {
                    if (whatTable == "Врач" && Vrach1.Text != "" && Vrach2.Text != "" && Vrach3.Text != "" && Vrach4.Text != "" && Vrach5.Text != "" && Vrach6.Text != "" && Vrach5.Text[0] == '+')
                    {
                        dt_user = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Специальность.Наименование as [Специальность],Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности AND Уволен is NULL");


                        for (int i = 0; i < dt_user.Rows.Count; i++)
                        {
                            string s = (string)dt_user.Rows[i][5];
                            if (s.IndexOf(Vrach5.Text) != (-1))
                            {
                                //Проверка на наличие врача  в базе
                                have = true;
                            }
                        }

                        if (!have)
                        {
                            insert();
                            dt_user = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Специальность.Наименование as [Специальность],Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности AND Уволен is NULL");

                            DataGrid1.ItemsSource = dt_user.DefaultView;
                            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                            vt.Visibility = Visibility.Hidden;
                            Vrach1.Text = "";
                            Vrach2.Text = "";
                            Vrach3.Text = "";
                            Vrach4.Text = "";
                            Vrach5.Text = "";
                            Vrach6.Text = "";
                        }
                        else MessageBox.Show("Врач с таким номером уже есть");

                    }
                    else MessageBox.Show("Данные введены некорректно либо не введены");
                }

                if (whatTable == "Пациент")
                {
                    if (whatTable == "Пациент" && Pacient1.Text != "" && Pacient2.Text != "" && Pacient3.Text != "" && Pacient4.Text != "" && Pacient5.Text != "" && Pacient6.Text != "" && Pacient7.Text != "" && Pacient4.Text[0] == '+')
                    {
                        dt_user = Select("SELECT Пациент.Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка and Пациент.Исключен is NULL");

                        for (int i = 0; i < dt_user.Rows.Count; i++)
                        {
                            string s = (string)dt_user.Rows[i][4];
                            if (s.IndexOf(Pacient4.Text) != (-1))
                            {

                                have = true;
                            }
                        }
                        if (!have)
                        {
                            insert();
                            dt_user = Select("SELECT Пациент.Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка AND Пациент.Исключен is NULL");
                            DataGrid1.ItemsSource = dt_user.DefaultView;
                            vt.Visibility = Visibility.Hidden;
                            // DataGrid1.Columns[1].Visibility = Visibility.Hidden;
                            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                            (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

                        }
                        else MessageBox.Show("Пациент с таким номером уже есть");
                    }
                    else MessageBox.Show("Данные введены некорректно либо не введены");
                }
                Cleaning2();
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }

        }
        Boolean tl = false; string kek = "";
        public void delete(object ind)
        {
            //удаление метод
            System.Data.SqlClient.SqlConnection sqlConnection1 =
            new System.Data.SqlClient.SqlConnection("server=DESKTOP-JRUISRI\\SQLEXPRESS;Trusted_Connection=Yes;DataBase=TRPO;");

            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
            cmd.CommandType = System.Data.CommandType.Text;

            if (whatTable == "Специальность")                           //специальность
            {
                cmd.CommandText = "UPDATE Специальность SET Состояние = @av WHERE Id_Специальности = @am";
                cmd.Parameters.AddWithValue("@av", "Исключена");
                cmd.Parameters.AddWithValue("@am", ind); //сверяет переданный id  и по нему удаляет
                tl = true;
            }
            if (whatTable == "Участок")
            {
                cmd.CommandText = "UPDATE Участок SET Участок.Исключен = @dv WHERE Id_Участка = @dm";
                cmd.Parameters.AddWithValue("@dv", "Исключен");
                cmd.Parameters.AddWithValue("@dm", ind); //сверяет переданный id  и по нему удаляет
                tl = true;
            }
            if (whatTable == "Врач" || kek == "Врач")                                   //Врач
            {
                cmd.CommandText = "UPDATE  Врач Set Врач.Уволен = @uv WHERE Id_Врача = @nm";
                cmd.Parameters.AddWithValue("@uv", "Уволен");
                cmd.Parameters.AddWithValue("@nm", ind); //сверяет переданный id  и по нему удаляет
                tl = true;

            }
            if (whatTable == "Пациент" || kek == "Пациент")                                   //Врач
            {
                cmd.CommandText = "UPDATE  Пациент Set Пациент.Исключен = @pv WHERE Id_Пациента = @pm";
                cmd.Parameters.AddWithValue("@pv", "Исключен");
                cmd.Parameters.AddWithValue("@pm", ind); //сверяет переданный id  и по нему удаляет
                tl = true;

            }
            if (whatTable == "Талон")
            {

                cmd.CommandText = "DELETE FROM Талон WHERE Id_Талона = @sv";
                cmd.Parameters.AddWithValue("@sv", ind);
            }
            if (whatTable == "Медсестра")
            {
                cmd.CommandText = "UPDATE  Медсестра Set Медсестра.Уволен = @uv, Медсестра.Логин = NULL,Медсестра.Пароль = Null WHERE Id_Медсестры = @nm";
                cmd.Parameters.AddWithValue("@uv", "Уволен");
                cmd.Parameters.AddWithValue("@nm", ind);
            }
            cmd.Connection = sqlConnection1;

            sqlConnection1.Open();
            cmd.ExecuteNonQuery();
            sqlConnection1.Close();
            if (tl)
            {
                deleteT(ind);
            }

        }
        public void deleteT(object ind)
        {

            System.Data.SqlClient.SqlConnection sqlConnection1 =
            new System.Data.SqlClient.SqlConnection("server=DESKTOP-JRUISRI\\SQLEXPRESS;Trusted_Connection=Yes;DataBase=TRPO;");

            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
            cmd.CommandType = System.Data.CommandType.Text;

            if (whatTable == "Пациент" || kek == "Пациент")
            {
                cmd.Parameters.AddWithValue("@sm", ind);
                cmd.CommandText = "DELETE FROM Талон WHERE Id_Пациента = @sm AND [Дата приема]>@sv";
                cmd.Parameters.AddWithValue("@sv", DateTime.Now);

                // cmd.CommandText = "DELETE FROM [dbo].Талон WHERE Id_Пациента = @pk AND [Дата приема]> GETDATE()";
                //// cmd.Parameters.AddWithValue("@pv", DateTime.Now);
                // cmd.Parameters.AddWithValue("@pk", ind);
            }
            if (whatTable == "Участок" && kek == "")
            {
                cmd.CommandText = "UPDATE  Пациент Set Пациент.Исключен = @kv WHERE Id_Участка = @km";
                cmd.Parameters.AddWithValue("@kv", "Исключен");
                cmd.Parameters.AddWithValue("@km", ind); //сверяет переданный id  и по нему удаляет

            }
            if (whatTable == "Врач" || kek == "Врач")
            {
                cmd.CommandText = "DELETE FROM Талон WHERE Id_Врача = @sm AND [Дата приема]>@sv";
                cmd.Parameters.AddWithValue("@sv", DateTime.Now);
                cmd.Parameters.AddWithValue("@sm", ind);
            }
            if (whatTable == "Специальность" && kek == "")
            {
                cmd.CommandText = "UPDATE Врач Set Врач.Уволен = @uv WHERE Врач.Id_Специальности = @nm";
                cmd.Parameters.AddWithValue("@uv", "Уволен");
                cmd.Parameters.AddWithValue("@nm", ind);
            }
            cmd.Connection = sqlConnection1;

            sqlConnection1.Open();
            cmd.ExecuteNonQuery();


            sqlConnection1.Close();
            if (whatTable == "Специальность" && kek == "")
            {//!
                dt_user = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Врач.Id_Специальности,Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности");

                for (int i = 0; i < dt_user.Rows.Count - 1; i++)
                {
                    if ((int)dt_user.Rows[i][4] == (int)ind)
                    {
                        id1 = (int)dt_user.Rows[i][0];
                        kek = "Врач";
                        delete(id1);
                    }
                }
            }
            if (whatTable == "Участок" && kek == "")
            {//! //пациенты
                dt_user = Select("SELECT Пациент.Id_Пациента,Пациент.Id_Участка, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка");

                for (int i = 0; i < dt_user.Rows.Count - 1; i++)
                {
                    if ((int)dt_user.Rows[i][1] == (int)ind)
                    {
                        id1 = (int)dt_user.Rows[i][0];
                        kek = "Пациент";
                        delete(id1);
                    }
                }
            }
            tl = false; kek = "";
        }
        int index1;

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {    //Удаление
           
            try
            {
                if (whatTable == "Медсестра")
                {
                    index1 = DataGrid1.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid1.Items[index1];   /////////////////Meds
                    MessageBoxResult result = MessageBox.Show("Вы действительно хотите удалить медсестру \"" + row["Фамилия"].ToString() +" " + row["Имя"].ToString()+  "\"?", "Удаление", MessageBoxButton.YesNo);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:
                            id2 = (int)row["Id_Медсестры"];
                    if (id2 != nowid)
                    {
                        delete(id2);
                        dt_user = Select("SELECT id_Медсестры,Имя, Фамилия, Отчество FROM [dbo].[Медсестра] WHERE Медсестра.Уволен is NULL"); // получаем данные из таблицы
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                        DataGrid1.Columns[0].Visibility = Visibility.Hidden; 
                    }
                    else MessageBox.Show("Для удаления данной медсестры нужно выйти из учетной записи");
                            break;
                        case MessageBoxResult.No:
                            MessageBox.Show("Удаление отменено", "Отмена");
                            break;
                    }
                }
                if (whatTable == "Специальность")
                {
                    index1 = DataGrid1.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid1.Items[index1];   /////////////////Specialn
                    MessageBoxResult result = MessageBox.Show("Вы действительно хотите удалить специальность \"" + row["Наименование"].ToString() + "\"?", "Удаление", MessageBoxButton.YesNo);
                    switch (result)
                    {
                   case MessageBoxResult.Yes:
                            id2 = (int)row["Id_Специальности"];
                    delete(id2);

                    dt_user = Select("SELECT Id_Специальности, Наименование FROM [dbo].[Специальность] WHERE Состояние is NULL"); // получаем данные из таблицы

                    DataGrid1.ItemsSource = dt_user.DefaultView;
                    DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                            break;
                        case MessageBoxResult.No:
                            MessageBox.Show("Удаление отменено", "Отмена");
                            break;
                    }
                }
                if (whatTable == "Талон")
                {

                    index1 = DataGrid1.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid1.Items[index1];   /////////////////Talon
                    MessageBoxResult result = MessageBox.Show("Вы действительно хотите удалить талон для пациента \"" + row["Пациент"].ToString() + "\"?", "Удаление", MessageBoxButton.YesNo);
                    switch (result) {
                        case MessageBoxResult.Yes:
                            id2 = (int)row["Id_Талона"];
                    delete(id2);

                    dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                            (DataGrid1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

                            (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "HH:mm";
                            (DataGrid1.Columns[5] as DataGridTextColumn).Width = 130;
                            (DataGrid1.Columns[7] as DataGridTextColumn).Width = 95;
                            (DataGrid1.Columns[6] as DataGridTextColumn).Width = 130;
                            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                    DataGrid1.Columns[1].Visibility = Visibility.Hidden;
                            break;
                          case  MessageBoxResult.No:
		                    MessageBox.Show("Удаление отменено", "Отмена");
                            break;
                    }

                }
                if (whatTable == "Участок")
                {
                    index1 = DataGrid1.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid1.Items[index1];   /////////////////Uchastok
                    MessageBoxResult result = MessageBox.Show("Вы действительно хотите удалить участок \"" + row["Наименование"].ToString() + "\"?", "Удаление", MessageBoxButton.YesNo);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:
                            id2 = (int)row["Id_Участка"];
                            delete(id2);

                            dt_user = Select("SELECT Id_Участка, Наименование FROM [dbo].[Участок] WHERE Участок.Исключен is NULL"); // получаем данные из таблицы

                            DataGrid1.ItemsSource = dt_user.DefaultView;
                            DataGrid1.Columns[0].Visibility = Visibility.Hidden; break;
                        case MessageBoxResult.No:
                            MessageBox.Show("Удаление отменено", "Отмена");
                            break;
                    }
                    }
                if (whatTable == "Врач")
                {
                    // index1 = DataGrid1.SelectedIndex;
                    // dt_user = Select("SELECT * FROM [dbo].[Врач]");
                    index1 = DataGrid1.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid1.Items[index1];   /////////////////Vrach
                    MessageBoxResult result = MessageBox.Show("Вы действительно хотите удалить врача \"" + row["Фамилия"].ToString() + " " + row["Имя"].ToString() + "\"?", "Удаление", MessageBoxButton.YesNo);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:
                            id2 = (int)row["Id_Врача"];
                    delete(id2);//id
                    dt_user = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Специальность.Наименование as [Специальность],Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности AND Уволен is NULL");

                    DataGrid1.ItemsSource = dt_user.DefaultView;
                    DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                            break;
                        case MessageBoxResult.No:
                            MessageBox.Show("Удаление отменено", "Отмена");
                            break;
                    }
                }
                if (whatTable == "Пациент")
                {
                    // index1 = DataGrid1.SelectedIndex;
                    // dt_user = Select("SELECT * FROM [dbo].[Врач]");
                    index1 = DataGrid1.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid1.Items[index1];   /////////////////Pacient
                    MessageBoxResult result = MessageBox.Show("Вы действительно хотите удалить пациента \"" + row["Фамилия"].ToString() +" " +row["Имя"].ToString() + "\"?", "Удаление", MessageBoxButton.YesNo);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:
                            id2 = (int)row["Id_Пациента"];
                    delete(id2);//id
                    dt_user = Select("SELECT Пациент.Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка and Пациент.Исключен is NULL");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                    DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                    (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";break;
                        case MessageBoxResult.No:
                            MessageBox.Show("Удаление отменено", "Отмена");
                            break;
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void Image_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (vt.Visibility == Visibility.Hidden) vt.Visibility = Visibility.Visible;
            else vt.Visibility = Visibility.Hidden;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        { //Врач
            Cleaning();
            filt.Visibility = Visibility.Hidden;
            ButFilt.Visibility = Visibility.Visible;
            whatTable = "Врач";
            dt_user = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Специальность.Наименование as [Специальность],Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности AND Уволен is NULL");

            DataGrid1.ItemsSource = dt_user.DefaultView;
            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
            vt.Visibility = Visibility.Hidden;

            GridTalon.Visibility = Visibility.Hidden;
            Spec.Visibility = Visibility.Hidden;
            GridPac.Visibility = Visibility.Hidden;
            GridMed.Visibility = Visibility.Hidden;
            GridVrach.Visibility = Visibility.Visible;

            DataGrid2.ItemsSource = null;
            DataGrid2.Items.Refresh();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            //Пациент
            Cleaning();
            filt.Visibility = Visibility.Hidden;
            ButFilt.Visibility = Visibility.Visible;
            dt_user = Select("SELECT Пациент.Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка and Пациент.Исключен is NULL");
            DataGrid1.ItemsSource = dt_user.DefaultView;
            vt.Visibility = Visibility.Hidden;

            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
            (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";
            whatTable = "Пациент";
            GridTalon.Visibility = Visibility.Hidden;
            Spec.Visibility = Visibility.Hidden;
            GridVrach.Visibility = Visibility.Hidden;
            GridMed.Visibility = Visibility.Hidden;
            GridPac.Visibility = Visibility.Visible;
            DataGrid2.ItemsSource = null;
            DataGrid2.Items.Refresh();
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            DataGrid2.ItemsSource = null;
            DataGrid2.Items.Refresh();
            //Талон
            Cleaning();
            filt.Visibility = Visibility.Hidden;
            ButFilt.Visibility = Visibility.Visible;
            dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");
            DataGrid1.ItemsSource = dt_user.DefaultView;
            (DataGrid1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

            (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "HH:mm";
            (DataGrid1.Columns[5] as DataGridTextColumn).Width = 130;
            (DataGrid1.Columns[7] as DataGridTextColumn).Width = 95;
            (DataGrid1.Columns[6] as DataGridTextColumn).Width = 130;
            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
            DataGrid1.Columns[1].Visibility = Visibility.Hidden;
            vt.Visibility = Visibility.Hidden;
            whatTable = "Талон";
            GridVrach.Visibility = Visibility.Hidden;
            Spec.Visibility = Visibility.Hidden;
            GridPac.Visibility = Visibility.Hidden;
            GridMed.Visibility = Visibility.Hidden;
            GridTalon.Visibility = Visibility.Visible;
            //DataRowView row = (DataRowView)DataGrid1.Items[0];
            //MessageBox.Show(row["Фамилия медсестры"].ToString());
            //dt_user = Select("SELECT Талон.[Время приёма], Талон.[Дата приема] as [Дата приема] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры");
            //List list = new List();
            //list. = dt_user.DefaultView;

        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            DataGrid2.ItemsSource = null;
            DataGrid2.Items.Refresh();
            //Специальность
            dt_user = Select("SELECT Id_Специальности, Наименование FROM [dbo].[Специальность] WHERE Состояние is NULL"); // получаем данные из таблицы
          
            
            Cleaning();
            filt.Visibility = Visibility.Hidden;
            ButFilt.Visibility = Visibility.Hidden;
            DataGrid1.ItemsSource = dt_user.DefaultView;
            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
            vt.Visibility = Visibility.Hidden;
            whatTable = "Специальность";
            GridTalon.Visibility = Visibility.Hidden;
            GridVrach.Visibility = Visibility.Hidden;
            GridPac.Visibility = Visibility.Hidden;
            GridMed.Visibility = Visibility.Hidden;
            Spec.Visibility = Visibility.Visible;
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            DataGrid2.ItemsSource = null;
            DataGrid2.Items.Refresh();
            //Участок
            Cleaning();
            filt.Visibility = Visibility.Hidden;
            ButFilt.Visibility = Visibility.Hidden;
            dt_user = Select("SELECT Id_Участка,Наименование FROM [dbo].[Участок] WHERE Участок.Исключен is NULL"); // получаем данные из таблицы
            DataGrid1.ItemsSource = dt_user.DefaultView;
            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
            vt.Visibility = Visibility.Hidden;
            whatTable = "Участок";
            GridTalon.Visibility = Visibility.Hidden;
            GridVrach.Visibility = Visibility.Hidden;
            GridPac.Visibility = Visibility.Hidden;
            GridMed.Visibility = Visibility.Hidden;
            Spec.Visibility = Visibility.Visible;
        }

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            DataGrid2.ItemsSource = null;
            DataGrid2.Items.Refresh();
            //Медсестра
            Cleaning();
            filt.Visibility = Visibility.Hidden;
            ButFilt.Visibility = Visibility.Hidden;
            dt_user = Select("SELECT id_Медсестры,Имя, Фамилия, Отчество FROM [dbo].[Медсестра] WHERE Медсестра.Уволен is NULL"); // получаем данные из таблицы
            DataGrid1.ItemsSource = dt_user.DefaultView;
            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
            vt.Visibility = Visibility.Hidden;
            whatTable = "Медсестра";
            GridTalon.Visibility = Visibility.Hidden;
            GridVrach.Visibility = Visibility.Hidden;
            GridPac.Visibility = Visibility.Hidden;
            Spec.Visibility = Visibility.Hidden;
            GridMed.Visibility = Visibility.Visible;
        }



     

        public void Vremya()
        {
            Vr_800.IsEnabled = true; Vr_800.Foreground = Brushes.White;
            Vr_825.IsEnabled = true; Vr_825.Foreground = Brushes.White;
            Vr_850.IsEnabled = true; Vr_850.Foreground = Brushes.White;
            Vr_915.IsEnabled = true; Vr_915.Foreground = Brushes.White;
            Vr_940.IsEnabled = true; Vr_940.Foreground = Brushes.White;
            Vr_1005.IsEnabled = true; Vr_1005.Foreground = Brushes.White;
            Vr_1030.IsEnabled = true; Vr_1030.Foreground = Brushes.White;
            Vr_1055.IsEnabled = true; Vr_1055.Foreground = Brushes.White;
            Vr_1120.IsEnabled = true; Vr_1120.Foreground = Brushes.White;
            Vr_1145.IsEnabled = true; Vr_1145.Foreground = Brushes.White;
            Vr_1210.IsEnabled = true; Vr_1210.Foreground = Brushes.White;
            Vr_1235.IsEnabled = true; Vr_1235.Foreground = Brushes.White;
            Vr_1300.IsEnabled = true; Vr_1300.Foreground = Brushes.White;
            Vr_1400.IsEnabled = true; Vr_1400.Foreground = Brushes.White;
            Vr_1425.IsEnabled = true; Vr_1425.Foreground = Brushes.White;
            Vr_1450.IsEnabled = true; Vr_1450.Foreground = Brushes.White;
            Vr_1515.IsEnabled = true; Vr_1515.Foreground = Brushes.White;
            Vr_1540.IsEnabled = true; Vr_1540.Foreground = Brushes.White;
            Vr_1605.IsEnabled = true; Vr_1605.Foreground = Brushes.White;
            Vr_1630.IsEnabled = true; Vr_1630.Foreground = Brushes.White;
            //  DataRowView row = (DataRowView)DataGrid1.Items[1];   /////////////////
            // MessageBox.Show(row["Дата приема"].ToString());
            dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента, Врач.Id_Врача,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");

            for (int i = 0; i < dt_user.Rows.Count; i++)
            {
               
                //MessageBox.Show(row["Дата приема"].ToString());
                string s = TB1.Text;
                string ss1 = dt_user.Rows[i][7].ToString();
                int Idp = (int)dt_user.Rows[i][1];
                int ID = (int)dt_user.Rows[i][2];
                if (oper == "Edit")
                {
                    index1 = DataGrid1.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid1.Items[index1];   /////////////////
                    int idPac = (int)row["Id_Пациента"];
                    if (ss1.IndexOf(s) != -1 && ID == id2 && idPac != Idp)
                    {
                        string[] ss = dt_user.Rows[i][6].ToString().Split(' ');
                      //  MessageBox.Show(ss[1]);
                        switch (ss[1].ToString())
                        {
                            case "8:00:00": Vr_800.IsEnabled = false; Vr_800.Foreground = Brushes.Black; break;
                            case "8:25:00": Vr_825.IsEnabled = false; Vr_825.Foreground = Brushes.Black; break;
                            case "8:50:00": Vr_850.IsEnabled = false; Vr_850.Foreground = Brushes.Black; break;
                            case "9:15:00": Vr_915.IsEnabled = false; Vr_915.Foreground = Brushes.Black; break;
                            case "9:40:00": Vr_940.IsEnabled = false; Vr_940.Foreground = Brushes.Black; break;
                            case "10:05:00": Vr_1005.IsEnabled = false; Vr_1005.Foreground = Brushes.Black; break;
                            case "10:30:00": Vr_1030.IsEnabled = false; Vr_1030.Foreground = Brushes.Black; break;
                            case "10:55:00": Vr_1055.IsEnabled = false; Vr_1055.Foreground = Brushes.Black; break;
                            case "11:20:00": Vr_1120.IsEnabled = false; Vr_1120.Foreground = Brushes.Black; break;
                            case "11:45:00": Vr_1145.IsEnabled = false; Vr_1145.Foreground = Brushes.Black; break;
                            case "12:10:00": Vr_1210.IsEnabled = false; Vr_1210.Foreground = Brushes.Black; break;
                            case "12:35:00": Vr_1235.IsEnabled = false; Vr_1235.Foreground = Brushes.Black; break;
                            case "13:00:00": Vr_1300.IsEnabled = false; Vr_1300.Foreground = Brushes.Black; break;
                            case "14:00:00": Vr_1400.IsEnabled = false; Vr_1400.Foreground = Brushes.Black; break;
                            case "14:25:00": Vr_1425.IsEnabled = false; Vr_1425.Foreground = Brushes.Black; break;
                            case "14:50:00": Vr_1450.IsEnabled = false; Vr_1450.Foreground = Brushes.Black; break;
                            case "15:15:00": Vr_1515.IsEnabled = false; Vr_1515.Foreground = Brushes.Black; break;
                            case "15:40:00": Vr_1540.IsEnabled = false; Vr_1540.Foreground = Brushes.Black; break;
                            case "16:05:00": Vr_1605.IsEnabled = false; Vr_1605.Foreground = Brushes.Black; break;
                            case "16:30:00": Vr_1630.IsEnabled = false; Vr_1630.Foreground = Brushes.Black; break;

                        }
                    }
                }
                else
                {
                    if (ss1.IndexOf(s) != -1 && ID == id2)
                    {
                        string[] ss = dt_user.Rows[i][6].ToString().Split(' ');
                        //MessageBox.Show(ss[1]);
                        switch (ss[1].ToString())
                        {
                            case "8:00:00": Vr_800.IsEnabled = false; Vr_800.Foreground = Brushes.Black; break;
                            case "8:25:00": Vr_825.IsEnabled = false; Vr_825.Foreground = Brushes.Black; break;
                            case "8:50:00": Vr_850.IsEnabled = false; Vr_850.Foreground = Brushes.Black; break;
                            case "9:15:00": Vr_915.IsEnabled = false; Vr_915.Foreground = Brushes.Black; break;
                            case "9:40:00": Vr_940.IsEnabled = false; Vr_940.Foreground = Brushes.Black; break;
                            case "10:05:00": Vr_1005.IsEnabled = false; Vr_1005.Foreground = Brushes.Black; break;
                            case "10:30:00": Vr_1030.IsEnabled = false; Vr_1030.Foreground = Brushes.Black; break;
                            case "10:55:00": Vr_1055.IsEnabled = false; Vr_1055.Foreground = Brushes.Black; break;
                            case "11:20:00": Vr_1120.IsEnabled = false; Vr_1120.Foreground = Brushes.Black; break;
                            case "11:45:00": Vr_1145.IsEnabled = false; Vr_1145.Foreground = Brushes.Black; break;
                            case "12:10:00": Vr_1210.IsEnabled = false; Vr_1210.Foreground = Brushes.Black; break;
                            case "12:35:00": Vr_1235.IsEnabled = false; Vr_1235.Foreground = Brushes.Black; break;
                            case "13:00:00": Vr_1300.IsEnabled = false; Vr_1300.Foreground = Brushes.Black; break;
                            case "14:00:00": Vr_1400.IsEnabled = false; Vr_1400.Foreground = Brushes.Black; break;
                            case "14:25:00": Vr_1425.IsEnabled = false; Vr_1425.Foreground = Brushes.Black; break;
                            case "14:50:00": Vr_1450.IsEnabled = false; Vr_1450.Foreground = Brushes.Black; break;
                            case "15:15:00": Vr_1515.IsEnabled = false; Vr_1515.Foreground = Brushes.Black; break;
                            case "15:40:00": Vr_1540.IsEnabled = false; Vr_1540.Foreground = Brushes.Black; break;
                            case "16:05:00": Vr_1605.IsEnabled = false; Vr_1605.Foreground = Brushes.Black; break;
                            case "16:30:00": Vr_1630.IsEnabled = false; Vr_1630.Foreground = Brushes.Black; break;

                        }

                    }
                }
            }
        }

        string who;
        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            vr.Visibility = Visibility.Hidden;
            if (TB1.Text != "")
            {
                TB4.Text = "";
                dt2 = Select("SELECT Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка and Пациент.Исключен is NULL");
                DataGrid2.ItemsSource = dt2.DefaultView;
                DataGrid2.Columns[0].Visibility = Visibility.Hidden;
                who = "Пациент";
                
            }
            else MessageBox.Show("Выберите дату");
        }

        private void Button_Click_9(object sender, RoutedEventArgs e)
        {
            vr.Visibility = Visibility.Hidden;
            if (TB1.Text != "")
            {
                if (TB3.Text != "")
                {

                    dt2 = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Специальность.Наименование as [Специальность],Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности and Уволен is NULL");
                    DataGrid2.ItemsSource = dt2.DefaultView;
                    DataGrid2.Columns[0].Visibility = Visibility.Hidden;
                    who = "Врач";

                }
                else MessageBox.Show("Выберите пациента");
            }
            else MessageBox.Show("Выберите дату");
        }

        //private void Button_Click_10(object sender, RoutedEventArgs e)
        //{
        //    who = "Медсестра";
        //    id3 = nowid;
        //    dt2 = Selectt("SELECT Id_Медсестры, Имя, Фамилия, Отчество FROM [dbo].[Медсестра] WHERE id_Медсестры = @am",id3); // получаем данные из таблицы
        //    TB5.Text = dt2.Rows[0][2].ToString();
        //    DataGrid2.ItemsSource = null;
        //    DataGrid2.Items.Refresh();
        //    who = "Медсестра";
             
        //}
        public int id1, id2, id3, h, m;
        public string vremya;
        private void Vr_850_Click(object sender, RoutedEventArgs e)
        {
            h = 8;
            m = 50;
            vremya = "8:50";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_915_Click(object sender, RoutedEventArgs e)
        {
            h = 9;
            m = 15;
            vremya = "9:15";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_940_Click(object sender, RoutedEventArgs e)
        {
            h = 9;
            m = 40;
            vremya = "9:40";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1005_Click(object sender, RoutedEventArgs e)
        {
            h = 10;
            m = 05;
            vremya = "10:05";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1030_Click(object sender, RoutedEventArgs e)
        {
            h = 10;
            m = 30;
            vremya = "10:30";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1055_Click(object sender, RoutedEventArgs e)
        {
            h = 10;
            m = 55;
            vremya = "10:55";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1120_Click(object sender, RoutedEventArgs e)
        {
            h = 11;
            m = 20;
            vremya = "11:20";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1145_Click(object sender, RoutedEventArgs e)
        {
            h = 11;
            m = 45;
            vremya = "11:45";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1210_Click(object sender, RoutedEventArgs e)
        {
            h = 12;
            m = 10;
            vremya = "12:10";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1235_Click(object sender, RoutedEventArgs e)
        {
            h = 12;
            m = 35;
            vremya = "12:35";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1300_Click(object sender, RoutedEventArgs e)
        {
            h = 13;
            m = 00;
            vremya = "13:00";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1400_Click(object sender, RoutedEventArgs e)
        {
            h = 14;
            m = 00;
            vremya = "14:00";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1425_Click(object sender, RoutedEventArgs e)
        {
            h = 14;
            m = 25;
            vremya = "14:25";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1450_Click(object sender, RoutedEventArgs e)
        {
            h = 14;
            m = 50;
            vremya = "14:50";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1515_Click(object sender, RoutedEventArgs e)
        {
            h = 15;
            m = 15;
            vremya = "15:15";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1540_Click(object sender, RoutedEventArgs e)
        {
            h = 15;
            m = 40;
            vremya = "15:40";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_1605_Click(object sender, RoutedEventArgs e)
        {
            h = 16;
            m = 05;
            vremya = "16:05";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void TB2_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void TB3_TextChanged(object sender, TextChangedEventArgs e)
        {


        }

        private void Vrach1_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if (Vrach1.Text.Length == 1)
            {
                Vrach1.Text = Vrach1.Text.ToUpper();
                Vrach1.Select(Vrach1.Text.Length, 0);
            }
        }

        private void Button_Click_12(object sender, RoutedEventArgs e)
        {
            dt2 = Select("SELECT  Id_Специальности, Наименование FROM [dbo].Специальность WHERE Состояние is NULL");
            DataGrid2.ItemsSource = dt2.DefaultView;
            DataGrid2.Columns[0].Visibility = Visibility.Hidden;
            who = "Специальность";
        }

        private void Button_Click_13(object sender, RoutedEventArgs e)
        {
            dt2 = Select("SELECT Id_Участка, Наименование FROM [dbo].[Участок] WHERE Исключен is NULL"); // получаем данные из таблицы

            DataGrid2.ItemsSource = dt2.DefaultView;
            DataGrid2.Columns[0].Visibility = Visibility.Hidden;
            who = "Участок";
        }

        private void TB1_CalendarOpened(object sender, RoutedEventArgs e)
        {
            vr.Visibility = Visibility.Hidden;
            TB4.Text = "";
            var minDate = TB1.DisplayDateStart ?? DateTime.MinValue;
            var maxDate = TB1.DisplayDateEnd ?? DateTime.MaxValue;

            for (var d = minDate; d <= maxDate && DateTime.MaxValue > d; d = d.AddDays(1))
            {
                if (d.DayOfWeek == DayOfWeek.Saturday || d.DayOfWeek == DayOfWeek.Sunday)
                {
                    TB1.BlackoutDates.Add(new CalendarDateRange(d));
                }
            }
        }

        private void Vrach1_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Vrach2_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;


            if (Vrach2.Text.Length == 1)
            {
                Vrach2.Text = Vrach2.Text.ToUpper();

                Vrach2.Select(Vrach2.Text.Length, 0);
            }
        }

        private void Vrach3_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if (Vrach3.Text.Length == 1)
            {
                Vrach3.Text = Vrach3.Text.ToUpper();
                Vrach3.Select(Vrach3.Text.Length, 0);
            }
        }

        private void Vrach5_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsDigit(inp))
                e.Handled = true;


            //if (Vrach5.Text.Length == 4)
            //{
            //    Vrach5.Text = Vrach5.Text+"(";
            //    Vrach5.Select(Vrach5.Text.Length, 1);
            //}
            //if (Vrach5.Text.Length == 7)
            //{
            //    Vrach5.Text = Vrach5.Text + ")";
            //    Vrach5.Select(Vrach5.Text.Length, 1);
            //}
            //if (Vrach5.Text.Length == 11)
            //{
            //    Vrach5.Text = Vrach5.Text + "-";
            //    Vrach5.Select(Vrach5.Text.Length, 1);
            //}
            //if (Vrach5.Text.Length == 14)
            //{
            //    Vrach5.Text = Vrach5.Text + "-";
            //    Vrach5.Select(Vrach5.Text.Length, 1);
            //}

        }

        private void Vrach6_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsDigit(inp))
                e.Handled = true;
        }

        private void Vrach6_TextChanged(object sender, TextChangedEventArgs e)
        {


        }

        private void Vrach5_SelectionChanged(object sender, RoutedEventArgs e)
        {
            e.Handled = true;

            if ((sender as TextBox).SelectionLength != 0)
                (sender as TextBox).SelectionLength = 0;
        }

        private void Vrach5_PreviewKeyDown(object sender, KeyEventArgs e)
        {

            if (e.Key == Key.Space)
            {
                e.Handled = true;
            }
            int k = 0;
            if (Vrach5.Text.Length != 0 && Vrach5.Text[0] != '+')
            {
                for (int i = 0; i < Vrach5.Text.Length - 1; i++)
                {
                    if (Vrach5.Text[i] == '+') k++;
                }
                if (k == 0)
                {
                    Vrach5.Text = "+" + Vrach5.Text;
                    Vrach5.Select(Vrach5.Text.Length, 1);
                }
            }
        }

        private void Pacient1_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if (Pacient1.Text.Length == 1)
            {
                Pacient1.Text = Pacient1.Text.ToUpper();
                Pacient1.Select(Pacient1.Text.Length, 0);
            }
        }

        private void Pacient2_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if (Pacient2.Text.Length == 1)
            {
                Pacient2.Text = Pacient2.Text.ToUpper();
                Pacient2.Select(Pacient2.Text.Length, 0);
            }
        }

        private void Pacient3_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if (Pacient3.Text.Length == 1)
            {
                Pacient3.Text = Pacient3.Text.ToUpper();
                Pacient3.Select(Pacient3.Text.Length, 0);
            }
        }

        private void Pacient4_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                e.Handled = true;
            }
            int k = 0;
            if (Pacient4.Text.Length != 0 && Pacient4.Text[0] != '+')
            {
                for (int i = 0; i < Pacient4.Text.Length - 1; i++)
                {
                    if (Pacient4.Text[i] == '+') k++;
                }
                if (k == 0)
                {
                    Pacient4.Text = "+" + Pacient4.Text;
                    Pacient4.Select(Pacient4.Text.Length, 1);
                }
            }
        }

        private void Pacient4_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsDigit(inp))
                e.Handled = true;
        }

        private void Pacient4_SelectionChanged(object sender, RoutedEventArgs e)
        {
            e.Handled = true;

            if ((sender as TextBox).SelectionLength != 0)
                (sender as TextBox).SelectionLength = 0;
        }

        private void Pacient5_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {

        }

        private void Med1_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if (Med1.Text.Length == 1)
            {
                Med1.Text = Med1.Text.ToUpper();
                Med1.Select(Med1.Text.Length, 0);
            }
        }

        private void Med2_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if (Med2.Text.Length == 1)
            {
                Med2.Text = Med2.Text.ToUpper();
                Med2.Select(Med2.Text.Length, 0);
            }
        }

        private void Med3_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if (Med3.Text.Length == 1)
            {
                Med3.Text = Med3.Text.ToUpper();
                Med3.Select(Med3.Text.Length, 0);
            }
        }

        private void TBS_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if (TBS.Text.Length == 1)
            {
                TBS.Text = TBS.Text.ToUpper();
                TBS.Select(TBS.Text.Length, 0);
            }
        }

        public void Edit(int id)
        {                                               //Редактирование
            SqlConnection sqlConnection1 =
            new SqlConnection("server=DESKTOP-JRUISRI\\SQLEXPRESS;Trusted_Connection=Yes;DataBase=TRPO;");

            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            if (whatTable == "Врач")
            {
                cmd.CommandText = "UPDATE Врач SET Фамилия = @id3, Имя = @id1, Отчество = @id2, Id_Специальности = @dt, Телефон = @vr, Кабинет = @kab  WHERE Id_Врача = @am";
                cmd.Parameters.AddWithValue("@id3", Vrach1.Text.Trim());
                cmd.Parameters.AddWithValue("@id1", Vrach2.Text.Trim());
                cmd.Parameters.AddWithValue("@id2", Vrach3.Text.Trim());
                cmd.Parameters.AddWithValue("@kab", int.Parse(Vrach6.Text.Trim()));

                cmd.Parameters.AddWithValue("@am", id);
                cmd.Parameters.AddWithValue("@dt", id1);
                cmd.Parameters.AddWithValue("@vr", Vrach5.Text.Trim());

            }

            if (whatTable == "Пациент")
            {
                cmd.CommandText = "UPDATE Пациент SET Фамилия = @id3, Имя = @id1, Отчество = @id2, Id_Участка = @dt, Телефон = @vr, Адрес = @kab, [Дата Рождения] = @dr  WHERE Id_Пациента = @am";
                cmd.Parameters.AddWithValue("@id3", Pacient1.Text.Trim());
                cmd.Parameters.AddWithValue("@id1", Pacient2.Text.Trim());
                cmd.Parameters.AddWithValue("@id2", Pacient3.Text.Trim());
                cmd.Parameters.AddWithValue("@kab", Pacient5.Text.Trim());
                cmd.Parameters.AddWithValue("@dr", Pacient6.Text);
                cmd.Parameters.AddWithValue("@am", id);
                cmd.Parameters.AddWithValue("@dt", id1);
                cmd.Parameters.AddWithValue("@vr", Pacient4.Text.Trim());
            }
            if (whatTable == "Талон")
            {
                cmd.CommandText = "UPDATE Талон SET Id_Медсестры = @id3, Id_Пациента = @id1, Id_Врача = @id2, [Время приёма] = @dt, [Дата приема] = @vr  WHERE Id_Талона = @am";
                cmd.Parameters.AddWithValue("@id3", id3);
                cmd.Parameters.AddWithValue("@id1", id1);
                cmd.Parameters.AddWithValue("@id2", id2);
                DateTime dt = new DateTime(2021, 1, 21, h, m, 0);
                cmd.Parameters.AddWithValue("@am", id);
                cmd.Parameters.AddWithValue("@dt", dt);
                cmd.Parameters.AddWithValue("@vr", TB1.Text.Trim());
            }
            if (whatTable == "Специальность")
            {
                cmd.CommandText = "UPDATE Специальность SET Наименование = @nm WHERE Id_Специальности = @am";
                cmd.Parameters.AddWithValue("@nm", TBS.Text);
                cmd.Parameters.AddWithValue("@am", id);
            }
            if (whatTable == "Участок")
            {
                cmd.CommandText = "UPDATE Участок SET Наименование = @nm WHERE Id_Участка = @am";
                cmd.Parameters.AddWithValue("@nm", TBS.Text.Trim());
                cmd.Parameters.AddWithValue("@am", id);
            }
            if (whatTable == "Медсестра")
            {
                cmd.CommandText = "UPDATE Медсестра SET Имя = @nm, Фамилия = @fm, Отчество = @om WHERE Id_Медсестры = @am";
                cmd.Parameters.AddWithValue("@nm", Med1.Text.Trim());
                cmd.Parameters.AddWithValue("@fm", Med2.Text.Trim());
                cmd.Parameters.AddWithValue("@om", Med3.Text.Trim());
                cmd.Parameters.AddWithValue("@am", id);
            }
            cmd.Connection = sqlConnection1;

            sqlConnection1.Open();
            cmd.ExecuteNonQuery();
            sqlConnection1.Close();

        }

        private void Button_Click_14(object sender, RoutedEventArgs e)
        {
            //Выбор для редактирования
            try
            {
                index1 = DataGrid1.SelectedIndex;
                DataRowView row = (DataRowView)DataGrid1.Items[index1];
                AcceptBut.Visibility = Visibility.Visible;
                DisAcceptBut.Visibility = Visibility.Visible;
                if (whatTable == "Талон")
                {
                    dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");

                    int idt = (int)row["Id_Талона"];
                    dt_user = Selectt("SELECT Талон.Id_Талона,Медсестра.Id_Медсестры,Пациент.Id_Пациента,Врач.Id_Врача, Медсестра.Фамилия, Врач.Фамилия,Пациент.Фамилия FROM [dbo].[Талон], Врач, Пациент, Медсестра WHERE [dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры AND Талон.Id_Талона = @am", idt);
                    id1 = (int)dt_user.Rows[0][2];
                    id2 = (int)dt_user.Rows[0][3];
                    id3 = (int)dt_user.Rows[0][1];
                    string s = dt_user.Rows[0][6].ToString();
                    TB3.Text = s.Trim();
                    //s = dt_user.Rows[0][4].ToString();
                    //TB5.Text = s.Trim();
                    s = dt_user.Rows[0][5].ToString();
                    TB4.Text = s.Trim();
                    s = row["Дата приема"].ToString();
                    TB1.Text = s.Trim();
                    s = row["Время приёма"].ToString();
                    s = s.Trim();
                    string[] ss = s.Split(' ');
                     ss = ss[1].Split(':');
                    h = int.Parse(ss[0]);
                    m = int.Parse(ss[1]);
                    TB2.Text = ss[0] + ":" + ss[1];
                }

                if (whatTable == "Специальность")
                {

                    string s = (string)row["Наименование"];
                    TBS.Text = s.Trim();
                }
                if (whatTable == "Участок")
                {

                    string s = (string)row["Наименование"];
                    TBS.Text = s.Trim();
                }
                if (whatTable == "Пациент")
                {
                    int b = (int)row["Id_Пациента"];
                    dt_user = Selectt("Select Id_Участка From Пациент Where Id_Пациента = @am ", b);
                    id1 = (int)dt_user.Rows[0][0];

                    string s = (string)row["Фамилия"];
                    Pacient1.Text = s.Trim();
                    s = (string)row["Имя"];
                    Pacient2.Text = s.Trim();
                    s = (string)row["Отчество"];
                    Pacient3.Text = s.Trim();
                    s = (string)row["Адрес"];
                    Pacient5.Text = s.Trim();
                    s = (string)row["Телефон"];
                    Pacient4.Text = s.Trim();
                    s = row["Дата Рождения"].ToString();
                    Pacient6.Text = s.Trim();
                    s = (string)row["Участок"];
                    Pacient7.Text = s.Trim();
                }
                if (whatTable == "Врач")
                {
                    int b = (int)row["Id_Врача"];
                    dt_user = Selectt("Select Id_Специальности From Врач Where Id_Врача = @am ", b);
                    id1 = (int)dt_user.Rows[0][0];

                    string s = (string)row["Фамилия"];
                    Vrach1.Text = s.Trim();
                    s = (string)row["Имя"];
                    Vrach2.Text = s.Trim();
                    s = (string)row["Отчество"];
                    Vrach3.Text = s.Trim();
                    s = (string)row["Специальность"];
                    Vrach4.Text = s.Trim();
                    s = (string)row["Телефон"];
                    Vrach5.Text = s.Trim();

                    Vrach6.Text = Convert.ToString(row["Кабинет"]);
                }
                if (whatTable == "Медсестра")
                {
                    idr = (int)row["Id_Медсестры"];
                    if (nowid != idr)
                    {

                        string s = (string)row["Фамилия"];
                        Med2.Text = s.Trim();
                        s = (string)row["Имя"];
                        Med1.Text = s.Trim();
                        s = (string)row["Отчество"];
                        Med3.Text = s.Trim();
                    }
                    else MessageBox.Show("Для редактирования данной записи необходимо авторизоваться под другим аккаунтом");
                }
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void ExportToExcelAndCsv()
        {
            try
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                Excel.Range rangeToHoldHyperlink;
                Excel.Range CellInstance;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlApp.DisplayAlerts = false;
                //Dummy initialisation to prevent errors.
                rangeToHoldHyperlink = xlWorkSheet.get_Range("A1", Type.Missing);
                CellInstance = xlWorkSheet.get_Range("A1", Type.Missing);
                int j = 1;

                if (whatTable == "Талон")
                {

                    foreach (DataRowView row in DataGrid1.Items)
                    {

                        for (int i = 2; i < DataGrid1.Columns.Count; i++)
                        {
                            Excel.Range Range2 = xlWorkSheet.get_Range("A1");
                            Excel.Range Range3 = xlWorkSheet.get_Range("B1");
                            Excel.Range Range4 = xlWorkSheet.get_Range("C1");
                            Excel.Range Range5 = xlWorkSheet.get_Range("D1");
                            Excel.Range Range6 = xlWorkSheet.get_Range("E1");
                            Excel.Range Range7 = xlWorkSheet.get_Range("F1");
                            Range2.ColumnWidth = 30;
                            Range2.EntireRow.AutoFit();
                            Range3.ColumnWidth = 30;
                            Range3.EntireRow.AutoFit();
                            Range4.ColumnWidth = 30;
                            Range4.EntireRow.AutoFit();
                            Range5.ColumnWidth = 20;
                            Range5.EntireRow.AutoFit();
                            Range6.ColumnWidth = 20;
                            Range6.EntireRow.AutoFit();
                            Range6.ColumnWidth = 18;
                            Range6.EntireRow.AutoFit();
                            Range7.ColumnWidth = 15;
                            Range7.EntireRow.AutoFit();
                            xlWorkSheet.Cells[1, 1] = "Медсестра";
                            xlWorkSheet.Cells[1, 2] = "Пациент";
                            xlWorkSheet.Cells[1, 3] = "Врач";
                            xlWorkSheet.Cells[1, 4] = "Время  приема";
                            xlWorkSheet.Cells[1, 5] = "Дата приема";
                            xlWorkSheet.Cells[1, 6] = "Кабинет";
                            string[] conv;
                            xlWorkSheet.Cells[j + 1, i - 1] = row[i].ToString();
                            if (i == 6) { conv = row[i].ToString().Split(' '); xlWorkSheet.Cells[j + 1, i-1] = conv[0]; }
                            if (i == 5) { conv = row[i].ToString().Split(' '); conv = conv[1].Split(':'); xlWorkSheet.Cells[j + 1, i - 1] = conv[0]+":"+conv[1]; }
                        }
                        j++;
                    }

                }

                if (whatTable == "Врач")
                {
                    foreach (DataRowView row in DataGrid1.Items)
                    {
                        for (int i = 1; i < DataGrid1.Columns.Count; i++)
                        {
                            Excel.Range Range2 = xlWorkSheet.get_Range("A1");
                            Excel.Range Range3 = xlWorkSheet.get_Range("B1");
                            Excel.Range Range4 = xlWorkSheet.get_Range("C1");
                            Excel.Range Range5 = xlWorkSheet.get_Range("D1");
                            Excel.Range Range6 = xlWorkSheet.get_Range("E1");
                            Excel.Range Range7 = xlWorkSheet.get_Range("F1");
                            Range2.ColumnWidth = 20;
                            Range2.EntireRow.AutoFit();
                            Range3.ColumnWidth = 20;
                            Range3.EntireRow.AutoFit();
                            Range4.ColumnWidth = 20;
                            Range4.EntireRow.AutoFit();
                            Range5.ColumnWidth = 20;
                            Range5.EntireRow.AutoFit();
                            Range6.ColumnWidth = 15;
                            Range6.EntireRow.AutoFit();
                            Range7.ColumnWidth = 15;
                            Range7.EntireRow.AutoFit();
                            xlWorkSheet.Cells[1, 1] = "Фамилия";
                            xlWorkSheet.Cells[1, 2] = "Имя";
                            xlWorkSheet.Cells[1, 3] = "Отчество";
                            xlWorkSheet.Cells[1, 4] = "Специальность";
                            xlWorkSheet.Cells[1, 5] = "Телефон";
                            (xlWorkSheet.Cells[j + 1, 5] as Excel.Range).NumberFormat = "@";
                            xlWorkSheet.Cells[1, 6] = "Кабинет";
                            xlWorkSheet.Cells[j + 1, i] = row[i].ToString();
                        }
                        j++;
                    }

                }
                if (whatTable == "Пациент")
                {
                    foreach (DataRowView row in DataGrid1.Items)
                    {
                        for (int i = 1; i < DataGrid1.Columns.Count; i++)
                        {
                            Excel.Range Range2 = xlWorkSheet.get_Range("A1");
                            Excel.Range Range3 = xlWorkSheet.get_Range("B1");
                            Excel.Range Range4 = xlWorkSheet.get_Range("C1");
                            Excel.Range Range5 = xlWorkSheet.get_Range("D1");
                            Excel.Range Range6 = xlWorkSheet.get_Range("E1");
                            Excel.Range Range7 = xlWorkSheet.get_Range("F1");
                            Excel.Range Range8 = xlWorkSheet.get_Range("G1");
                            Range2.ColumnWidth = 20;
                            Range2.EntireRow.AutoFit();
                            Range3.ColumnWidth = 20;
                            Range3.EntireRow.AutoFit();
                            Range4.ColumnWidth = 20;
                            Range4.EntireRow.AutoFit();
                            Range5.ColumnWidth = 20;
                            Range5.EntireRow.AutoFit();
                            Range6.ColumnWidth = 20;
                            Range6.EntireRow.AutoFit();
                            Range7.ColumnWidth = 15;
                            Range7.EntireRow.AutoFit();
                            Range8.ColumnWidth = 15;
                            Range8.EntireRow.AutoFit();
                            xlWorkSheet.Cells[1, 1] = "Фамилия";
                            xlWorkSheet.Cells[1, 2] = "Имя";
                            xlWorkSheet.Cells[1, 3] = "Отчество";
                            xlWorkSheet.Cells[1, 5] = "Дата рождения";
                            xlWorkSheet.Cells[1, 4] = "Телефон";
                            (xlWorkSheet.Cells[j + 1, 4] as Excel.Range).NumberFormat = "@";
                            //не работает (xlWorkSheet.Cells[j + 1, 5] as Excel.Range).NumberFormat = "ДД.ММ.ГГГГ";
                            xlWorkSheet.Cells[1, 6] = "Адрес";
                            xlWorkSheet.Cells[1, 7] = "Участок";
                            string[] conv;

                            xlWorkSheet.Cells[j + 1, i] = row[i].ToString();
                            if (i == 5) { conv = row[i].ToString().Split(' '); xlWorkSheet.Cells[j + 1, i] = conv[0]; }

                        }
                        j++;
                    }
                }
                if (whatTable == "Специальность")
                {
                    foreach (DataRowView row in DataGrid1.Items)
                    {
                        for (int i = 1; i < DataGrid1.Columns.Count; i++)
                        {
                            Excel.Range Range2 = xlWorkSheet.get_Range("A1");


                            Range2.ColumnWidth = 20;
                            Range2.EntireRow.AutoFit();


                            xlWorkSheet.Cells[1, 1] = "Наименование";



                            xlWorkSheet.Cells[j + 1, i] = row[i].ToString();
                        }
                        j++;
                    }
                }
                if (whatTable == "Участок")
                {
                    foreach (DataRowView row in DataGrid1.Items)
                    {
                        for (int i = 1; i < DataGrid1.Columns.Count; i++)
                        {
                            Excel.Range Range2 = xlWorkSheet.get_Range("A1");


                            Range2.ColumnWidth = 20;
                            Range2.EntireRow.AutoFit();


                            xlWorkSheet.Cells[1, 1] = "Наименование";



                            xlWorkSheet.Cells[j + 1, i] = row[i].ToString();
                        }
                        j++;
                    }
                }
                if (whatTable == "Медсестра")
                {
                    foreach (DataRowView row in DataGrid1.Items)
                    {
                        for (int i = 1; i < DataGrid1.Columns.Count; i++)
                        {
                            Excel.Range Range2 = xlWorkSheet.get_Range("A1");
                            Excel.Range Range3 = xlWorkSheet.get_Range("B1");
                            Excel.Range Range4 = xlWorkSheet.get_Range("C1");

                            Range2.ColumnWidth = 20;
                            Range2.EntireRow.AutoFit();
                            Range3.ColumnWidth = 20;
                            Range3.EntireRow.AutoFit();
                            Range4.ColumnWidth = 20;
                            Range4.EntireRow.AutoFit();

                            xlWorkSheet.Cells[1, 1] = "Имя";
                            xlWorkSheet.Cells[1, 2] = "Фамилия";
                            xlWorkSheet.Cells[1, 3] = "Отчество";


                            xlWorkSheet.Cells[j + 1, i] = row[i].ToString();
                        }
                        j++;
                    }
                }
                Excel.Range Range1 = xlWorkSheet.get_Range("A1");
               
                Range1.EntireRow.Font.Size = 14;
                
                Range1.EntireRow.AutoFit();
                Excel.Range tRange = xlWorkSheet.UsedRange;
                tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                SaveFileDialog sv = new SaveFileDialog();
                sv.Filter = "Excel files(*.xls)|*.xls|All files(*.*)|*.*";
                if (sv.ShowDialog() == true)
                {
                    string path = sv.FileName;
                    xlWorkBook.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    MessageBox.Show("Файл сохранен");
                }
                else MessageBox.Show("Сохранение отменено");
                xlWorkBook.Close();
               
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

        int idr;
        private void Button_Click_15(object sender, RoutedEventArgs e)
        {

            index1 = DataGrid1.SelectedIndex;
            DataRowView row = (DataRowView)DataGrid1.Items[index1];
            if (whatTable == "Пациент")
            {
                Boolean have = false;
                if (whatTable == "Пациент" && Pacient1.Text != "" && Pacient2.Text != "" && Pacient3.Text != "" && Pacient4.Text != "" && Pacient5.Text != "" && Pacient6.Text != "" && Pacient7.Text != "" && Pacient4.Text[0] == '+')
                {
                    dt_user = Select("SELECT Пациент.Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка and Пациент.Исключен is NULL");
                    int idd = (int)row["Id_Пациента"];
                    for (int i = 0; i < dt_user.Rows.Count; i++)
                    {
                        int a = (int)dt_user.Rows[i][0];
                        string s = (string)dt_user.Rows[i][4];
                        if (s.IndexOf(Pacient4.Text) != (-1) && a != idd)
                        {

                            have = true;
                        }
                    }
                    if (!have)
                    {
                        idr = idd;
                        Edit(idd);
                        dt_user = Select("SELECT Пациент.Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка AND Пациент.Исключен is NULL");
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                        vt.Visibility = Visibility.Hidden;
                        DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                        (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

                    }
                    else MessageBox.Show("Пациент с таким номером уже есть");
                }
                else MessageBox.Show("Данные введены некорректно либо не введены");
            }
            if (whatTable == "Врач")
            {
                Boolean have = false;
                if (whatTable == "Врач" && Vrach1.Text != "" && Vrach2.Text != "" && Vrach3.Text != "" && Vrach4.Text != "" && Vrach5.Text != "" && Vrach6.Text != "" && Vrach5.Text[0] == '+')
                {
                    dt_user = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Специальность.Наименование as [Специальность],Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности AND Уволен is NULL");
                    int idd = (int)row["Id_Врача"];


                    for (int i = 0; i < dt_user.Rows.Count; i++)
                    {
                        int a = (int)dt_user.Rows[i][0];
                        string s = (string)dt_user.Rows[i][5];
                        if (s.IndexOf(Vrach5.Text) != (-1) && a != idd)
                        {
                            //Проверка на наличие врача  в базе
                            have = true;
                        }
                    }

                    if (!have)
                    {
                        idr = idd;
                        Edit(idd);
                        dt_user = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Специальность.Наименование as [Специальность],Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности AND Уволен is NULL");

                        DataGrid1.ItemsSource = dt_user.DefaultView;
                        DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                        Vrach1.Text = "";
                        Vrach2.Text = "";
                        Vrach3.Text = "";
                        Vrach4.Text = "";
                        Vrach5.Text = "";
                        Vrach6.Text = "";
                        AcceptBut.Visibility = Visibility.Hidden;
                        DisAcceptBut.Visibility = Visibility.Hidden;
                    }
                    else MessageBox.Show("Врач с таким номером уже есть");

                }
                else MessageBox.Show("Данные введены некорректно либо не введены");
            }
            if (whatTable == "Талон")
            {
                if (TB1.Text != "" && TB2.Text != "" && TB3.Text != "" && TB4.Text != "" )
                {
                    Boolean ch = true;
                    dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,Медсестра.Фамилия as [Фамилия медсестры], Пациент.Фамилия as [Фамилия пациента], Врач.Фамилия as [Фамилия врача], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");

                    for (int i = 0; i < dt_user.Rows.Count; i++)
                    {

                        string s = TB2.Text; string s1 = TB1.Text;
                        string r = dt_user.Rows[i][5].ToString();
                        string r1 = dt_user.Rows[i][6].ToString();                    //проверка по времени
                        index1 = DataGrid1.SelectedIndex;
                        /////////////////Talon

                        idr = (int)row["Id_Талона"];
                        int idd = (int)dt_user.Rows[i][0];
                        if (r.IndexOf(s) != -1 && r1.IndexOf(s1) != -1 && id2 != idd)
                        {

                            int pac = (int)dt_user.Rows[i][0];
                            if (pac == id1 && idr != idd) { MessageBox.Show("На это время вы уже заказали талон"); ch = false; break; }
                            TB1.Text = "";
                            TB2.Text = "";
                            TB3.Text = "";
                            TB4.Text = "";
                            
                            AcceptBut.Visibility = Visibility.Hidden;
                            DisAcceptBut.Visibility = Visibility.Hidden;
                        }
                    }
                    if (ch)
                    {

                        idr = (int)row["Id_Талона"];
                        Edit(idr);
                        dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                        (DataGrid1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

                        (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "HH:mm";
                        (DataGrid1.Columns[5] as DataGridTextColumn).Width = 130;
                        (DataGrid1.Columns[7] as DataGridTextColumn).Width = 95;
                        (DataGrid1.Columns[6] as DataGridTextColumn).Width = 130;
                        DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                        DataGrid1.Columns[1].Visibility = Visibility.Hidden;
                        TB1.Text = "";
                        TB2.Clear();
                        TB3.Clear();
                        TB4.Clear();
                       
                        AcceptBut.Visibility = Visibility.Hidden;
                        DisAcceptBut.Visibility = Visibility.Hidden;
                    }
                }
                else MessageBox.Show("Введите все данные");
            }
            if (whatTable == "Участок")
            {
                if (TBS.Text != "")
                {
                    idr = (int)row["Id_Участка"];
                    Edit(idr);
                    dt_user = Select("SELECT Id_Участка, Наименование FROM [dbo].[Участок] WHERE Исключен is NULL"); // получаем данные из таблицы

                    DataGrid1.ItemsSource = dt_user.DefaultView;
                    DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                    TBS.Text = "";
                    AcceptBut.Visibility = Visibility.Hidden;
                    DisAcceptBut.Visibility = Visibility.Hidden;
                }
                else MessageBox.Show("Введите данные");
            }
            if (whatTable == "Специальность")
            {
                if (TBS.Text != "")
                {
                    idr = (int)row["Id_Специальности"];
                    Edit(idr);
                    dt_user = Select("SELECT Id_Специальности, Наименование FROM [dbo].[Специальность] WHERE Состояние is NULL"); // получаем данные из таблицы

                    DataGrid1.ItemsSource = dt_user.DefaultView;
                    DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                    TBS.Text = "";
                    AcceptBut.Visibility = Visibility.Hidden;
                    DisAcceptBut.Visibility = Visibility.Hidden;
                }
                else MessageBox.Show("Введите данные");
            }
            if (whatTable == "Медсестра")
            {
                if (Med1.Text != "" && Med2.Text != "" && Med3.Text != "")
                {
                    idr = (int)row["Id_Медсестры"];


                    Edit(idr);
                    dt_user = Select("SELECT id_Медсестры,Имя, Фамилия, Отчество FROM [dbo].[Медсестра] WHERE Медсестра.Уволен is NULL"); // получаем данные из таблицы
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                    DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                    Med1.Text = "";
                    Med2.Text = "";
                    Med3.Text = "";
                    AcceptBut.Visibility = Visibility.Hidden;
                    DisAcceptBut.Visibility = Visibility.Hidden;

                }
                else MessageBox.Show("Введите все данные");
            }
        }
        string oper = "";
        private void Button_Click_16(object sender, RoutedEventArgs e)
        {
            SelDelete.Foreground = Brushes.White;
            SelEdit.Foreground = Brushes.White;
            SelAdd.Foreground = Brushes.DarkOrange;
            //UIElement uie = new UIElement();
            //uie.Effect = 
            DropShadowEffect shd = new DropShadowEffect();
            shd.ShadowDepth = 0;
            shd.Opacity = 1;
            shd.BlurRadius = 6;
            shd.Color = Colors.DarkOrange;
            SelAdd.Effect = shd;
            DropShadowEffect shd1 = new DropShadowEffect();

            shd1.Opacity = 0;
            SelEdit.Effect = shd1;
            SelDelete.Effect = shd1;
            oper = "Plus";
            EditBut.Visibility = Visibility.Hidden;
            AcceptBut.Visibility = Visibility.Hidden;
            DeleteBut.Visibility = Visibility.Hidden;
            DisAcceptBut.Visibility = Visibility.Hidden;

            AddBut.Visibility = Visibility.Visible;
            Cleaning2();
        }

        private void Button_Click_17(object sender, RoutedEventArgs e)
        {
            DropShadowEffect shd = new DropShadowEffect();
            shd.ShadowDepth = 0;
            shd.Opacity = 1;
            shd.BlurRadius = 6;
            shd.Color = Colors.DarkOrange;
            SelEdit.Effect = shd;

            DropShadowEffect shd1 = new DropShadowEffect();
            SelAdd.Foreground = Brushes.White;
            shd1.Opacity = 0;
            SelAdd.Effect = shd1;
            SelDelete.Effect = shd1;
            SelDelete.Foreground = Brushes.White;
            SelEdit.Foreground = Brushes.DarkOrange;
            oper = "Edit";
            EditBut.Visibility = Visibility.Visible;
            AcceptBut.Visibility = Visibility.Hidden;
            DeleteBut.Visibility = Visibility.Hidden;
            DisAcceptBut.Visibility = Visibility.Hidden;

            AddBut.Visibility = Visibility.Hidden;
            Cleaning2();
        }

        private void Button_Click_18(object sender, RoutedEventArgs e)
        {
            SelAdd.Foreground = Brushes.White;
            SelEdit.Foreground = Brushes.White;
            SelDelete.Foreground = Brushes.DarkOrange;
            DropShadowEffect shd = new DropShadowEffect();
            shd.ShadowDepth = 0;
            shd.Opacity = 1;
            shd.BlurRadius = 6;
            shd.Color = Colors.DarkOrange;
            SelDelete.Effect = shd;

            DropShadowEffect shd1 = new DropShadowEffect();

            shd1.Opacity = 0;
            SelAdd.Effect = shd1;
            SelEdit.Effect = shd1;
            oper = "Delete";
            EditBut.Visibility = Visibility.Hidden;
            AcceptBut.Visibility = Visibility.Hidden;
            DisAcceptBut.Visibility = Visibility.Hidden;
            DeleteBut.Visibility = Visibility.Visible;

            AddBut.Visibility = Visibility.Hidden;
            Cleaning2();
        }

        private void DataGrid1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (oper == "Edit")
            {
                AcceptBut.Visibility = Visibility.Hidden;
                DisAcceptBut.Visibility = Visibility.Hidden;
                TB1.Text = "";
                TB2.Text = "";
                TB3.Text = "";
                TB4.Text = "";
               
                TBS.Text = "";
                Med1.Text = "";
                Med2.Text = "";
                Med3.Text = "";
                Vrach1.Text = "";
                Vrach2.Text = "";
                Vrach3.Text = "";
                Vrach4.Text = "";
                Vrach5.Text = "";
                Vrach6.Text = "";
                Pacient1.Text = "";
                Pacient2.Text = "";
                Pacient3.Text = "";
                Pacient4.Text = "";
                Pacient5.Text = "";
                Pacient6.Text = "";
                Pacient7.Text = "";
                DataGrid2.ItemsSource = null;
                DataGrid2.Items.Refresh();

            }
        }

        private void DataGrid2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            vr.Visibility = Visibility.Hidden;
            if (who == "Врач" || who == "Пациент")
            {
                TB2.Text = "";
            }
        }

        private void TB1_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            TB3.Text = "";
        }

        public void Cleaning()
        {
            TB1.Text = "";
            TB2.Text = "";
            TB3.Text = "";
            TB4.Text = "";
           
            TBS.Text = "";
            Med1.Text = "";
            Med2.Text = "";
            Med3.Text = "";
            Vrach1.Text = "";
            Vrach2.Text = "";
            Vrach3.Text = "";
            Vrach4.Text = "";
            Vrach5.Text = "";
            Vrach6.Text = "";
            Pacient1.Text = "";
            Pacient2.Text = "";
            Pacient3.Text = "";
            Pacient4.Text = "";
            Pacient5.Text = "";
            Pacient6.Text = "";
            Pacient7.Text = "";
            DataGrid2.ItemsSource = null;
            DataGrid2.Items.Refresh();
            Date1.Text = "";
            Date2.Text = "";
            ChVr.Text = "";
            ChPac.Text = "";
            ChSpec.Text = "";
            ChUch.Text = "";
            ChDr1.Text = "";
            ChDr2.Text = "";
            DataGrid2.ItemsSource = null;
            DataGrid2.Items.Refresh();
        }

        public void Cleaning2()
        {
            TB1.Text = "";
            TB2.Text = "";
            TB3.Text = "";
            TB4.Text = "";
    
            TBS.Text = "";
            Med1.Text = "";
            Med2.Text = "";
            Med3.Text = "";
            Vrach1.Text = "";
            Vrach2.Text = "";
            Vrach3.Text = "";
            Vrach4.Text = "";
            Vrach5.Text = "";
            Vrach6.Text = "";
            Pacient1.Text = "";
            Pacient2.Text = "";
            Pacient3.Text = "";
            Pacient4.Text = "";
            Pacient5.Text = "";
            Pacient6.Text = "";
            Pacient7.Text = "";
            DataGrid2.ItemsSource = null;
            DataGrid2.Items.Refresh();
        }
        private void DisAcceptBut_Click(object sender, RoutedEventArgs e)
        {
            Cleaning2();
            AcceptBut.Visibility = Visibility.Hidden;
            DisAcceptBut.Visibility = Visibility.Hidden;
        }

        private void DataGrid1_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            load();
            if (oper == "Plus" && whatTable == "Медсестра")
            {
               
                DropShadowEffect shd = new DropShadowEffect();
                shd.ShadowDepth = 0;
                shd.Opacity = 1;
                shd.BlurRadius = 6;
                shd.Color = Colors.DarkOrange;
                SelEdit.Effect = shd;

                DropShadowEffect shd1 = new DropShadowEffect();
                SelAdd.Foreground = Brushes.White;
                shd1.Opacity = 0;
                SelAdd.Effect = shd1;
                SelDelete.Effect = shd1;
                SelDelete.Foreground = Brushes.White;
                SelEdit.Foreground = Brushes.DarkOrange;
                oper = "Edit";
                EditBut.Visibility = Visibility.Visible;
                AcceptBut.Visibility = Visibility.Hidden;
                DeleteBut.Visibility = Visibility.Hidden;
                DisAcceptBut.Visibility = Visibility.Hidden;

                AddBut.Visibility = Visibility.Hidden;
                SelAdd.Visibility = Visibility.Hidden;
            }
            //else
            //{
            //    SelAdd.Visibility = Visibility.Visible;
            //}

            if (oper == "Edit" || oper == "Delete" && whatTable == "Медсестра")
            {
                SelAdd.Visibility = Visibility.Hidden;
            }
            if(whatTable != "Медсестра")
            {
                SelAdd.Visibility = Visibility.Visible;
            }


        }

       

        private void Vr_1630_Click(object sender, RoutedEventArgs e)
        {
            h = 16;
            m = 30;
            vremya = "16:30";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }

        private void Vr_825_Click(object sender, RoutedEventArgs e)
        {
            h = 8;
            m = 25;
            vremya = "8:25";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }
        string filtWhat ="";
        private void Button_Click_19(object sender, RoutedEventArgs e)
        {
            dt2 = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Специальность.Наименование as [Специальность],Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности and Уволен is NULL");
            DataGrid2.ItemsSource = dt2.DefaultView;
            DataGrid2.Columns[0].Visibility = Visibility.Hidden;
            filtWhat = "Врач";
        }

        private void Vr_800_Click(object sender, RoutedEventArgs e)
        {
            h = 8;
            m = 00;
            vremya = "8:00";
            TB2.Text = vremya;
            vr.Visibility = Visibility.Hidden;
        }
        
        private void ChVr_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                 
                dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");

                if (Date1.Text == "" && Date2.Text == "" && ChPac.Text == "")
                {
                    dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }
                if (Date1.Text != "" && Date2.Text != "" && ChPac.Text != "")
                {
                    dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%' AND [Врач] LIKE '%{ChPac.Text}%' AND [Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }
                if (Date1.Text == "" && Date2.Text == "" && ChPac.Text != "")
                {
                    dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%' AND [Врач] LIKE '%{ChPac.Text}%'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }
                if (Date1.Text != "" && Date2.Text != "" && ChPac.Text == "")
                {
                    dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%' AND [Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;

                }
                if (ChVr.Text =="")
                {
                    if (Date1.Text == "" && Date2.Text == "" && ChPac.Text == "")
                    {
                       
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                    }
                    if (Date1.Text != "" && Date2.Text != "" && ChPac.Text != "")
                    {
                        dt_user.DefaultView.RowFilter = string.Format($"[Врач] LIKE '%{ChPac.Text}%' AND [Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                    }
                    if (Date1.Text == "" && Date2.Text == "" && ChPac.Text != "")
                    {
                        dt_user.DefaultView.RowFilter = string.Format($"[Врач] LIKE '%{ChPac.Text}%'");
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                    }
                    if (Date1.Text != "" && Date2.Text != "" && ChPac.Text == "")
                    {
                        dt_user.DefaultView.RowFilter = string.Format($"[Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                        DataGrid1.ItemsSource = dt_user.DefaultView;

                    }
                }

                 (DataGrid1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

                (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "HH:mm";
                (DataGrid1.Columns[5] as DataGridTextColumn).Width = 130;
                (DataGrid1.Columns[7] as DataGridTextColumn).Width = 95;
                (DataGrid1.Columns[6] as DataGridTextColumn).Width = 130;
                DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                DataGrid1.Columns[1].Visibility = Visibility.Hidden;
            }
            catch
            {
                MessageBox.Show("Недопустимый знак");
                ChVr.Text = "";
            }
        }
        
        private void ChPac_TextChanged(object sender, TextChangedEventArgs e)
        {
            try { 
            dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");


            if (Date1.Text == "" && Date2.Text == "" && ChVr.Text == "")
            {
                dt_user.DefaultView.RowFilter = string.Format($"[Врач] LIKE '%{ChPac.Text}%'");
                DataGrid1.ItemsSource = dt_user.DefaultView;
            }
            if (Date1.Text != "" && Date2.Text != "" && ChVr.Text != "")
            {
                dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%' AND [Врач] LIKE '%{ChPac.Text}%' AND [Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                DataGrid1.ItemsSource = dt_user.DefaultView;
            }
            if (Date1.Text == "" && Date2.Text == "" && ChVr.Text != "")
            {
                dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%' AND [Врач] LIKE '%{ChPac.Text}%'");
                DataGrid1.ItemsSource = dt_user.DefaultView;
            }
            if (Date1.Text != "" && Date2.Text != "" && ChVr.Text == "")
            {
                dt_user.DefaultView.RowFilter = string.Format($"[Врач] LIKE '%{ChPac.Text}%' AND [Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                DataGrid1.ItemsSource = dt_user.DefaultView;
            }
            if(ChPac.Text == "")
                {
                    if (Date1.Text == "" && Date2.Text == "" && ChVr.Text == "")
                    {
                      
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                    }
                    if (Date1.Text != "" && Date2.Text != "" && ChVr.Text != "")
                    {
                        dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%' AND [Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                    }
                    if (Date1.Text == "" && Date2.Text == "" && ChVr.Text != "")
                    {
                        dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%'");
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                    }
                    if (Date1.Text != "" && Date2.Text != "" && ChVr.Text == "")
                    {
                        dt_user.DefaultView.RowFilter = string.Format($"[Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                    }
                }
                (DataGrid1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

            (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "HH:mm";
            (DataGrid1.Columns[5] as DataGridTextColumn).Width = 130;
            (DataGrid1.Columns[7] as DataGridTextColumn).Width = 95;
            (DataGrid1.Columns[6] as DataGridTextColumn).Width = 130;
            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
            DataGrid1.Columns[1].Visibility = Visibility.Hidden;
            }
            catch
            {
                MessageBox.Show("Недопустимый знак");
                ChPac.Text = "";
            }
        }

        private void Date2_TextInput(object sender, TextCompositionEventArgs e)
        {
           
        }

        private void SelAdd_Copy_Click(object sender, RoutedEventArgs e)
        {
            if (oper == "Edit")
            {
                Cleaning2();
                AcceptBut.Visibility = Visibility.Hidden;
                DisAcceptBut.Visibility = Visibility.Hidden;
            }
            FiltTalon.Visibility = Visibility.Hidden;
            FiltVrach.Visibility = Visibility.Hidden;
            FiltPac.Visibility = Visibility.Hidden;
            if (filt.Visibility == Visibility.Hidden) filt.Visibility = Visibility.Visible;
            else filt.Visibility = Visibility.Hidden;
            if (whatTable == "Талон")
            {
                dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");
                FiltTalon.Visibility = Visibility.Visible;
            }
            if (whatTable == "Врач")
            {
                dt_user = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Специальность.Наименование as [Специальность],Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности AND Уволен is NULL");
                FiltVrach.Visibility = Visibility.Visible; 
            }
            if (whatTable == "Пациент")
            {
                dt_user = Select("SELECT Пациент.Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка and Пациент.Исключен is NULL");
                FiltPac.Visibility = Visibility.Visible;
            }
            
        }

        private void ChSpec_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                dt_user = Select("SELECT Id_Врача, Врач.Фамилия, Врач.Имя, Врач.Отчество, Специальность.Наименование as [Специальность],Телефон, Кабинет FROM [dbo].Специальность, [dbo].Врач WHERE[dbo].[Специальность].Id_Специальности = [dbo].[Врач].Id_Специальности AND Уволен is NULL");

                if (ChSpec.Text != "")
                {
                    dt_user.DefaultView.RowFilter = string.Format($"[Специальность] LIKE '%{ChSpec.Text}%'");
                   
                }
                DataGrid1.ItemsSource = dt_user.DefaultView;
                DataGrid1.Columns[0].Visibility = Visibility.Hidden;
            }
            catch
            {
                MessageBox.Show("Недопустимый знак");
                ChSpec.Text = "";
            }
        }

        private void SelAdd_Copy2_Click(object sender, RoutedEventArgs e)
        {
            try { 
            dt_user = Select("SELECT Пациент.Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка and Пациент.Исключен is NULL");

            if (ChDr1.Text != "" && ChDr2.Text != "")
            {
                if (ChDr1.Text != "" && ChDr2.Text != "" && ChUch.Text != "")
                {
                    dt_user.DefaultView.RowFilter = string.Format($"[Участок] LIKE '%{ChUch.Text}%' AND [Дата рождения] >= '{ChDr1.Text}' AND [Дата рождения] <= '{ChDr2.Text}'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }
                if (ChDr1.Text != "" && ChDr2.Text != "" && ChUch.Text == "")
                {
                    dt_user.DefaultView.RowFilter = string.Format($"[Дата рождения] >= '{ChDr1.Text}' AND [Дата рождения] <= '{ChDr2.Text}'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                    DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                    (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";
                }
                DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";
            }
            }
            catch
            {
                MessageBox.Show("Недопустимый знак");
            }
        }

        private void ChUch_TextChanged(object sender, TextChangedEventArgs e)
        {
            try { 
            dt_user = Select("SELECT Пациент.Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка and Пациент.Исключен is NULL");

            if (ChDr1.Text != "" && ChDr2.Text != "" && ChUch.Text != "")
            {
                dt_user.DefaultView.RowFilter = string.Format($"[Участок] LIKE '%{ChUch.Text}%' AND [Дата рождения] >= '{ChDr1.Text}' AND [Дата рождения] <= '{ChDr2.Text}'");
                DataGrid1.ItemsSource = dt_user.DefaultView;
            }
            if (ChDr1.Text == "" && ChDr2.Text == "" && ChUch.Text !="")
            {
                dt_user.DefaultView.RowFilter = string.Format($"[Участок] LIKE '%{ChUch.Text}%'");
                DataGrid1.ItemsSource = dt_user.DefaultView;
            }

            if (ChUch.Text == "")
                {
                    if (ChDr1.Text != "" && ChDr2.Text != ""  )
                    {
                        dt_user.DefaultView.RowFilter = string.Format($"[Дата рождения] >= '{ChDr1.Text}' AND [Дата рождения] <= '{ChDr2.Text}'");
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                    }
                    if (ChDr1.Text == "" && ChDr2.Text == "" && ChUch.Text == "")
                    {
                        
                        DataGrid1.ItemsSource = dt_user.DefaultView;
                    }
                }
            DataGrid1.Columns[0].Visibility = Visibility.Hidden;
            (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";
            }
            catch
            {
                MessageBox.Show("Недопустимый знак");
                ChUch.Text = "";
            }
        }

        private void ToEx_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcelAndCsv();
        }

        private void ToEx_Copy_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int indexw = DataGrid1.SelectedIndex;
                DataRowView row = (DataRowView)DataGrid1.Items[indexw];
                string sourcePath = @"C:\Users\Diana\Desktop\Талон.docx";
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Word files(*.docx)|*.docx|All files(*.*)|*.*";
                if (saveFileDialog.ShowDialog() == true)
                {
                    string path = saveFileDialog.FileName;
                    //string path = @"C:\Users\Diana\Desktop\" + row["Пациент"] + ".docx";
                    //System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) +

                    File.Copy(sourcePath, path, true);

                    Word.Application wordApp = new Word.Application();
                    wordApp.Visible = false;

                    
                    Word.Document wordDoucment = wordApp.Documents.Open(path);
                    Word.Range range = wordDoucment.Content;
                    range = wordDoucment.Content;
                    range.Find.ClearFormatting();
                    range.Find.Execute(FindText: "{WSpec}", ReplaceWith: (string)row["Врач"]);
                    range = wordDoucment.Content;
                    range.Find.ClearFormatting();
                    string[] s = row["Дата приема"].ToString().Split(' ');
                    range.Find.Execute(FindText: "{WDate}", ReplaceWith: s[0]);
                    range = wordDoucment.Content;
                    range.Find.ClearFormatting();
                     s = row["Время приёма"].ToString().Split(' ');
                    s = s[1].Split(':');
                    string ss = s[0] + ":" + s[1];
                    range.Find.Execute(FindText: "{WTime}", ReplaceWith: ss);
                    range = wordDoucment.Content;
                    range.Find.ClearFormatting();
                    range.Find.Execute(FindText: "{WKab}", ReplaceWith: (string)row["Кабинет"].ToString());
                    range = wordDoucment.Content;
                    range.Find.ClearFormatting();
                    range.Find.Execute(FindText: "{WPac}", ReplaceWith: (string)row["Пациент"]);
                    range = wordDoucment.Content;
                    range.Find.ClearFormatting();
                   


                    wordDoucment.Save();
                    wordDoucment.Close();
                    wordApp.Quit();

                    MessageBox.Show("Файл сохранен");
                }
                else MessageBox.Show("Сохранение отменено");
            }
            catch
            {
                MessageBox.Show("Выберите талон");
            }
         }

        

        private void Rectangle_MouseDown_1(object sender, MouseButtonEventArgs e)
        {
            vt.Visibility = Visibility.Hidden;
            filt.Visibility = Visibility.Hidden;
            vr.Visibility = Visibility.Hidden;
        }

        private void ButFilt_Copy_Click(object sender, RoutedEventArgs e)
        {
            if (vt.Visibility == Visibility.Hidden) vt.Visibility = Visibility.Visible;
            else vt.Visibility = Visibility.Hidden;
        }

        private void ChVr_TextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if (inp == (char)Key.Space && ChVr.Text == "")
            {
                e.Handled = true;
            }
            else
            {
                if (!Char.IsLetter(inp))
                    e.Handled = true;
            }
        }

        private void ChPac_TextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            
            if(inp == (char)Key.Space && ChPac.Text == "")
            {
                e.Handled = true;
            }
            else
            {
                if (!Char.IsLetter(inp))
                    e.Handled = true;
            }
        }

        private void SelAdd_Copy_Click_1(object sender, RoutedEventArgs e)
        {
            Date1.Text = "";
            Date2.Text = "";
            dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");

            
                if (ChVr.Text != "" && ChPac.Text != "")
                {

                    dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%' AND [Врач] LIKE '%{ChPac.Text}%'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }
                if (ChVr.Text != "" && ChPac.Text == "")
                {

                    dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }
                if (ChVr.Text == "" && ChPac.Text != "")
                {

                    dt_user.DefaultView.RowFilter = string.Format($"[Врач] LIKE '%{ChPac.Text}%'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }
                if (ChVr.Text == "" && ChPac.Text == "")
                {

                    
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }

                (DataGrid1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

                (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "HH:mm";
                (DataGrid1.Columns[5] as DataGridTextColumn).Width = 130;
                (DataGrid1.Columns[7] as DataGridTextColumn).Width = 95;
                (DataGrid1.Columns[6] as DataGridTextColumn).Width = 130;
                DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                DataGrid1.Columns[1].Visibility = Visibility.Hidden;
            }

        private void SelAdd_Copy3_Click(object sender, RoutedEventArgs e)
        {
            ChDr1.Text = "";
            ChDr2.Text = "";
            dt_user = Select("SELECT Пациент.Id_Пациента, Пациент.Фамилия, Пациент.Имя, Пациент.Отчество, Пациент.Телефон, Пациент.[Дата рождения] as [Дата рождения], Пациент.Адрес, Участок.Наименование as [Участок] FROM [dbo].Пациент, [dbo].Участок WHERE[dbo].[Участок].Id_Участка = [dbo].[Пациент].Id_Участка and Пациент.Исключен is NULL");

            
                if (ChDr1.Text != "" && ChDr2.Text != "" && ChUch.Text != "")
                {
                    dt_user.DefaultView.RowFilter = string.Format($"[Участок] LIKE '%{ChUch.Text}%'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";
            }
                if (ChDr1.Text == "" && ChDr2.Text == "" && ChUch.Text == "")
                {
                   
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                    DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                    (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";
                }
               
            
        }

        private void ChUch_TextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];

            if (inp == (char)Key.Space && ChUch.Text == "")
            {
                e.Handled = true;
            }
            else
            {
                if (!Char.IsLetter(inp))
                    e.Handled = true;
            }
        }

        private void picTime_Click(object sender, RoutedEventArgs e)
        {
            if (TB4.Text != "")
            {

                if (vr.Visibility == Visibility.Hidden) vr.Visibility = Visibility.Visible;
                else vr.Visibility = Visibility.Hidden;
                Vremya();

            }
            else MessageBox.Show("Выберите врача");
        }

        int filtid;
        private void Button_Click_11(object sender, RoutedEventArgs e)
        {
            //ghbftn
            try
            {
               // TB2.Text = "";
                if (filtWhat == "Врач")
                {
                    index1 = DataGrid2.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid2.Items[index1];
                    filtid = (int)row["Id_Врача"];
                    ChVr.Text = row["Фамилия"].ToString();
                }
                if (who == "Пациент")
                {
                    index1 = DataGrid2.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid2.Items[index1];   /////////////////


                    TB3.Text = row["Фамилия"].ToString();
                    id1 = (int)row["Id_Пациента"];

                }
                if (who == "Врач")
                {
                    Boolean b = false; Boolean bb = false;
                    if (oper == "Edit")
                    {
                        dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,Медсестра.Фамилия as [Фамилия медсестры], Пациент.Фамилия as [Фамилия пациента], Врач.Фамилия as [Фамилия врача], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");

                        for (int i = 0; i < dt_user.Rows.Count; i++)
                        {

                            index1 = DataGrid2.SelectedIndex;
                            DataRowView row1 = (DataRowView)DataGrid2.Items[index1];   /////////////////неправильно
                            int index2 = DataGrid1.SelectedIndex;
                            dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,Медсестра.Фамилия as [Фамилия медсестры], Пациент.Фамилия as [Фамилия пациента], Врач.Фамилия as [Фамилия врача], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");
                            DataRowView rowdt = (DataRowView)DataGrid1.Items[index2];

                            DataRowView row = (DataRowView)DataGrid1.Items[i];
                            string s = TB1.Text;
                            string ss1 = row["Дата приема"].ToString();
                            int idt = (int)rowdt["Id_Талона"];
                            dt_user = Select("Select Id_Врача From Врач");
                            int id22 = (int)row1["Id_Врача"];
                            //  int id11 = (int)dt_user.Rows[index2][1];
                            dt_user = Select("Select Id_Врача, Id_Пациента, Id_Талона From Талон");
                            int IDVR = (int)dt_user.Rows[i][0];
                            int IDPC = (int)dt_user.Rows[i][1];
                            int IDTL = (int)dt_user.Rows[i][2];
                            if (ss1.IndexOf(s) != -1 && IDVR == id22 && id1 == IDPC && idt != IDTL)
                            {
                                bb = true;
                            }
                        }
                    }
                    if (oper != "Edit")
                    {
                        dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,Медсестра.Фамилия as [Фамилия медсестры], Пациент.Фамилия as [Фамилия пациента], Врач.Фамилия as [Фамилия врача], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");

                        for (int i = 0; i < dt_user.Rows.Count; i++)
                        {

                            index1 = DataGrid2.SelectedIndex;
                            DataRowView row1 = (DataRowView)DataGrid2.Items[index1];   /////////////////

                            int id22 = (int)row1["Id_Врача"];
                            // int id11 = (int)dt_user.Rows[index1][1];

                            DataRowView row = (DataRowView)DataGrid1.Items[i];
                            string s = TB1.Text;
                            string ss1 = row["Дата приема"].ToString();
                            dt_user = Select("Select Id_Врача, Id_Пациента From Талон");
                            int IDVR = (int)dt_user.Rows[i][0];
                            int IDPC = (int)dt_user.Rows[i][1];
                            if (ss1.IndexOf(s) != -1 && IDVR == id22 && id1 == IDPC)
                            {
                                b = true;
                            }
                        }
                    }
                    if (!b && !bb)
                    {
                        index1 = DataGrid2.SelectedIndex;
                        DataRowView row = (DataRowView)DataGrid2.Items[index1];   /////////////////

                        TB4.Text = row["Фамилия"].ToString();
                        id2 = (int)row["Id_Врача"];
                    }
                    else MessageBox.Show("На данную дату к этому врачу талон заказан");

                }
                //if (who == "Медсестра")
                //{
                //    index1 = DataGrid2.SelectedIndex;
                //    DataRowView row = (DataRowView)DataGrid2.Items[index1];   /////////////////

                //    TB5.Text = row["Фамилия"].ToString();
                //    id3 = (int)row["Id_Медсестры"];
                //}
                if (who == "Специальность")
                {
                    index1 = DataGrid2.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid2.Items[index1];   /////////////////

                    Vrach4.Text = row["Наименование"].ToString();
                    id1 = (int)row["Id_Специальности"];
                }
                if (who == "Участок")
                {
                    index1 = DataGrid2.SelectedIndex;
                    DataRowView row = (DataRowView)DataGrid2.Items[index1];   /////////////////

                    Pacient7.Text = row["Наименование"].ToString();
                    id1 = (int)row["Id_Участка"];
                }
                DataGrid2.ItemsSource = null;
                DataGrid2.Items.Refresh();
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }
        private void SelAdd_Copy1_Click(object sender, RoutedEventArgs e)
        {
            try { 
            dt_user = Select("SELECT Талон.Id_Талона,Пациент.Id_Пациента,(Медсестра.Фамилия+' '+Медсестра.Имя+' '+Медсестра.Отчество) as [Медсестра], (Пациент.Фамилия+' '+Пациент.Имя+' '+Пациент.Отчество) as [Пациент], (Врач.Фамилия+' '+Врач.Имя+' '+Врач.Отчество) as [Врач], Талон.[Время приёма], Талон.[Дата приема] as [Дата приема], Врач.Кабинет as [Кабинет] FROM [dbo].Талон, [dbo].Пациент, [dbo].Врач, [dbo].Медсестра WHERE[dbo].[Врач].Id_Врача = [dbo].[Талон].Id_Врача AND [dbo].[Пациент].Id_Пациента = [dbo].[Талон].Id_Пациента AND [dbo].[Медсестра].Id_Медсестры = [dbo].[Талон].Id_Медсестры ");

            if (Date1.Text !="" && Date2.Text != "")
            {
                if(ChVr.Text !="" && ChPac.Text != "")
                {
                    
                    dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%' AND [Врач] LIKE '%{ChPac.Text}%' AND [Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }
                if(ChVr.Text !="" && ChPac.Text == "")
                {
                    
                    dt_user.DefaultView.RowFilter = string.Format($"[Пациент] LIKE '%{ChVr.Text}%' AND [Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }
                if (ChVr.Text == "" && ChPac.Text != "")
                {

                    dt_user.DefaultView.RowFilter = string.Format($"[Врач] LIKE '%{ChPac.Text}%' AND [Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }
                if (ChVr.Text == "" && ChPac.Text == "")
                {

                    dt_user.DefaultView.RowFilter = string.Format($"[Дата приема] >= '{Date1.Text}' AND [Дата приема] <= '{Date2.Text}'");
                    DataGrid1.ItemsSource = dt_user.DefaultView;
                }

                (DataGrid1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy";

                (DataGrid1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "HH:mm";
                (DataGrid1.Columns[5] as DataGridTextColumn).Width = 130;
                (DataGrid1.Columns[7] as DataGridTextColumn).Width = 95;
                (DataGrid1.Columns[6] as DataGridTextColumn).Width = 130;
                DataGrid1.Columns[0].Visibility = Visibility.Hidden;
                DataGrid1.Columns[1].Visibility = Visibility.Hidden;

            }
            }
            catch
            {
                MessageBox.Show("Недопустимый знак");
            }
        }
        



    }
    public static class DataGridTextSearch
    {
        // Using a DependencyProperty as the backing store for SearchValue.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SearchValueProperty =
            DependencyProperty.RegisterAttached("SearchValue", typeof(string), typeof(DataGridTextSearch),
                new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.Inherits));

        public static string GetSearchValue(DependencyObject obj)
        {
            return (string)obj.GetValue(SearchValueProperty);
        }

        public static void SetSearchValue(DependencyObject obj, string value)
        {
            obj.SetValue(SearchValueProperty, value);
        }

        // Using a DependencyProperty as the backing store for IsTextMatch.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsTextMatchProperty =
            DependencyProperty.RegisterAttached("IsTextMatch", typeof(bool), typeof(DataGridTextSearch), new UIPropertyMetadata(false));

        public static bool GetIsTextMatch(DependencyObject obj)
        {
            return (bool)obj.GetValue(IsTextMatchProperty);
        }

        public static void SetIsTextMatch(DependencyObject obj, bool value)
        {
            obj.SetValue(IsTextMatchProperty, value);
        }
    }

    public class SearchValueConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string cellText = values[0] == null ? string.Empty : values[0].ToString();
            string searchText = values[1] as string;

            if (!string.IsNullOrEmpty(searchText) && !string.IsNullOrEmpty(cellText))
            {
                if (cellText.ToLower().IndexOf(searchText.ToLower()) != -1)
                    return true;
                else return false;
              
            }
            return false;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }




}
