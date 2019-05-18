using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using Microsoft.Win32;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Windows.Forms;

namespace PasswordReset
{
    public partial class Form1 : Form
    {
        string cs = "";

        public Form1()
        {
            InitializeComponent();
            string opendentalmsipath;
            String opendentalpath = "";

            bool is64 = System.Environment.Is64BitOperatingSystem;
            if (is64)
            {
                opendentalmsipath = "Software\\Wow6432Node\\MakeMSI\\PersistingProperties\\OpenDental";
            }
            else opendentalmsipath = "Software\\MakeMSI\\PersistingProperties\\OpenDental";

            RegistryKey key = Registry.LocalMachine.OpenSubKey(opendentalmsipath);
            if (key != null)
            {
                Object o = key.GetValue("INSTALLDIR");
                if (o != null)
                {
                    opendentalpath = o.ToString();
                }
            }

            if (File.Exists(opendentalpath + "FreeDentalConfig.xml"))
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(opendentalpath + "FreeDentalConfig.xml");
                XmlNode node = doc.DocumentElement.SelectSingleNode("/ConnectionSettings/DatabaseConnection/ComputerName");
                string servername = node.InnerText;
                XmlNode node2 = doc.DocumentElement.SelectSingleNode("/ConnectionSettings/DatabaseConnection/Database");
                string database = node2.InnerText;
                XmlNode node3 = doc.DocumentElement.SelectSingleNode("/ConnectionSettings/DatabaseConnection/User");
                string username = node3.InnerText;
                XmlNode node4 = doc.DocumentElement.SelectSingleNode("/ConnectionSettings/DatabaseConnection/Password");
                string plainpassword = node4.InnerText;
                XmlNode node5 = doc.DocumentElement.SelectSingleNode("/ConnectionSettings/DatabaseConnection/MySQLPassHash");
                if (node5 == null)
                {
                    //Console.Write($"server={servername};userid={username};password={plainpassword};database={database}");
                    //File.WriteAllText("credentials.txt", $"server={servername};userid={username};password={plainpassword};database={database}");
                    cs = $"server={servername};userid={username};password={plainpassword};database={database}";
                }
                else
                {
                    string hashpassword = node5.InnerText;
                    string realpassword = "";

                    var DLL = Assembly.LoadFile(@opendentalpath + "CDT.dll");
                    foreach (Type type in DLL.GetExportedTypes())
                    {
                        var c = Activator.CreateInstance(type);
                        object[] methodData = null;
                        methodData = new object[2];
                        methodData[0] = hashpassword;
                        methodData[1] = realpassword;
                        type.InvokeMember("Decrypt", BindingFlags.InvokeMethod, null, c, methodData);
                        cs = $"server={servername};userid={username};password={methodData[1]};database={database}";
                        //Console.Write($"server={servername};userid={username};password={methodData[1]};database={database}");
                        //File.WriteAllText("credentials.txt", $"server={servername};userid={username};password={methodData[1]};database={database}");
                    }

                }

            }
            
            MySqlConnection conn = null;
            string query = "SELECT UserName FROM opendental.userod;";
            try
            {
                using (conn = new MySqlConnection(cs))
                {
                    using (MySqlDataAdapter adapter = new MySqlDataAdapter(query, conn))
                    {
                        DataSet ds = new DataSet();
                        adapter.Fill(ds);
                        dataGridView1.DataSource = ds.Tables[0];
                        dataGridView1.AllowUserToAddRows = false;
                    }
                }
            }
            catch (MySqlException)
            {


            }
            finally
            {

                if (conn != null)
                {
                    conn.Close();
                }

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string username = dataGridView1.SelectedCells[0].Value.ToString();
            MySqlConnection conn = null;
            try
            {
                conn = new MySqlConnection(cs);
                conn.Open();

                string query = "UPDATE `opendental`.`userod` SET `Password`= '' WHERE `UserName`= '" + username + "';";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                MySqlDataReader MyReader;
                MyReader = cmd.ExecuteReader();


            }
            catch (MySqlException)
            {


            }
            finally
            {

                if (conn != null)
                {
                    conn.Close();
                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string username = dataGridView1.SelectedCells[0].Value.ToString();
            MySqlConnection conn = null;
            try
            {
                conn = new MySqlConnection(cs);
                conn.Open();

                string query = "UPDATE `opendental`.`userod` SET `UserGroupNum`= '1' WHERE `UserName`= '" + username + "';";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                MySqlDataReader MyReader;
                MyReader = cmd.ExecuteReader();


            }
            catch (MySqlException)
            {


            }
            finally
            {

                if (conn != null)
                {
                    conn.Close();
                }

            }
        }
    }
}
