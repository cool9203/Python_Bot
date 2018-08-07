using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using MySql.Data.MySqlClient;

namespace ToPython
{
    [Guid("05064E38-5E93-4659-8137-75CB4D46EB62")]

    public interface TOPYTHON
    {
        void SetConnectionString(string str);
        string GetConnectionString();
        Boolean Connect();
        Boolean Command(string commandString);
        int Select(string selectString);
    }


    [ClassInterface(ClassInterfaceType.None)]
    [Guid("880997A4-343B-4904-A84B-A3F4F02B2450")]
    [ProgId("ToPython.Application")]
    public class ToPython : TOPYTHON
    {
        DataTable dataTable = new DataTable();
        string connectionString = null;

        public void SetConnectionString(string connectionString)
        {
            this.connectionString = connectionString;
        }


        public string GetConnectionString()
        {
            return connectionString;
        }


        public Boolean Connect()
        {
            try
            {
                MySqlConnection connection = new MySqlConnection();
                connection.ConnectionString = connectionString;
                connection.Open();
                connection.Close();
                connection.Dispose();
            }
            catch (Exception e)
            {
                SaveErrorText(e);
                return false;
            }
            return true;
        }

        public Boolean Command(string commandString)
        {
            MySqlConnection connection = new MySqlConnection();
            MySqlCommand command = new MySqlCommand();
            MySqlDataAdapter dataAdapter = new MySqlDataAdapter();
            try
            {
                connection.ConnectionString = connectionString;
                connection.Open();
                command.Connection = connection;
                command.CommandText = commandString;
                dataAdapter.SelectCommand = command;
                dataAdapter.Fill(dataTable);

                connection.Close();
                connection.Dispose();
                command.Dispose();
                dataAdapter.Dispose();
            }
            catch (Exception e)
            {
                SaveErrorText(e);
                return false;
            }
            return true;
        }


        public int Select(string selectString)
        {
            int count = -2;
            try
            {
                DataRow[] temp = dataTable.Select(selectString);
                count = temp.Count();
            }
            catch (Exception e)
            {
                SaveErrorText(e);
                return -1;
            }
            return count;
        }


        void SaveErrorText(Exception e)
        {
            using (FileStream fs = new FileStream(@"C:\Users\s0011\Desktop\DllErrormsg.txt", FileMode.Append, FileAccess.Write))
            {
                fs.Write(Encoding.UTF8.GetBytes(DateTime.Today.ToLongTimeString().ToString()), 0, Encoding.UTF8.GetByteCount(DateTime.Today.ToLongTimeString().ToString()));
                fs.Write(Encoding.UTF8.GetBytes(e.ToString()), 0, Encoding.UTF8.GetByteCount(e.ToString()));
                fs.Flush();
            }
        }
    }
}
