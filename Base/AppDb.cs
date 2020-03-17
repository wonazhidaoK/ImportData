using MySql.Data.MySqlClient;
using System;

namespace ImportData.Base
{
    public class AppDb : IDisposable
    {
        public MySqlConnection Connection;
        public AppDb(string connectionString)
        {
            Connection = new MySqlConnection(connectionString);
        }
        public void Dispose()
        {
            Connection.Close();
        }
    }
}
