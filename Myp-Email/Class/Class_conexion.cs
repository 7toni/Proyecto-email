using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace Myp_Email.Class
{
    public class Class_conexion
    {
        private string servidor = "192.232.243.185";
        private string puerto = "3306";
        private string usuario = "mypsa_app2";
        private string pass = "TL5ZU9J4H2WV";

        static public string strConexion;

        public Class_conexion()
        {
        }

        public string _conexion(string nom_bd="")
        {
            string db = "";
            if (nom_bd == "2") { db = "mypsa_recordatorios"; }
            else { db = "mypsa_bitacoramyp"; }
            

            strConexion = String.Format("server={0};" +
                "port={1};" +
                "user id={2};" +
                "password={3};" +
                "database={4}; pooling=false;" +
                "Allow Zero Datetime=false; Convert Zero Datetime=true", this.servidor, this.puerto, this.usuario, this.pass, db);

            return strConexion;
        }       

    }
}
