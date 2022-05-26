using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace pruebasExcel
{
   





    class Conexion
    {

        public static MySqlConnection conexion()
        {

            string servidor = "localhost";
            string bd = "TMPermisos";
            string usuario = "root";
            string password = "280720";

            string cadenaConexion = "DataBase=" + bd + "; Data Source=" + servidor + "; User Id =" + usuario + "; Password=" + password + "";

            try
            {
                MySqlConnection conexionBD = new MySqlConnection(cadenaConexion);

                return conexionBD;
            }
            catch (MySqlException ex)
            {

                Console.WriteLine("Error :" + ex.Message);

                return null;
            }

        }

    }
}
