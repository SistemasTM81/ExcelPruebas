
using DocumentFormat.OpenXml.Spreadsheet;
using MySql.Data.MySqlClient;
using SpreadsheetLight;
using SpreadsheetLight.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pruebasExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnGen_Click(object sender, EventArgs e)
        {



            //Empezar a usar excel
            SLDocument sl = new SLDocument();

            //Directorio para cargar imagen en el Excel
            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(@"C:\Users\Ing. Osky Lopez\Downloads\mu\logo2.png");

            //Ingresar Imagen
            byte[] ba;
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                bm.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Close();
                ba = ms.ToArray();
            }

            SLPicture pic = new SLPicture(ba, DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
            pic.SetPosition(0, 0);
            pic.ResizeInPixels(300, 60);
            sl.InsertPicture(pic);
            //Ingresar Imagen

            //Titulo de la tabla
            sl.SetCellValue("G2", "Reporte de Ususarios");

            //Estilos para titulo de la tabla
            SLStyle estiloT = sl.CreateStyle();
            estiloT.Font.FontName = "Arial";
            estiloT.Font.FontSize = 16;
            estiloT.Font.Bold = true;
            sl.SetCellStyle("G2", estiloT);
            sl.MergeWorksheetCells("G2", "I2");
            //Estilos para titulo de la tabla

            //Para saber en que celda iniciar
            int celdaCabecera = 5, celdaInicial = 5;

            //Nombre de la Hoja de Excel
            sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "PRUEBAS");

            //Encabezados
            sl.SetCellValue("H" + celdaCabecera, "IdEmpleado");
            sl.SetCellValue("I" + celdaCabecera, "Nombre");
            sl.SetCellValue("J" + celdaCabecera, "IdPuesto1");

            //Estilos de la tabla 
            SLStyle estiloCa = sl.CreateStyle();
            estiloCa.Font.FontName = "Arial";
            estiloCa.Font.FontSize = 14;
            estiloCa.Font.Bold = true;
            estiloCa.Font.FontColor = System.Drawing.Color.White;
            estiloCa.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Crimson, System.Drawing.Color.Crimson);
            sl.SetCellStyle("H" + celdaCabecera, "J" + celdaCabecera, estiloCa);
            //Estilos de la tabla 

            //Consulta de la base datos
            string sql = "SELECT idEmpleado, NombreEm, idPuesto1 from Registro";

            //Conexion a la base de datos
            MySqlConnection conexionBD = Conexion.conexion();
            conexionBD.Open();


            MySqlCommand comando = new MySqlCommand(sql, conexionBD);
            MySqlDataReader reader = comando.ExecuteReader();
            //Conexion a la base de datos

            //Arreglo para recorrer los resultados de la consulta
            while (reader.Read())
            {
                celdaCabecera++;

                sl.SetCellValue("H" + celdaCabecera, reader["idEmpleado"].ToString());
                sl.SetCellValue("I" + celdaCabecera, reader["NombreEm"].ToString());
                sl.SetCellValue("J" + celdaCabecera, reader["IdPuesto1"].ToString());
            }
            //Arreglo para recorrer los resultados de la consulta

            //Estilos Para bordes de la tabla
            SLStyle EstiloB = sl.CreateStyle();

            EstiloB.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            EstiloB.Border.LeftBorder.Color = System.Drawing.Color.Black;

            EstiloB.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            EstiloB.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            EstiloB.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;

            sl.SetCellStyle("H" + celdaInicial, "J" + celdaCabecera, EstiloB);

            //Personalizar celdas



            //Personalizar celdas

            sl.AutoFitColumn("H", "J");
            //Estilos Para bordes de la tabla

            //Directorio para Guardar el Excel
            sl.SaveAs(@"C:\Users\Ing. Osky Lopez\Desktop\Pruebas\oscar.xlsx");

            // guardar en directorio 
            /*
             Stream myStream;
             SaveFileDialog saveFileDialog1 = new SaveFileDialog();

             saveFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
             saveFileDialog1.FilterIndex = 2;
             saveFileDialog1.RestoreDirectory = true;

             if (saveFileDialog1.ShowDialog() == DialogResult.OK)
             {

                 if ((myStream = saveFileDialog1.OpenFile()) != null)
                 {

                     // Code to write the stream goes here.
                     myStream.Close();
                 }
             }
                 */


        }
    }
}