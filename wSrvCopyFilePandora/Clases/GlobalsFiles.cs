
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wSrvCopyFilePandora.Modelos;

namespace wSrvCopyFilePandora.Clases
{
    public class GlobalsFiles
    {
        public static bool DeleteLogFile(string LogFile)
        {
            bool bDeleted = false;

            try
            {

                if (System.IO.File.Exists(LogFile))
                {
                    System.IO.File.Delete(LogFile);
                    bDeleted = true;
                }

                return bDeleted;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return bDeleted;
            }

        }

        public static ObjRespuesta CreateLogFile(LogFileModel objLogFile)
        {
            ObjRespuesta objRes = new ObjRespuesta();

            try
            {
                string path = string.Empty;
                string fileName = string.Empty;


                DataSet dtsDefiniciones = GlobalsSQL.GetDefincion("", objLogFile.LogDefin, "Pandora");


                if (dtsDefiniciones.Tables.Count == 0)
                {
                    objRes.CodError = 0;
                    objRes.Mensaje = "No existe configurada un Ruta para el log de envio de correos de Ordenes de Venta";
                    return objRes;
                }


                var data = (from deta in dtsDefiniciones.Tables[0].AsEnumerable()
                            select new
                            {
                                path = deta.Field<string>("CTxValor")
                            }).ToList().FirstOrDefault();

                if (objLogFile.TipoTrans == "SRV_COPY_FILE")
                {
                    path = data.path;
                    string currentDate = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss.fff").Replace(":", "_").Replace("-", "_").Replace(".", "_");
                    fileName = objLogFile.TipoTrans + "__" + currentDate + ".txt";
                }
                if (objLogFile.TipoTrans == "SRV-OF")
                {
                    path = data.path;
                    string currentDate = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss.fff").Replace(":", "_").Replace("-", "_").Replace(".", "_");
                    fileName = objLogFile.TipoTrans + "__" + currentDate + ".txt";
                }

                if (objLogFile.TipoTrans == "SRV-SEND-OR")
                {
                    path = data.path;
                    string currentDate = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss.fff").Replace(":", "_").Replace("-", "_").Replace(".", "_");
                    fileName = objLogFile.TipoTrans + "__" + currentDate + ".txt";
                }

                string rutaCompleta = path + fileName;

                // Check if file already exists. If yes, delete it.     
                if (File.Exists(rutaCompleta))
                {
                    File.Delete(fileName);
                }

                // Create a new file
                using (FileStream fs = File.Create(rutaCompleta))
                {
                    fs.Close();
                    objRes.CodError = 1;
                    objRes.Mensaje = "Archivo creado correctamente";
                    objRes.MensajeAdic = rutaCompleta;
                    objRes.RutaLog = rutaCompleta;
                }

                return objRes;
            }
            catch (Exception ex)
            {
                objRes.CodError = -5;
                objRes.Mensaje = "Ocurrió un error al crear el archivo: " + ex.Message;
                return objRes;
            }

        }

        public static void WriteLogFile(string Ruta, string Msg)
        {

            try
            {
                string currentDate = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss.fff").Replace(":", "_").Replace("-", "_").Replace(".", "_");
                Msg = currentDate + " | " + Msg;

                if (File.Exists(Ruta))
                {
                    File.AppendAllText(Ruta, Msg + Environment.NewLine);
                }

            }
            catch (Exception)
            {
                Console.WriteLine("Ocurrió un error al editar el archivo de log");
            }


        }

    }
}
