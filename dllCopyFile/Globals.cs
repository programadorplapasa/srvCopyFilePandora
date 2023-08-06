using dllCopyFile.Modelos;
using Pandora;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dllCopyFile
{
    public class Globals
    {
        public static ObjRespuestaSrv CreateLogFile(LogFileModel objLogFile)
        {
            ObjRespuestaSrv objRes = new ObjRespuestaSrv();

            try
            {

                ZPLA_DETA_DEFINICIONES objDef = new ZPLA_DETA_DEFINICIONES();
                string path = string.Empty;
                string fileName = string.Empty;

                using (PandoraEntities db = new PandoraEntities())
                {
                    objDef = db.ZPLA_DETA_DEFINICIONES.Where(x => x.CCiDetDefinicion == objLogFile.LogDefin).FirstOrDefault();

                    if (objDef.CTxValor != "")
                    {

                        if (objLogFile.TipoTrans == "SRV_COPY_FILE")
                        {

                            path = objDef.CTxValor;
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

                    }

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


    }
}
