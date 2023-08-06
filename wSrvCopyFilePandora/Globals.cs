using Admin;
using dllCopyFile.Modelos;
using Pandora;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace dllCopyFile
{
    public class Globals
    {
        private static Company oCompanyOfVta = new Company();
        private static string ServerSap = string.Empty;
        private static string CompanyDB = string.Empty;
        private static string DbServerType = string.Empty;
        private static string DbUserName = string.Empty;
        private static string DbPassword = string.Empty;
        private static string UserName = string.Empty;
        private static string Password = string.Empty;
        private static string LicenseServer = string.Empty;

        private static void GetVariablesSapOfVta()
        {
            ServerSap = ConfigurationManager.AppSettings["Server"];
            CompanyDB = ConfigurationManager.AppSettings["CompanyDB"];
            DbServerType = ConfigurationManager.AppSettings["DbServerType"];
            DbUserName = ConfigurationManager.AppSettings["DbUserName"];
            DbPassword = ConfigurationManager.AppSettings["DbPassword"];
            UserName = ConfigurationManager.AppSettings["UserName"];
            Password = ConfigurationManager.AppSettings["Password"];
            LicenseServer = ConfigurationManager.AppSettings["LicenseServer"];
        }

        public static void DisconnectOCompanyOfVta(Company oCompanyOfVta, string RutaLog)
        {

            //Log: Desconecta compañia actual
            Globals.WriteLogFile(RutaLog, "Desconecta compañia actual");

            if (oCompanyOfVta != null)
            {
                if (oCompanyOfVta.Connected)
                {
                    //Log: Desconecta compañia actual
                    Globals.WriteLogFile(RutaLog, "Desconecta compañia");
                    oCompanyOfVta.Disconnect();
                    oCompanyOfVta = null;
                }
            }

        }
        public static string ConnectSap(string RutaLog, string CCiDetDefinicion)
        {
            string errorcdb = string.Empty;
            string passwnc = string.Empty;
            string username = string.Empty;

            using (PandoraEntities db = new PandoraEntities())
            {

                Globals.WriteLogFile(RutaLog, "Se obtiene usuario de sap para conexion");

                var listUsers = (from cabDef in db.TblGeCabDefinicion.AsNoTracking()
                                 join detDef in db.TblGeDetDefinicion.AsNoTracking() on cabDef.CCiDefinicion equals detDef.CCiDefinicion
                                 where cabDef.CCiDefinicion == "USERS_OFVTA"
                                    && detDef.CCiDetDefinicion == CCiDetDefinicion
                                    && cabDef.CCeDefinicon == "A"
                                    && detDef.CCeDetDefinicion == "A"
                                 select new
                                 {
                                     detDef.CTxValor
                                 }
                                ).ToList();

                db.Dispose();

                if (listUsers.Count() == 0)
                {
                    errorcdb = "Error, configuración de usuario SAP";
                    Globals.WriteLogFile(RutaLog, "No se encuentra configurado uusarios de registro de oferta de SAP");
                    return errorcdb;
                }

                if (listUsers.Count() > 0)
                {
                    username = listUsers.FirstOrDefault().CTxValor;

                }

            }

            Globals.WriteLogFile(RutaLog, "Se obtiene clave de usuario");

            using (AdminEntities dbAdmin = new AdminEntities())
            {
                var Usuario = (from usr in dbAdmin.tblAdUsuario
                               where usr.CCiUserSAP.Contains(username)
                               && usr.CCeUsuario == "A"
                               select usr).ToList().FirstOrDefault();

                if (Usuario == null)
                {
                    Globals.WriteLogFile(RutaLog, "Usuario SAP no configurado ");
                }

                if (Usuario.CTxClaveSAP == "")
                {
                    Globals.WriteLogFile(RutaLog, "El usuario: " + Usuario.CCiUsuario + ", no posee clave de SAP no registrada");
                }

                passwnc = Usuario.CTxClaveSAP;
                dbAdmin.Dispose();

            }

            if (oCompanyOfVta != null)
            {
                if (!oCompanyOfVta.Connected)
                {
                    Globals.WriteLogFile(RutaLog, "Valida inicio de conexión");
                    errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);

                    if (errorcdb != "")
                    {
                        DisconnectOCompanyOfVta(oCompanyOfVta, RutaLog);
                        Globals.WriteLogFile(RutaLog, "Valida inicio de conexión");
                        errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);
                    }

                }

                else
                {
                    DisconnectOCompanyOfVta(oCompanyOfVta, RutaLog);
                    Globals.WriteLogFile(RutaLog, "Valida inicio de conexión");
                    errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);
                }

            }
            else
            {
                Globals.WriteLogFile(RutaLog, "Valida inicio de conexión");
                errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);
            }

            return errorcdb;
        }
        public static string ConnectSap(string RutaLog, string CCiDetDefinicion, Company oCompanyOfVta)
        {
            string errorcdb = string.Empty;
            string passwnc = string.Empty;
            string username = string.Empty;

            using (PandoraEntities db = new PandoraEntities())
            {

                Globals.WriteLogFile(RutaLog, "Se obtiene usuario de sap para conexion");

                var listUsers = (from cabDef in db.TblGeCabDefinicion.AsNoTracking()
                                 join detDef in db.TblGeDetDefinicion.AsNoTracking() on cabDef.CCiDefinicion equals detDef.CCiDefinicion
                                 where cabDef.CCiDefinicion == "USERS_OFVTA"
                                    && detDef.CCiDetDefinicion == CCiDetDefinicion
                                    && cabDef.CCeDefinicon == "A"
                                    && detDef.CCeDetDefinicion == "A"
                                 select new
                                 {
                                     detDef.CTxValor
                                 }
                                ).ToList();

                db.Dispose();

                if (listUsers.Count() == 0)
                {
                    errorcdb = "Error, configuración de usuario SAP";
                    Globals.WriteLogFile(RutaLog, "No se encuentra configurado uusarios de registro de oferta de SAP");
                    return errorcdb;
                }

                if (listUsers.Count() > 0)
                {
                    username = listUsers.FirstOrDefault().CTxValor;

                }

            }

            Globals.WriteLogFile(RutaLog, "Se obtiene clave de usuario");

            using (AdminEntities dbAdmin = new AdminEntities())
            {
                var Usuario = (from usr in dbAdmin.tblAdUsuario
                               where usr.CCiUserSAP.Contains(username)
                               && usr.CCeUsuario == "A"
                               select usr).ToList().FirstOrDefault();

                if (Usuario == null)
                {
                    Globals.WriteLogFile(RutaLog, "Usuario SAP no configurado ");
                }

                if (Usuario.CTxClaveSAP == "")
                {
                    Globals.WriteLogFile(RutaLog, "El usuario: " + Usuario.CCiUsuario + ", no posee clave de SAP no registrada");
                }

                passwnc = Usuario.CTxClaveSAP;
                dbAdmin.Dispose();

            }

            if (oCompanyOfVta != null)
            {
                if (!oCompanyOfVta.Connected)
                {
                    Globals.WriteLogFile(RutaLog, "Valida inicio de conexión");
                    errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);

                    if (errorcdb != "")
                    {
                        DisconnectOCompanyOfVta(oCompanyOfVta, RutaLog);
                        Globals.WriteLogFile(RutaLog, "Valida inicio de conexión");
                        errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);
                    }

                }

                else
                {
                    DisconnectOCompanyOfVta(oCompanyOfVta, RutaLog);
                    Globals.WriteLogFile(RutaLog, "Valida inicio de conexión");
                    errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);
                }

            }
            else
            {
                Globals.WriteLogFile(RutaLog, "Valida inicio de conexión");
                errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);
            }

            return errorcdb;
        }
        public static string ConnectSapOfVta(string username, string password, Company oCompanyOfVta)
        {
            string error = string.Empty;
            GetVariablesSapOfVta();

            try
            {

                if (oCompanyOfVta == null)
                {
                    oCompanyOfVta = new Company();
                }

                //otherCompany.Server = "10.72.20.211:30015";
                oCompanyOfVta.Server = ServerSap;
                oCompanyOfVta.CompanyDB = CompanyDB;
                oCompanyOfVta.UserName = username;
                oCompanyOfVta.Password = password;
                oCompanyOfVta.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                //otherCompany.DbUserName = "SYSTEM";
                oCompanyOfVta.DbUserName = DbUserName;
                //otherCompany.DbPassword = "P1ApasA181409";
                oCompanyOfVta.DbPassword = DbPassword;
                oCompanyOfVta.UseTrusted = false;
                oCompanyOfVta.language = SAPbobsCOM.BoSuppLangs.ln_English;

                //int iErr;
                //int lRetCode2 = otherCompany.Connect();

                int lRetCode2 = 0;

                if (!oCompanyOfVta.Connected)
                {
                    lRetCode2 = oCompanyOfVta.Connect();
                }

                if (lRetCode2 != 0)
                {
                    oCompanyOfVta.GetLastError(out lRetCode2, out error);

                }

            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            return error;
        }


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

        public static string GetRutaDestino(string CCiDetDefincion, ObjRespuestaSrv ObjRespuestaSrv)
        {
            ZPLA_DETA_DEFINICIONES DetDefiniones = new ZPLA_DETA_DEFINICIONES();
            string fileGoogle = string.Empty;

            using (PandoraEntities pandora = new PandoraEntities())
            {
                var detDefiniciones = (from deta in pandora.ZPLA_DETA_DEFINICIONES
                                       where deta.CCiDetDefinicion == CCiDetDefincion // "PATCH_ANEXOS"
                                       select deta).ToList();

                if (detDefiniciones == null || detDefiniciones.ToList().Count == 0)
                {
                    Globals.WriteLogFile(ObjRespuestaSrv.MensajeAdic, "Ruta de Anexos, no se encuentra configurada. Favor revisar con Sistemas");
                    return fileGoogle;
                }

                fileGoogle = detDefiniciones.FirstOrDefault().CTxValor.ToString();
                Globals.WriteLogFile(ObjRespuestaSrv.MensajeAdic, "Recupero ruta de Destino: " + fileGoogle);
                Console.WriteLine("Recupero ruta de Destino: " + fileGoogle);

                pandora.Dispose();
            }

            return fileGoogle;

        }


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


      

        public static List<AnexoCopido> CopyFilesList(List<AnexoModel> listaAnexo, string PathFileDrive, ObjRespuestaSrv objRespCreateTxt)
        {
            AnexoCopido anexo = new AnexoCopido();
            List<AnexoCopido> list = new List<AnexoCopido>();
            ObjRespuestaSrv objRespuesta = new ObjRespuestaSrv();

            string routeAnexos = string.Empty;
            string routeAnexoDest = string.Empty;
            string FileName = string.Empty;

            try
            {
                var group = (from deta in listaAnexo
                             select deta).GroupBy(x => new { x.NNuIdReferncia, x.CCiAplicacion, x.DocEntry })
                             .Select(deta => new AnexoCopido
                             {
                                 NNuIdReferncia = deta.Key.NNuIdReferncia,
                                 CCiAplicacion = deta.Key.CCiAplicacion,
                                 DocEntry = deta.Key.DocEntry
                             });


                foreach (var item in group)
                {
                    int NNuIdReferencia = item.NNuIdReferncia;
                    string CCiAplicacion = item.CCiAplicacion;
                    string sResult = string.Empty;

                    WriteLogFile(objRespCreateTxt.MensajeAdic, "=========================================================================================");

                    var DetListAnexos = (from deta in listaAnexo
                                         where deta.NNuIdReferncia == NNuIdReferencia
                                         && deta.CCiAplicacion == CCiAplicacion
                                         select deta).ToList();

                    sResult = string.Concat(sResult, "Referencia Oferta Vta Pandora ", item.NNuIdReferncia, "; DocEntrySAP: " + item.DocEntry.ToString());
                    WriteLogFile(objRespCreateTxt.MensajeAdic, sResult);
                    WriteLogFile(objRespCreateTxt.MensajeAdic, "=========================================================================================");

                    foreach (var deta in DetListAnexos)
                    {
                        routeAnexos = Path.Combine(deta.CTxRutaOrigen, deta.CTxFolderOrigen, deta.CTxNombreArchivo);
                        string route = Path.Combine(PathFileDrive, deta.CTxNombreArchivo);
                        string CTxExtension = System.IO.Path.GetExtension(route).Substring(1);

                        routeAnexoDest = route;
                        string fileName = string.Concat(deta.CTxNombreArchivo, ".", deta.CTxExtension.Trim());
                        

                        sResult = string.Concat("Se procesa el archivo: ", deta.CTxNombreArchivo);
                        WriteLogFile(objRespCreateTxt.MensajeAdic, sResult);

                        if (!System.IO.File.Exists(routeAnexoDest))
                        {
                            WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo no existe, se copia archivo.");
                            System.IO.File.Copy(routeAnexos, route);
                        }
                        else
                        {
                            WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo existe. Se copia pero con parametro overWrite: True");
                            System.IO.File.Copy(routeAnexos, routeAnexoDest, true);
                        }

                        WriteLogFile(objRespCreateTxt.MensajeAdic, "Confirmo que el archivo existe, para pasarlo una lista de archivos copiados");
                        if (System.IO.File.Exists(routeAnexoDest))
                        {
                            anexo = new AnexoCopido();
                            anexo.NNuIdReferncia = NNuIdReferencia;
                            anexo.CCiAplicacion = CCiAplicacion;
                            anexo.RutaOrigen = routeAnexos;
                            anexo.RutaDestino = routeAnexoDest;
                            anexo.DocEntry = deta.DocEntry;
                            list.Add(anexo);
                        }
                    }

                    WriteLogFile(objRespCreateTxt.MensajeAdic, "=========================================================================================");
                }

                WriteLogFile(objRespCreateTxt.MensajeAdic, "Termina proceso de copy");
                return list;

            }
            catch (Exception ex)
            {
                WriteLogFile(objRespCreateTxt.MensajeAdic, "error - CopyFilesList: " + ex.Message);
                list = new List<AnexoCopido>();
                return list;
            }

        }


        public static ObjRespuestaSrv GenerarArchivosAdjuntoSap(List<AnexoCopido> _tblGeAnexos, Company oCompanyOfVta, Attachments2 poOfAttachment, ObjRespuestaSrv objRespCreateTxt)
        {
            ObjRespuestaSrv objRespTrans = new ObjRespuestaSrv();
            string respuesta = string.Empty;

            try
            {
                int iRow = 0;
                foreach (var deta in _tblGeAnexos)
                {

                    string CTxRutaOrigen = string.Empty;
                    string CTxRutaDestino = string.Empty;
                    string CTxNombreArchivo = string.Empty;
                    // Path.GetDirectoryName(deta.RutaDestino); 

                    CTxRutaOrigen = Path.GetDirectoryName(deta.RutaOrigen);
                    CTxRutaDestino = Path.GetDirectoryName(deta.RutaDestino);
                    CTxNombreArchivo = Path.GetFileName(deta.RutaDestino);

                    string FileName = Path.Combine(CTxRutaDestino, CTxNombreArchivo);

                    poOfAttachment.Lines.Add();
                    poOfAttachment.Lines.FileName = System.IO.Path.GetFileNameWithoutExtension(FileName);
                    poOfAttachment.Lines.FileExtension = System.IO.Path.GetExtension(FileName).Substring(1);  // file_extesion.Trim();
                    poOfAttachment.Lines.SourcePath = System.IO.Path.GetDirectoryName(FileName);
                    poOfAttachment.Lines.Override = BoYesNoEnum.tYES;

                    iRow++;
                }

                int error = 0;
                error = poOfAttachment.Add();


                if (error != 0)
                {
                    oCompanyOfVta.GetLastError(out error, out respuesta);

                    //Log: Transacción incorrecta (Grabar código y mensaje)
                    //Globals.WriteLogFile(RutaLog, "Error al grabar oferta en Sap: " + respuesta);
                    objRespTrans.Mensaje = "Error al grabar oferta en Sap: " + respuesta;
                    objRespTrans.CodError = -5;
                }
                else
                {
                    //Log: Transacción grabada correctamente
                    oCompanyOfVta.GetLastError(out error, out respuesta);
                    //Globals.WriteLogFile(RutaLog, "Oferta grabada correctamente");
                    objRespTrans.DocEntry = int.Parse(oCompanyOfVta.GetNewObjectKey());
                    objRespTrans.CodError = 1;
                }

                return objRespTrans;
            }
            catch (Exception ex)
            {
                objRespTrans.Mensaje = "Error al grabar Archivos Adjuntos en Sap: " + ex.Message;
                return objRespTrans;

            }
        }


        public static int GetAttachmentEntry(List<AnexoCopido> listAnexos, long NNuIdReferencia, string CTpOrigen, string CTxRutaOrigen, ObjRespuestaSrv objRespCreateTxt, Company oCompanyOfVta)
        {
            int AttachmentEntry = 0;


            try
            {
                List<AnexoCopido> _Anexos = new List<AnexoCopido>();
                string CTxRutaDestino = string.Empty;
                Attachments2 poOfAttachment = oCompanyOfVta.GetBusinessObject(BoObjectTypes.oAttachments2);

                var objRespOfVtaAnexos = GenerarArchivosAdjuntoSap(listAnexos, oCompanyOfVta, poOfAttachment, objRespCreateTxt);
                if (objRespOfVtaAnexos.CodError == 1)
                {
                    //devuelve el doc entry AbsEntry del documento ,  mediante le objeto respuesta
                    AttachmentEntry = objRespOfVtaAnexos.DocEntry;
                    Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo adjuntos generado con exito en sap " + objRespOfVtaAnexos.DocEntry);
                }

            }
            catch (Exception ex)
            {
                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo adjuntos generado con exito en sap:  " + ex.Message);
            }
            return AttachmentEntry;
        }
        public  static DataSet GetDataSetAnexosPend()
        {
            string sRuta = string.Empty;
            DataSet dtsConsulta = new DataSet();
            DataTable table = new DataTable();
            string strConnSQL = string.Empty;
            string sQuery = string.Empty;

            try
            {
                // strConnSQL = ConfigurationManager.ConnectionStrings["Hermes"].ConnectionString;
                strConnSQL = ConfigurationManager.ConnectionStrings["Pandora"].ConnectionString;


                sQuery = string.Empty;
                sQuery = string.Concat(sQuery, " Select Top 1 * ", Environment.NewLine);
                sQuery = string.Concat(sQuery, " from viewDetAnexosOfertaVtaPend where 1=1   ", Environment.NewLine);

                using (SqlCommand cmd = new SqlCommand(sQuery, new SqlConnection(strConnSQL)))
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (!reader.IsClosed)
                        {
                            DataTable dt = new DataTable();
                            dt.Load(reader);
                            dtsConsulta.Tables.Add(dt);
                        }
                    }
                    cmd.Connection.Close();
                    cmd.Connection.Dispose();
                }

                return dtsConsulta;
            }
            catch (Exception)
            {
                dtsConsulta = new DataSet();
                return dtsConsulta;
            }


        }
        public static string UpdateOfertaVtaSAP(int NNuIdReferencia, int DocEntrOfertaVta, int AttachmentEntry, Company poOfVta, ObjRespuestaSrv objRespCreateTxt)
        {
            int errorUpdate = 0;
            string respuesta = string.Empty;

            try
            {
                Documents oDoc = oCompanyOfVta.GetBusinessObject(BoObjectTypes.oQuotations);
                if (oDoc.GetByKey(DocEntrOfertaVta))
                {
                    oDoc.AttachmentEntry = AttachmentEntry;
                    errorUpdate = oDoc.Update();
                    if (errorUpdate != 0)
                    {
                        oCompanyOfVta.GetLastError(out errorUpdate, out respuesta);
                        //throw new Exception("Something Wrong");
                        WriteLogFile(objRespCreateTxt.MensajeAdic, "Error al cancelar la oferta" + respuesta);
                    }
                    else
                    {
                        WriteLogFile(objRespCreateTxt.MensajeAdic, "Se actuliza correctamente en SAP");
                    }

                }
            }
            catch (Exception ex)
            {
                respuesta = ex.Message;
                WriteLogFile(objRespCreateTxt.MensajeAdic, "Error:" + respuesta);
            }
            return respuesta;
        }

        public static DataSet UpdateOfertaVta(int NNuIdReferencia, int DocEntrOfertaVta, int AttachmentEntry, ObjRespuestaSrv objRespCreateTxt)
        {
            string respuesta = string.Empty;
            string sQuery = string.Empty;
            DataSet dtsConsulta = new DataSet();

            try
            {
                string ConnectionStringSQL = ConfigurationManager.ConnectionStrings["Pandora"].ConnectionString;
                SqlConnection conn = new SqlConnection(ConnectionStringSQL);
                conn.Open();

                sQuery = string.Empty;
                sQuery = string.Concat("Exec sppvtaDetActualizanexo", Environment.NewLine);
                sQuery = string.Concat(sQuery, " @NNuIdReferencia =", NNuIdReferencia, Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @AtcEntry =", AttachmentEntry, Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @CTpOrigen ='OF_VENTA' ", Environment.NewLine);

                using (SqlCommand cmd = new SqlCommand(sQuery, new SqlConnection(ConnectionStringSQL)))
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (!reader.IsClosed)
                        {
                            DataTable dt = new DataTable();
                            // DataTable.Load automatically advances the reader to the next result set
                            dt.Load(reader);
                            dtsConsulta.Tables.Add(dt);
                        }
                    }
                    cmd.Connection.Close();
                    cmd.Connection.Dispose();
                }

                conn.Dispose();

                return dtsConsulta;
            }
            catch (Exception ex)
            {

                dtsConsulta = new DataSet();
                return dtsConsulta;


            }
            return dtsConsulta;

        }


    }
}
