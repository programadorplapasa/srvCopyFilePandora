
using Admin;
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
using wSrvCopyFilePandora.Modelos;

namespace wSrvCopyFilePandora.Clases
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
            GlobalsFiles.WriteLogFile(RutaLog, "Desconecta compañia actual");

            if (oCompanyOfVta != null)
            {
                if (oCompanyOfVta.Connected)
                {
                    //Log: Desconecta compañia actual
                    GlobalsFiles.WriteLogFile(RutaLog, "Desconecta compañia");
                    oCompanyOfVta.Disconnect();
                    oCompanyOfVta = null;
                }
            }

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


        public static string ConnectSap(string RutaLog, string CCiDetDefinicion, Company oCompanyOfVta)
        {
            string errorcdb = string.Empty;
            string passwnc = string.Empty;
            string username = string.Empty;

            using (PandoraEntities db = new PandoraEntities())
            {

                GlobalsFiles.WriteLogFile(RutaLog, "Se obtiene usuario de sap para conexion");

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
                    GlobalsFiles.WriteLogFile(RutaLog, "No se encuentra configurado uusarios de registro de oferta de SAP");
                    return errorcdb;
                }

                if (listUsers.Count() > 0)
                {
                    username = listUsers.FirstOrDefault().CTxValor;

                }

            }

            GlobalsFiles.WriteLogFile(RutaLog, "Se obtiene clave de usuario");

            using (AdminEntities dbAdmin = new AdminEntities())
            {
                var Usuario = (from usr in dbAdmin.tblAdUsuario
                               where usr.CCiUserSAP.Contains(username)
                               && usr.CCeUsuario == "A"
                               select usr).ToList().FirstOrDefault();

                if (Usuario == null)
                {
                    GlobalsFiles.WriteLogFile(RutaLog, "Usuario SAP no configurado ");
                }

                if (Usuario.CTxClaveSAP == "")
                {
                    GlobalsFiles.WriteLogFile(RutaLog, "El usuario: " + Usuario.CCiUsuario + ", no posee clave de SAP no registrada");
                }

                passwnc = Usuario.CTxClaveSAP;
                dbAdmin.Dispose();

            }

            if (oCompanyOfVta != null)
            {
                if (!oCompanyOfVta.Connected)
                {
                    GlobalsFiles.WriteLogFile(RutaLog, "Valida inicio de conexión");
                    errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);

                    if (errorcdb != "")
                    {
                        DisconnectOCompanyOfVta(oCompanyOfVta, RutaLog);
                        GlobalsFiles.WriteLogFile(RutaLog, "Valida inicio de conexión");
                        errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);
                    }

                }

                else
                {
                    DisconnectOCompanyOfVta(oCompanyOfVta, RutaLog);
                    GlobalsFiles.WriteLogFile(RutaLog, "Valida inicio de conexión");
                    errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);
                }

            }
            else
            {
                GlobalsFiles.WriteLogFile(RutaLog, "Valida inicio de conexión");
                errorcdb = ConnectSapOfVta(username, passwnc, oCompanyOfVta);
            }

            return errorcdb;
        }
        public static DataSet GetDataSetAnexosPend()
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
        public static DataSet GetDataSetAnexosPend(string CCiAplicacion, string CTpOrigen, string dbPandora)
        {
            string sRuta = string.Empty;
            DataSet dtsConsulta = new DataSet();
            DataTable table = new DataTable();
            string strConnSQL = string.Empty;
            string sQuery = string.Empty;

            try
            {
                strConnSQL = ConfigurationManager.ConnectionStrings["Hermes"].ConnectionString;
                // strConnSQL = ConfigurationManager.ConnectionStrings["Pandora"].ConnectionString;


                sQuery = string.Empty;
                sQuery = string.Concat(sQuery, " exec spcDetAnexosPend ", Environment.NewLine);
                sQuery = string.Concat(sQuery, " @CCiAplicacion = '", CCiAplicacion, "'", Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @CTpOrigen = '", CTpOrigen, "'", Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @PandoraDb = '", dbPandora, "'", Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @NNuIdReferencia = 0 ", Environment.NewLine);
                dtsConsulta = GlobalsSQL.GetDataSet(strConnSQL, sQuery);
                return dtsConsulta;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                dtsConsulta = new DataSet();
                return dtsConsulta;
            }


        }
        public static string GetRutaDestino(string CCiDetDefincion, ObjRespuesta ObjRespuestaSrv)
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
                    GlobalsFiles.WriteLogFile(ObjRespuestaSrv.MensajeAdic, "Ruta de Anexos, no se encuentra configurada. Favor revisar con Sistemas");
                    return fileGoogle;
                }

                fileGoogle = detDefiniciones.FirstOrDefault().CTxValor.ToString();
                pandora.Dispose();
            }

            return fileGoogle;

        }
        public static ObjRespuesta EjecutaCopy(int NNuIdReferncia, int DocEntry,  string CTpOrigen, string CCiAplicacion,  List<AnexoModel> listAnexos, ObjRespuesta objRespCreateTxt)
        {
            ObjRespuesta objRespuesta = new ObjRespuesta();

            bool bDocSAP = GetEstadoDocumentoSAP(DocEntry, CTpOrigen);
            if (bDocSAP)
            {
                var ResultUpdateCCeDocSap = UpdateOfertaVta(NNuIdReferncia, DocEntry, 0, "N", objRespCreateTxt);
               


                objRespuesta.CodError = -2;
                objRespuesta.Mensaje = "El documento se Origen, no puede actualizado, debido a que se encuentra Cerrado / Cancelado en SAP";
                return objRespuesta;
            }

            Company oCompanyOfVta = new Company();
            string strConn = Globals.ConnectSap(objRespCreateTxt.MensajeAdic, "USER_PROD", oCompanyOfVta);

            if (strConn != "")
            {
                string strResult = "Ocurrió un error al conectarse a Sap: " + strConn;
               //strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", strResult, Environment.NewLine);

                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, strResult);

                objRespuesta.CodError = -3;
                objRespuesta.Mensaje = strResult;
                // objListRespuesta.Add(objRespuesta);
                return objRespuesta;
            }


            Documents poOfVta = oCompanyOfVta.GetBusinessObject(BoObjectTypes.oQuotations);
            GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Conectado a Sap. registro en tabla de Anexos SAP");
            // strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Conectado a Sap. registro en tabla de Anexos SAP", Environment.NewLine);
            int AttachmentEntry = 0;
            GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "DocEntry:" + DocEntry + "; Id Referencia: " + NNuIdReferncia);            
            GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Recupero Ruta de Origen del Archivo");
            string CTxRutaOrigen = listAnexos.FirstOrDefault().CTxRutaOrigen.ToString();
            GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Ejecuto proceso: GetAttachmentEntry");
            AttachmentEntry = Globals.GetAttachmentEntry(listAnexos, NNuIdReferncia, "OF_VENTA", objRespCreateTxt, oCompanyOfVta);
            GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Respuesta GetAttachmentEntry: AttachmentEntry = " + AttachmentEntry);

            if (AttachmentEntry > 0)
            {
                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Actualizo AtcEntry en la Tabla de la Oferta de SAP ");
                string updateOQUT = Globals.UpdateOfertaVtaSAP(NNuIdReferncia, DocEntry, AttachmentEntry, oCompanyOfVta, objRespCreateTxt);
                // strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Se actualiza el registro de referencia en SAP ", Environment.NewLine);
                if (updateOQUT == "")
                {
                    GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Actualizo el AtcEntry en la tabla de Anexos");
                    var ResultUpdate = UpdateOfertaVta(NNuIdReferncia, DocEntry, AttachmentEntry, objRespCreateTxt);
                    if (ResultUpdate.CodError != 0)
                    {
                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Error al actualizar Anexos: " + ResultUpdate.Mensaje);
                        objRespuesta.CodError = ResultUpdate.CodError;
                        objRespuesta.Mensaje = ResultUpdate.Mensaje;
                        return objRespuesta;
                    }

                    GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Desconecto la conexión SAP ");
                    Globals.DisconnectOCompanyOfVta(oCompanyOfVta, objRespCreateTxt.MensajeAdic);
                    GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Eliminación de objeto oCompanyOfVta de memoria");
                    // strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Eliminación de objeto oCompanyOfVta de memoria", Environment.NewLine);
                    Marshal.ReleaseComObject(oCompanyOfVta);


                }

            }

            return objRespuesta;



        }
        public static DataSet GetDataSetUpdateOfertaVta(int NNuIdReferencia, int DocEntrOfertaVta, int AttachmentEntry, ObjRespuesta objRespCreateTxt)
        {
            string respuesta = string.Empty;
            string sQuery = string.Empty;
            DataSet dtsConsulta = new DataSet();

            try
            {
                string ConnectionStringSQL = ConfigurationManager.ConnectionStrings["Pandora"].ConnectionString;
                sQuery = string.Empty;
                sQuery = string.Concat("Exec sppvtaDetActualizanexo", Environment.NewLine);
                sQuery = string.Concat(sQuery, " @NNuIdReferencia =", NNuIdReferencia, Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @AtcEntry =", AttachmentEntry, Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @CTpOrigen ='OF_VENTA' ", Environment.NewLine);
                dtsConsulta = GlobalsSQL.GetDataSet(ConnectionStringSQL, sQuery);

                return dtsConsulta;
            }
            catch (Exception ex)
            {

                dtsConsulta = new DataSet();
                return dtsConsulta;


            }
        }

        public static dynamic UpdateOfertaVta(int NNuIdReferencia, int DocEntrOfertaVta, int AttachmentEntry, ObjRespuesta objRespCreateTxt)
        {
            string respuesta = string.Empty;
            string sQuery = string.Empty;
            DataSet dtsConsulta = new DataSet();
            int CodError ;
            string Mensaje;



            var objResult = new List<object>().Select(t => new
            {
                CodError = default(int),
                Mensaje = default(string),

            }).ToList();

            

            try
            {
                string ConnectionStringSQL = ConfigurationManager.ConnectionStrings["Pandora"].ConnectionString;
                sQuery = string.Empty;
                sQuery = string.Concat("Exec sppvtaDetActualizanexo", Environment.NewLine);
                sQuery = string.Concat(sQuery, " @NNuIdReferencia =", NNuIdReferencia, Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @AtcEntry =", AttachmentEntry, Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @CTpOrigen ='OF_VENTA' ", Environment.NewLine);
                dtsConsulta = GlobalsSQL.GetDataSet(ConnectionStringSQL, sQuery);

                if (dtsConsulta.Tables.Count > 0)
                {
                    int iCodError = 0;
                    string MsjError = "";

                    foreach (DataRow itemUpdate in dtsConsulta.Tables[0].Rows)
                    {
                        iCodError = Int32.Parse(itemUpdate["CodError"].ToString());
                        MsjError = itemUpdate["MsjError"].ToString();

                        objResult.Add(new
                        {
                            CodError = iCodError,
                            Mensaje = MsjError
                        });
                    }
                    
                }
            }
            catch (Exception ex)
            {
                objResult.Add(new
                {
                    CodError = -1,
                    Mensaje = "Error: " + ex.Message
                });

                //return objResult;
                return new { lista = objResult, bResult = false, strResult = "", CodError = -1, MsjError = "Error: " + ex.Message };

            }

            var result = objResult.FirstOrDefault();
            // urn objResult;
            return new { lista = objResult, bResult = true, strResult = "", CodError = result.CodError, MsjError = result.Mensaje };

        }
        public static dynamic UpdateOfertaVta(int NNuIdReferencia, int DocEntrOfertaVta, int AttachmentEntry, string CCeDocumento,  ObjRespuesta objRespCreateTxt)
        {
            string respuesta = string.Empty;
            string sQuery = string.Empty;
            DataSet dtsConsulta = new DataSet();
            int CodError;
            string Mensaje;



            var objResult = new List<object>().Select(t => new
            {
                CodError = default(int),
                Mensaje = default(string),

            }).ToList();



            try
            {
                string ConnectionStringSQL = ConfigurationManager.ConnectionStrings["Pandora"].ConnectionString;
                sQuery = string.Empty;
                sQuery = string.Concat("Exec sppvtaDetActualizanexo", Environment.NewLine);
                sQuery = string.Concat(sQuery, " @NNuIdReferencia =", NNuIdReferencia, Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @AtcEntry =", AttachmentEntry, Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @CTpOrigen ='OF_VENTA' ", Environment.NewLine);
                sQuery = string.Concat(sQuery, " , @CCeAnexoUpdate ='", CCeDocumento, "'", Environment.NewLine);
                dtsConsulta = GlobalsSQL.GetDataSet(ConnectionStringSQL, sQuery);

                if (dtsConsulta.Tables.Count > 0)
                {
                    int iCodError = 0;
                    string MsjError = "";

                    foreach (DataRow itemUpdate in dtsConsulta.Tables[0].Rows)
                    {
                        iCodError = Int32.Parse(itemUpdate["CodError"].ToString());
                        MsjError = itemUpdate["MsjError"].ToString();

                        objResult.Add(new
                        {
                            CodError = iCodError,
                            Mensaje = MsjError
                        });
                    }

                }
            }
            catch (Exception ex)
            {
                objResult.Add(new
                {
                    CodError = -1,
                    Mensaje = "Error: " + ex.Message
                });

                //return objResult;
                return new { lista = objResult, bResult = false, strResult = "", CodError = -1, MsjError = "Error: " + ex.Message };

            }

            var result = objResult.FirstOrDefault();
            // urn objResult;
            return new { lista = objResult, bResult = true, strResult = "", CodError = result.CodError, MsjError = result.Mensaje };

        }

        public static string UpdateOfertaVtaSAP(int NNuIdReferencia, int DocEntrOfertaVta, int AttachmentEntry, Company oCompanyOfVta, ObjRespuesta objRespCreateTxt)
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
                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Error al cancelar la oferta" + respuesta);
                    }
                    else
                    {
                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Se actuliza correctamente en SAP");
                    }

                }
            }
            catch (Exception ex)
            {
                respuesta = ex.Message;
                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Error:" + respuesta);
            }
            return respuesta;
        }
        public static List<AnexoModel> CopyFilesList(List<AnexoModel> listaAnexo, string PathFileDrive, ObjRespuesta objRespCreateTxt)
        {
            AnexoModel anexo = new AnexoModel();
            List<AnexoModel> list = new List<AnexoModel>();
            ObjRespuestaSrv objRespuesta = new ObjRespuestaSrv();

            string routeAnexos = string.Empty;
            string routeAnexoDest = string.Empty;
            string FileName = string.Empty;

            try
            {
                var group = (from deta in listaAnexo
                             select deta).GroupBy(x => new { x.NNuIdReferncia, x.CCiAplicacion, x.DocEntry, x.CTpOrigen })
                             .Select(deta => new AnexoModel
                             {
                                 NNuIdReferncia = deta.Key.NNuIdReferncia,
                                 CCiAplicacion = deta.Key.CCiAplicacion,
                                 DocEntry = deta.Key.DocEntry, 
                                 CTpOrigen = deta.Key.CTpOrigen
                             });


                foreach (var item in group)
                {
                    int NNuIdReferencia = item.NNuIdReferncia;
                    string CCiAplicacion = item.CCiAplicacion;
                    string CTpOrigen = item.CTpOrigen; 

                    string sResult = string.Empty;

                    GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "******************************************************************************************");

                    var DetListAnexos = (from deta in listaAnexo
                                         where deta.NNuIdReferncia == NNuIdReferencia
                                         && deta.CCiAplicacion == CCiAplicacion
                                         && deta.CTpOrigen == CTpOrigen
                                         select deta).ToList();

                    sResult = string.Concat(sResult, "Referencia Oferta Vta Pandora ", item.NNuIdReferncia, "; DocEntrySAP: " + item.DocEntry.ToString());
                    GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, sResult);
                    GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "=========================================================================================");

                    foreach (var deta in DetListAnexos)
                    {
                        routeAnexos = Path.Combine(deta.CTxRutaOrigen, deta.CTxFolderOrigen, deta.CTxNombreArchivo);
                        string route = Path.Combine(PathFileDrive, deta.CTxNombreArchivo);
                        string CTxExtension = System.IO.Path.GetExtension(route).Substring(1);

                        routeAnexoDest = route;
                        string fileName = string.Concat(deta.CTxNombreArchivo, ".", deta.CTxExtension.Trim());


                        sResult = string.Concat("Se procesa el archivo: ", deta.CTxNombreArchivo);
                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, sResult);

                        try
                        {

                            if (!System.IO.File.Exists(routeAnexos))
                            {
                                
                                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo no existe en la Ruta de origen.");
                                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Salta al siguiente");
                                continue;
                            }


                            if (!System.IO.File.Exists(routeAnexoDest))
                            {
                                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo no existe, se copia archivo.");
                                System.IO.File.Copy(routeAnexos, route);
                            }
                            else
                            {
                                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo existe. Se copia pero con parametro overWrite: True");
                                System.IO.File.Copy(routeAnexos, routeAnexoDest, true);
                            }

                            GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Confirmo que el archivo existe, para pasarlo una lista de archivos copiados");
                            if (System.IO.File.Exists(routeAnexoDest))
                            {
                                anexo = new AnexoModel();
                                anexo.NNuIdReferncia = NNuIdReferencia;
                                anexo.CCiAplicacion = CCiAplicacion;
                                anexo.CTxRutaOrigen = routeAnexos;
                                anexo.CTxRutaDestino = routeAnexoDest;
                                anexo.DocEntry = deta.DocEntry;
                                anexo.CTpOrigen = deta.CTpOrigen;
                                anexo.CTxNombreArchivo = deta.CTxNombreArchivo;
                                anexo.CTxFolderOrigen = deta.CTxFolderOrigen;
                                anexo.CTxExtension= deta.CTxExtension.Trim();

                                list.Add(anexo);
                            }


                        }
                        catch (Exception ex)
                        {
                            sResult = string.Concat("El archivo, ", deta.CTxNombreArchivo, ", no pede ser copiado: ", ex.Message);
                            GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Termina proceso de copy");
                        }



                    }

                    GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "******************************************************************************************");
                }

                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Termina proceso de copy");
                return list;

            }
            catch (Exception ex)
            {
                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "error - CopyFilesList: " + ex.Message);
                //list = new List<AnexoCopido>();
                return list;
            }

        }
        public static int GetAttachmentEntry(List<AnexoModel> listAnexos, long NNuIdReferencia, string CTpOrigen, ObjRespuesta objRespCreateTxt, Company oCompanyOfVta)
        {
            int AttachmentEntry = 0;


            try
            {
                List<AnexoModel> _Anexos = new List<AnexoModel>();
                string CTxRutaDestino = string.Empty;
                Attachments2 poOfAttachment = oCompanyOfVta.GetBusinessObject(BoObjectTypes.oAttachments2);

                var objRespOfVtaAnexos = GenerarArchivosAdjuntoSap(listAnexos, oCompanyOfVta, poOfAttachment, objRespCreateTxt);
                if (objRespOfVtaAnexos.CodError == 1)
                {
                    //devuelve el doc entry AbsEntry del documento ,  mediante le objeto respuesta
                    AttachmentEntry = objRespOfVtaAnexos.DocEntry;
                    GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo adjuntos generado con exito en sap " + objRespOfVtaAnexos.DocEntry);
                }

            }
            catch (Exception ex)
            {
                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo adjuntos generado con exito en sap:  " + ex.Message);
            }
            return AttachmentEntry;
        }

        public static ObjRespuestaSrv GenerarArchivosAdjuntoSap(List<AnexoModel> _tblGeAnexos, Company oCompanyOfVta, Attachments2 poOfAttachment, ObjRespuesta objRespCreateTxt)
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

                    CTxRutaOrigen = Path.GetDirectoryName(deta.CTxRutaOrigen);
                    CTxRutaDestino = Path.GetDirectoryName(deta.CTxRutaDestino);
                    CTxNombreArchivo = Path.GetFileName(deta.CTxNombreArchivo);

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


        public static bool GetEstadoDocumentoSAP(int DocEntry, string CTpOrigen)
        {
            bool bDocSAP = false;
            try
            {
                string sQuery = string.Empty;
                string ConnectionStrinHANA = ConfigurationManager.ConnectionStrings["Hana"].ConnectionString;
                string CompanyDB = ConfigurationManager.AppSettings["CompanyDB"];
                DataSet dtsConsulta = new DataSet();

                if (CTpOrigen == "OF_VENTA")
                {
                    
                    sQuery = string.Concat(sQuery, "Select ");
                    sQuery = string.Concat(sQuery, '\u0022', "DocNum", '\u0022');
                    sQuery = string.Concat(sQuery, ", ", '\u0022', "DocEntry", '\u0022');
                    sQuery = string.Concat(sQuery, ", ", '\u0022', "DocStatus", '\u0022');
                    sQuery = string.Concat(sQuery, ", ", '\u0022', "CANCELED", '\u0022');
                    sQuery = string.Concat(sQuery, " From ", '\u0022', CompanyDB, '\u0022', ".");
                    sQuery = string.Concat(sQuery, '\u0022', "OQUT", '\u0022');
                    sQuery = string.Concat(sQuery, " where ", '\u0022', "DocEntry", '\u0022');
                    sQuery = string.Concat(sQuery, "= ", DocEntry);
                    dtsConsulta = GlobalsSQL.GetDataSetHana(ConnectionStrinHANA, sQuery);

                    foreach (DataRow data in dtsConsulta.Tables[0].Rows)
                    {
                        string CCeDocumentoSAP = data["DocStatus"].ToString();
                        string Canceled = data["CANCELED"].ToString();


                        bDocSAP = false;
                        if (Canceled == "Y")
                        {
                            bDocSAP = true;
                            
                        }
                        else {
                            if (CCeDocumentoSAP == "C") {
                                bDocSAP = true;
                            }
                        }
                        
                    }
                }

            }
            catch (Exception)
            {
                bDocSAP = true;
                return bDocSAP;
            }
            return bDocSAP;
        }

    }
}
