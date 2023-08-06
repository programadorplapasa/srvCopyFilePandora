using dllCopyFile;
using dllCopyFile.Modelos;
using Microsoft.SqlServer.Server;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace wSrvCopyFilePandora
{
    public partial class Service1 : ServiceBase
    {
        bool blBandera = false;
        string strLogEntry = string.Empty;
        private static Company oCompanyOfVta;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            slpLapso.Start();
            //strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, "Pandora => Inica Proceso para copiar Anexos", Environment.NewLine);

        }

        protected override void OnStop()
        {
            //blBandera = false;
            slpLapso.Stop();
            // strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, "Pandora => Finaliza proceso", Environment.NewLine);
        }

    
        private void slpLapso2_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            List<ObjRespuestaSrv> _objRespuestas = new List<ObjRespuestaSrv>();
            ObjRespuestaSrv objRespCreateTxt = new ObjRespuestaSrv();
            LogFileModel objLogFile = new LogFileModel();
            string fileGoogle = string.Empty;
            if (blBandera) return;

            try
            {
                blBandera = true;

                objLogFile.TipoTrans = "SRV_COPY_FILE";
                objLogFile.LogDefin = "RUTA_LOG_SRV_COPYFILE";
                objRespCreateTxt = Globals.CreateLogFile(objLogFile);

                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " PLAPASA => Plasticos Panamericanos. Servicio de Copia de Archivos ", Environment.NewLine);
                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " Inica Proceso para copiar Anexos", Environment.NewLine);

                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "PLAPASA => Plasticos Panamericanos. Servicio de Copia de Archivos ");
                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Inica Proceso para copiar Anexos ");

                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Recupero Ruta de Anexos Destino ");
                string fileAnexoServer = Globals.GetRutaDestino("PATCH_ANEXOS", objRespCreateTxt);

                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "fileAnexoServer: " , fileAnexoServer, Environment.NewLine);
                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "fileAnexoServer: " + fileAnexoServer);

                if (!string.IsNullOrEmpty(fileAnexoServer))
                {

                    if (!System.IO.Directory.Exists(fileAnexoServer))
                    {
                        strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Ruta de Anexos de Unidad G:\\ no puede ser contrada ", Environment.NewLine);
                        Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Ruta de Anexos de Unidad G:\\ no puede ser contrada ");

                        EventLog.WriteEntry(strLogEntry, EventLogEntryType.Information);
                        blBandera = false;
                        return;
                    }


                    
                    // EventLog.WriteEntry(strLogEntry, EventLogEntryType.Information);

                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Inicia Servicio: SrvCopyFiles ", Environment.NewLine);
                    Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Inicia Servicio: SrvCopyFiles ");
                    var resultado = SrvCopyFiles(fileAnexoServer, objRespCreateTxt, strLogEntry);
                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Sale del Servicio: SrvCopyFiles ", Environment.NewLine);
                    var strLogError1 = resultado.ToList().LastOrDefault().strLogEntry;
                    strLogEntry = string.Empty;
                    strLogEntry = strLogError1;

                    /* pregunto si CodError = -4: No encontro registros, elimino el archivo */
                    var iConResult = resultado.Where(x => x.CodError == -4).ToList().Count();

                    if (iConResult > 0)
                    {
                        var data = Globals.DeleteLogFile(objRespCreateTxt.RutaLog);
                    }

                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Fin de Servicio ", Environment.NewLine);
                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "======================================================================== ", Environment.NewLine);

                    Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Fin de Proceso");
                    Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "========================================================================");
                }
                else
                {
                    Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Ruta de Anexos, no se encuentra configurada. Favor revisar con Sistemas");
                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Ruta de Anexos, no se encuentra configurada. Favor revisar con Sistemas", Environment.NewLine);
                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "fileAnexoServer: ", fileAnexoServer, Environment.NewLine);
                    // EventLog.WriteEntry(strLogEntry, EventLogEntryType.Information);
                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "fileAnexoServer: ", fileAnexoServer, Environment.NewLine);
                    // EventLog.WriteEntry(strLogEntry, EventLogEntryType.Information);

                    blBandera = false;
                    return;
                }

                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Fin de proceso de Copia de Anexos", Environment.NewLine);
                // EventLog.WriteEntry(strLogEntry, EventLogEntryType.Information);

                blBandera = false;

            }
            catch (Exception ex)
            {
                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Error. " + ex.Message);
                blBandera = false;
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

                    Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "=========================================================================================");

                    var DetListAnexos = (from deta in listaAnexo
                                         where deta.NNuIdReferncia == NNuIdReferencia
                                         && deta.CCiAplicacion == CCiAplicacion
                                         select deta).ToList();

                    sResult = string.Concat(sResult, "Referencia Oferta Vta Pandora ", item.NNuIdReferncia, "; DocEntrySAP: " + item.DocEntry.ToString());
                    Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, sResult);
                    Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "=========================================================================================");

                    foreach (var deta in DetListAnexos)
                    {
                        routeAnexos = Path.Combine(deta.CTxRutaOrigen, deta.CTxFolderOrigen, deta.CTxNombreArchivo);
                        string route = Path.Combine(PathFileDrive, deta.CTxNombreArchivo);
                        string CTxExtension = System.IO.Path.GetExtension(route).Substring(1);

                        routeAnexoDest = route;
                        string fileName = string.Concat(deta.CTxNombreArchivo, ".", deta.CTxExtension.Trim());


                        sResult = string.Concat("Se procesa el archivo: ", deta.CTxNombreArchivo);
                        Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, sResult);

                        try
                        {

                            if (!System.IO.File.Exists(routeAnexoDest))
                            {
                                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo no existe, se copia archivo.");
                                System.IO.File.Copy(routeAnexos, route);
                            }
                            else
                            {
                                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Archivo existe. Se copia pero con parametro overWrite: True");
                                System.IO.File.Copy(routeAnexos, routeAnexoDest, true);
                            }

                            Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Confirmo que el archivo existe, para pasarlo una lista de archivos copiados");
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
                        catch (Exception ex)
                        {
                            sResult = string.Concat("El archivo, ", deta.CTxNombreArchivo, ", no pede ser copiado: ", ex.Message);
                            Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Termina proceso de copy");
                        }



                    }

                    Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "=========================================================================================");
                }

                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Termina proceso de copy");
                return list;

            }
            catch (Exception ex)
            {
                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "error - CopyFilesList: " + ex.Message);
                //list = new List<AnexoCopido>();
                return list;
            }

        }



        public List<ObjRespuestaSrv> SrvCopyFiles(string PathFileDrive, ObjRespuestaSrv objRespCreateTxt, string strLogEntry)
        {
            AnexoModel anexoModel = new AnexoModel();
            List<AnexoModel> listAnexos = new List<AnexoModel>();
            ObjRespuestaSrv objRespuesta = new ObjRespuestaSrv();
            List<ObjRespuestaSrv> objListRespuesta = new List<ObjRespuestaSrv>();
            int iConAnexosPend = 0;
            
            try
            {
                //Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Incia Proceso de Copiar Archivos");
                //strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " Pandora => Inica Proceso para copiar Anexos", Environment.NewLine);

                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Recupero de Anexos Pendientes de Copiar.");
                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " Recupero de Anexos Pendientes de Copiar.", Environment.NewLine);
                


                DataSet dtsConsulta = Globals.GetDataSetAnexosPend();

                if (dtsConsulta.Tables.Count > 0)
                {
                    iConAnexosPend = dtsConsulta.Tables[0].Rows.Count;
                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " Existen ", iConAnexosPend, ", Registros de Anexos Pendientes " , Environment.NewLine);
                    

                    if (iConAnexosPend == 0)
                    {

                        Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "No existen registros de Anexos Pendientes.");
                        strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " No existen registros de Anexos Pendientes. ", Environment.NewLine);

                        objRespuesta.CodError = -4;
                        objRespuesta.Mensaje = "No existen registros de Anexos Pendientes.";
                        objListRespuesta.Add(objRespuesta);
                        return objListRespuesta;
                    }

                    foreach (DataRow data in dtsConsulta.Tables[0].Rows)
                    {
                        anexoModel = new AnexoModel();
                        anexoModel.CCiAplicacion = data["CCiAplicacion"].ToString();
                        anexoModel.NNuIdReferncia = Int32.Parse(data["NNuIdReferencia"].ToString());
                        anexoModel.DocEntry = Int32.Parse(data["DocEntry"].ToString());
                        anexoModel.CTxRutaOrigen = data["CTxOrigenPath"].ToString();
                        anexoModel.CTxFolderOrigen = data["CTxForlderOrigenPath"].ToString();
                        anexoModel.CTxNombreArchivo = data["CtxNombreAnexos"].ToString();
                        anexoModel.CTxExtension = data["CTpExtension"].ToString();
                        listAnexos.Add(anexoModel);
                    }


                    Company oCompanyOfVta = new Company();

                    if (listAnexos.ToList().Count > 0)
                    {
                        oCompanyOfVta = new Company();
                        string strContAnexos = listAnexos.ToList().Count().ToString();
                        var ListCopyFile = CopyFilesList(listAnexos, PathFileDrive, objRespCreateTxt);

                        if (ListCopyFile.Count > 0)
                        {
                            Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Se van a procesar " + ListCopyFile.Count + ", datos de anexos");
                            strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Se van a procesar " + ListCopyFile.Count + ", datos de anexos", Environment.NewLine);

                            Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Comienza proceso para grabar a SAP");
                            strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Comienza proceso para grabar a SAP", Environment.NewLine);

                            Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Recupero datos de funciones ConnectSAP");
                            strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Recupero datos de funciones ConnectSAP", Environment.NewLine);
                            //EventLog.WriteEntry(strLogEntry, EventLogEntryType.Information);

                            //objRespuesta.strLogEntry = strLogEntry;
                            //objListRespuesta.Add(objRespuesta);
                            //return objListRespuesta;


                            string strConn = Globals.ConnectSap(objRespCreateTxt.MensajeAdic, "USER_PROD", oCompanyOfVta);

                            if (strConn != "")
                            {
                                string strResult = "Ocurrió un error al conectarse a Sap: " + strConn;
                                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", strResult, Environment.NewLine);

                                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, strResult);

                                objRespuesta.CodError = -3;
                                objRespuesta.Mensaje = strResult;
                                objListRespuesta.Add(objRespuesta);
                                return objListRespuesta;
                            }

                            var group = (from deta in ListCopyFile
                                         select deta).GroupBy(x => new { x.NNuIdReferncia, x.CCiAplicacion, x.DocEntry })
                            .Select(deta => new AnexoCopido
                            {
                                NNuIdReferncia = deta.Key.NNuIdReferncia,
                                CCiAplicacion = deta.Key.CCiAplicacion,
                                DocEntry = deta.Key.DocEntry
                            });


                            Documents poOfVta = oCompanyOfVta.GetBusinessObject(BoObjectTypes.oQuotations);
                            Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Conectado a Sap. registro en tabla de Anexos SAP");
                            strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Conectado a Sap. registro en tabla de Anexos SAP", Environment.NewLine);

                            foreach (var item in group)
                            {

                                int AttachmentEntry = 0;

                                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", " ******************************* ", Environment.NewLine);
                                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, " ******************************* " + item.CTxNombreArchivo);


                                var fileData = ListCopyFile.Where(x => x.DocEntry == item.DocEntry && x.NNuIdReferncia == item.NNuIdReferncia).ToList();

                                // Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, " Archivo: " + item.CTxNombreArchivo);
                                // strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Archivo: " + item.CTxNombreArchivo, Environment.NewLine);

                                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "IdReferencia: " + item.NNuIdReferncia + "; DocEntry Oferta: " + item.DocEntry);
                                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "IdReferencia: " + item.NNuIdReferncia + "; DocEntry: " + item.DocEntry, Environment.NewLine);

                                var detAnexoXRef = (from deta in ListCopyFile
                                                    where deta.NNuIdReferncia == item.NNuIdReferncia
                                                    && deta.CCiAplicacion == item.CCiAplicacion
                                                    select deta).ToList();

                                AttachmentEntry = Globals.GetAttachmentEntry(detAnexoXRef, item.NNuIdReferncia, "OF_VENTA", PathFileDrive, objRespCreateTxt, oCompanyOfVta);
                                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "AttachmentEntry: " + AttachmentEntry);
                                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "AttachmentEntry: " + AttachmentEntry, Environment.NewLine);

                                if (AttachmentEntry > 0)
                                {
                                    string updateOQUT = Globals.UpdateOfertaVtaSAP(item.NNuIdReferncia, item.DocEntry, AttachmentEntry, oCompanyOfVta, objRespCreateTxt);
                                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Se actualiza el registro de referencia en SAP ", Environment.NewLine);


                                    if (updateOQUT == "")
                                    {
                                        Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, updateOQUT);

                                        strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Se actualiza el registro de referencia en Hermes / Anexos ", Environment.NewLine);
                                        DataSet dtsConsultaUpdate = Globals.UpdateOfertaVta(item.NNuIdReferncia, item.DocEntry, AttachmentEntry, objRespCreateTxt);
                                        if (dtsConsultaUpdate.Tables.Count > 0)
                                        {
                                            int iCodError = 0;
                                            string MsjError = "";

                                            foreach (DataRow itemUpdate in dtsConsultaUpdate.Tables[0].Rows)
                                            {
                                                iCodError = Int32.Parse(itemUpdate["CodError"].ToString());
                                                MsjError = itemUpdate["MsjError"].ToString();
                                            }
                                            if (iCodError == 0)
                                            {
                                                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Datos de referencia AtcEntry, se han actualizado correctamente en Hermes.");
                                                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Datos de referencia AtcEntry, se han actualizado correctamente en Hermes.", Environment.NewLine);
                                            }

                                        }


                                    }

                                    objRespuesta.CodError = 0;
                                    objRespuesta.Mensaje = "Datos de referencia AtcEntry, se han actualizado correctamente en Hermes.";
                                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Datos de referencia AtcEntry, se han actualizado correctamente en Hermes.", Environment.NewLine);
                                    objRespuesta.strLogEntry = strLogEntry;
                                }
                                else
                                {
                                    Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Error al generar el Adjunto en SAP");
                                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Error al generar el Adjunto en SAP", Environment.NewLine);
                                    objRespuesta.strLogEntry = strLogEntry;

                                }

                            }

                            Globals.DisconnectOCompanyOfVta(oCompanyOfVta, objRespCreateTxt.MensajeAdic);
                            Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Eliminación de objeto oCompanyOfVta de memoria");
                            strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Eliminación de objeto oCompanyOfVta de memoria", Environment.NewLine);
                            objRespuesta.strLogEntry = strLogEntry;

                            Marshal.ReleaseComObject(oCompanyOfVta);

                        }
                    }
                    else
                    {
                        objRespuesta.CodError = -2;
                        objRespuesta.Mensaje = "No existsn Anexos pendientes de copiar";
                        strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "CodError: -2; No existsn Anexos pendientes de copiar", Environment.NewLine);
                        objRespuesta.strLogEntry = strLogEntry;

                    }
                }
                else
                {
                    objRespuesta.CodError = -4;
                    objRespuesta.Mensaje = "No existsn Anexos pendientes de copiar";
                    strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "CodError: -4; No existsn Anexos pendientes de copiar", Environment.NewLine);

                    objRespuesta.strLogEntry = strLogEntry;
                }

                objListRespuesta.Add(objRespuesta);

            }
            catch (Exception ex)
            {
                objRespuesta.CodError = -3;
                objRespuesta.Mensaje = "Error: " + ex.Message;
                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "CodError: -3; Error: " , ex.Message, Environment.NewLine);

                Globals.DisconnectOCompanyOfVta(oCompanyOfVta, objRespCreateTxt.MensajeAdic);
                Globals.WriteLogFile(objRespCreateTxt.MensajeAdic, "Eliminación de objeto oCompanyOfVta de memoria");
                strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "Eliminación de objeto oCompanyOfVta de memoria", Environment.NewLine);
                Marshal.ReleaseComObject(oCompanyOfVta);


                objRespuesta.strLogEntry = strLogEntry;
                // return objRespuesta;
                objListRespuesta.Add(objRespuesta);
            }

            return objListRespuesta;
        }

    }
}
