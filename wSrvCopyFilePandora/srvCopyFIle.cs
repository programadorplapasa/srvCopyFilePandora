using Microsoft.SqlServer.Server;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using wSrvCopyFilePandora.Clases;
using wSrvCopyFilePandora.Modelos;

namespace wSrvCopyFilePandora
{
    public partial class srvCopyFIle : ServiceBase
    {
        bool blBandera = false;
        string strLogEntry = string.Empty;
        private static Company oCompanyOfVta;

        public srvCopyFIle()
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

        private void slpLapso_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (blBandera) return;
            string strEventLog = string.Empty;
            ObjRespuesta objRespCreateTxt = new ObjRespuesta();
            LogFileModel objLogFile = new LogFileModel();

            try
            {
                blBandera = true;

                string DbPandora = ConfigurationManager.AppSettings["DataBasePandora"];
                DataSet dtsConsulta = Globals.GetDataSetAnexosPend("PAND", "OF_VENTA", DbPandora);

                int iCountTable = dtsConsulta.Tables.Count;
                
                if (iCountTable > 0)
                {
                    int iCountRegister = dtsConsulta.Tables[0].Rows.Count;

                    if (iCountRegister > 0) {
                        objLogFile.TipoTrans = "SRV_COPY_FILE";
                        objLogFile.LogDefin = "RUTA_LOG_SRV_COPYFILE";
                        objRespCreateTxt = GlobalsFiles.CreateLogFile(objLogFile);

                        strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " PLAPASA => Plasticos Panamericanos. Servicio de Copia de Archivos ", Environment.NewLine);
                        strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, "Se han encontrado " + iCountRegister + " Archivos pendientes de procesar ", Environment.NewLine);
                        strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, "Se inicia el proceso para copiar archivos", Environment.NewLine);

                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "PLAPASA => Plasticos Panamericanos. Servicio de Copia de Archivos ");
                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Se han encontrado " + iCountRegister + " Archivos pendientes de procesar ");
                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Se inicia el proceso para copiar archivos");

                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Recupero Ruta de Anexos Destino ");
                        string fileAnexoServer = Globals.GetRutaDestino("PATCH_ANEXOS", objRespCreateTxt);
                        strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, " ", "fileAnexoServer: ", fileAnexoServer, Environment.NewLine);
                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "fileAnexoServer: " + fileAnexoServer);

                        if (string.IsNullOrEmpty(fileAnexoServer))
                        {
                            blBandera = false;
                            GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Ruta de Anexo Server no configurada, se termina proceso" );
                            return;
                        }


                        List<AnexoModel> listAnexos = new List<AnexoModel>();
                        AnexoModel anexoModel = new AnexoModel();

                        foreach (DataRow data in dtsConsulta.Tables[0].Rows)
                        {
                            anexoModel = new AnexoModel();
                            anexoModel.CCiAplicacion = data["CCiAplicacion"].ToString();
                            anexoModel.NNuIdReferncia = Int32.Parse(data["NNuIdReferencia"].ToString());
                            anexoModel.DocEntry = Int32.Parse(data["DocEntry"].ToString());
                            anexoModel.CTxRutaOrigen = data["CTxOrigenPath"].ToString();
                            anexoModel.CTxRutaDestino = data["CTxDestinoPath"].ToString();
                            anexoModel.CTxFolderOrigen = data["CTxForlderOrigenPath"].ToString();
                            anexoModel.CTxNombreArchivo = data["CtxNombreAnexos"].ToString();
                            anexoModel.CTxExtension = data["CTpExtension"].ToString();
                            anexoModel.CTpOrigen = data["CTpOrigen"].ToString();
                            listAnexos.Add(anexoModel);
                        }


                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Se encontraron " + listAnexos.Count + ", pendientes de Procesar");
                        var ListCopyFile = Globals.CopyFilesList(listAnexos, fileAnexoServer, objRespCreateTxt);
                        

                        if (ListCopyFile.Count == 0)
                        {
                           
                            blBandera = false;
                            GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Uno o más archivos no pudo ser copiado a la ruta de destino. Favor revisar");
                            return;

                        }

                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Se procesaran a SAP " + ListCopyFile.Count);
                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Se Agrupan por Referencias de Origen segun la lista de archivos habilitados ");

                        var groupAnexos = (from deta in ListCopyFile
                                           select deta).GroupBy(x => new { x.NNuIdReferncia, x.CCiAplicacion, x.DocEntry, x.CTpOrigen })
                        .Select(deta => new                        {
                            NNuIdReferncia = deta.Key.NNuIdReferncia,
                            CTpOrigen = deta.Key.CTpOrigen,
                            CCiAplicacion = deta.Key.CCiAplicacion,
                            DocEntry = deta.Key.DocEntry
                        });


                        GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Referencias Agrupadas para procesar");

                        List<ObjRespuestaSrv> objRespuestaSrvs = new List<ObjRespuestaSrv>();
                        ObjRespuestaSrv objRespuestaSrv = new ObjRespuestaSrv();

                        foreach (var dato in groupAnexos)
                        {

                            var resultAnexo = (from deta in listAnexos
                                          where deta.NNuIdReferncia == dato.NNuIdReferncia
                                          && deta.CCiAplicacion == dato.CCiAplicacion
                                          && deta.CTpOrigen == dato.CTpOrigen
                                          select deta).ToList();





                            GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "EjecutaCopy");
                            var result = Globals.EjecutaCopy(dato.NNuIdReferncia,dato.DocEntry,  dato.CTpOrigen, dato.CCiAplicacion, resultAnexo, objRespCreateTxt);
                            objRespuestaSrv = new ObjRespuestaSrv();
                            objRespuestaSrv.CodError = result.CodError;
                            objRespuestaSrv.DocEntry = dato.DocEntry;
                            objRespuestaSrv.NNuSecuencia = dato.NNuIdReferncia;
                            objRespuestaSrv.Mensaje = result.Mensaje;
                            objRespuestaSrvs.Add(objRespuestaSrv);
                        }

                        int iContError = objRespuestaSrvs.Where(x => x.CodError != 0).ToList().Count();

                        if (iContError != 0) { GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Termina proceso con errores de Copiado"); }
                        else { GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Termina proceso sin Errores al copiar ");  }
                        


                    }
                    else
                    {
                        strLogEntry = string.Concat(strLogEntry, System.DateTime.Now, "No existen registros", Environment.NewLine);
                        //GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "No existen registros");
                        blBandera = false;
                        return;
                    }

                }
                else
                {
                    blBandera = false;
                    return;
                }

                blBandera = false;

            }
            catch (Exception ex)
            {
                GlobalsFiles.WriteLogFile(objRespCreateTxt.MensajeAdic, "Error Servicio: " + ex.Message);
                blBandera = false;
            }
        }






    }
}
