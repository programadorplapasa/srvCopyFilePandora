using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dllCopyFile.Modelos
{
    public class DetObjetos
    {
    }
    public class ObjRespuestaSrv
    {
        public string Mensaje { get; set; }
        public string MensajeAdic { get; set; }
        public int CodError { get; set; }
        public int NNuSecuencia { get; set; }
        public string RutaLog { get; set; }
        public int DocEntry { get; set; }
        public string strLogEntry { get; set; }
    }

    public class Response
    {
        public string Error { get; set; }
        public bool bResult { get; set; }
    }
    public class ObjEventLog
    {
        public string Mensaje { get; set; }
        public DateTime FechaRegistro { get; set; }
    }
}
