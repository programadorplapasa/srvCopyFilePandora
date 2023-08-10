using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wSrvCopyFilePandora.Modelos
{
    public class ObjRespuesta
    {
        public string Mensaje { get; set; }
        public string MensajeAdic { get; set; }
        public int CodError { get; set; }
        public int CodError2 { get; set; }
        public string RutaLog { get; set; }
    }
    public class objRespCopyFile
    {
        public string Mensaje { get; set; }
        public bool bCopyFile { get; set; }
        public int DocEntry { get; set; }
        public int DocNum { get; set; }
    }
}
