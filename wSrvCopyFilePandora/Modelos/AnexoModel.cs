using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wSrvCopyFilePandora.Modelos
{
    public class AnexoModel
    {
        public int NNuIdReferncia { get; set; }
        public int DocEntry { get; set; }
        public string CCiAplicacion { get; set; }
        public string CTxRutaOrigen { get; set; }
        public string CTxFolderOrigen { get; set; }
        public string CTxNombreArchivo { get; set; }
        public string CTxExtension { get; set; }
        public string CTxRutaDestino { get; set; }
        public string CTpOrigen { get; set; }
    }

    public class AnexoCopiado
    {
        public int NNuIdReferncia { get; set; }
        public string CCiAplicacion { get; set; }
        public string CTxRutaOrigen { get; set; }
        public string CTxNombreArchivo { get; set; }
        public string CTxExtension { get; set; }
        public string CSnCopiado { get; set; }
        public string RutaOrigen { get; set; }
        public string RutaDestino { get; set; }
        public int AtcEntry { get; set; }
        public int DocEntry { get; set; }
    }
}