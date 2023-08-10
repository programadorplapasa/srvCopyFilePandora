using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wSrvCopyFilePandora.Modelos
{
        public class LogFileModel
        {
            public string TipoTrans { get; set; }
            public long DocNumSap { get; set; }
            public string TipoMat { get; set; }
            public string LogDefin { get; set; }
            public long NNuIdDocTran { get; set; }
            public List<long> listDocsTran { get; set; }
            public List<string> listItemCodes { get; set; }
        }


    }
