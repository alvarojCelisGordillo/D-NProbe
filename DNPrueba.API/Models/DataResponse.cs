using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Threading.Tasks;

namespace DNPrueba.API.Models
{
    public class DataResponse
    {
        public int Code { get; set; }
        public string Message { get; set; }
        public string Base64 { get; set; }
    }
}
