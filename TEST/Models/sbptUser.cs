using Modelleyici;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEST.Models
{
    public class sbptUser
    {
        public int UserID { get; set; }
        [IgnoreColumn]
        public string UserName { get; set; }
    }
}
