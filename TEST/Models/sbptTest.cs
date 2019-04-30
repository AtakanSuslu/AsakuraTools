using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace TEST.Models
{
    
    class sbptTest
    {
        [Key]
        public int ID { get; set; }
        public string Isim { get; set; }
    }
}
