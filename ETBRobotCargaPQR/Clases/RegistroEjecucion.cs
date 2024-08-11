using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ETBRobotAsignarCasosPQR.Clases
{
    class RegistroEjecucion
    {

        public int id { get; set; }
        public string empresa { get; set; }
        public string tipo { get; set; }
        public int intentos { get; set; }
        public DateTime fechaproceso { get; set; }
        public bool procesado { get; set; }
        public DateTime fechahorainicio { get; set; }
        public DateTime fechahorafinal { get; set; }
        public string observaciones { get; set; }
        public string winuser { get; set; }
        public DateTime fecharegistro { get; set; }
    }
}
