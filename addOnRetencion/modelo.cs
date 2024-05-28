using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace addOnRetencion
{
    class modelo
    {
        //clases para el cuerpo del json
        public class atributos
        {
            public string fechaCreacion { get; set; }
            public string fechaHoraCreacion { get; set; }
        }

        public class informado
        {
            public string situacion { get; set; }
            public string nombre { get; set; }
            public string ruc { get; set; }
            public string dv { get; set; }
            public string domicilio { get; set; }
            public string direccion { get; set; }
            public string correoElectronico { get; set; }
            public string tipoIdentificacion { get; set; }
            public string identificacion { get; set; }
            public string pais { get; set; }
            public string telefono { get; set; }
        }

        public class transaccion
        {
            public string condicionCompra { get; set; }
            public string numeroComprobanteVenta { get; set; }
            public int cuotas { get; set; }
            public int tipoComprobante { get; set; }
            public string fecha { get; set; }
            public string numeroTimbrado { get; set; }
        }

        public class retencion
        {
            public string fecha { get; set; }
            public string moneda { get; set; }
            public int tipoCambio { get; set; }
            public bool retencionRenta { get; set; }
            public string conceptoRenta { get; set; }
            public bool retencionIva { get; set; }
            public string conceptoIva { get; set; }
            public double rentaPorcentaje { get; set; }
            public int rentaCabezasBase { get; set; }
            public int rentaCabezasCantidad { get; set; }
            public int rentaToneladasBase { get; set; }
            public int rentaToneladasCantidad { get; set; }
            public int ivaPorcentaje5 { get; set; }
            public int ivaPorcentaje10 { get; set; }
        }

        public class detalle
        {
            public int cantidad { get; set; }
            public string tasaAplica { get; set; }
            public double precioUnitario { get; set; }
            public string descripcion { get; set; }
        }

        //json general para el tesaka
        public class jsonTesaka
        {
            public object atributos { get; set; }
            public object informado { get; set; }
            public object transaccion { get; set; }
            public object retencion { get; set; }
            public List<detalle> detalle { get; set; }
        }

    }
}
