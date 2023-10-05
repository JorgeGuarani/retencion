using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;
using Aspose.Cells;
using Aspose.Cells.Cloud.SDK.Model;
using Newtonsoft.Json;
using System.IO;
using System.Data;

namespace addOnRetencion
{
    [FormAttribute("addOnRetencion.Form1", "exportartxt.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            this.lbldesde = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.lblhasta = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.txtdesde = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.txtOP = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.btnexport = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.btnexport.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnexport_ClickAfter);
            this.btncancelar = ((SAPbouiCOM.Button)(this.GetItem("Item_6").Specific));
            this.btncancelar.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btncancelar_ClickAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.txtcoti = ((SAPbouiCOM.EditText)(this.GetItem("Item_7").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        

        private void OnCustomInitialize()
        {
        }

        #region VARIABLES
        private SAPbouiCOM.StaticText lblhasta;
        private SAPbouiCOM.EditText txtdesde;
        private SAPbouiCOM.EditText txtOP;
        private SAPbouiCOM.Button btnexport;
        private SAPbouiCOM.Button btncancelar;
        private SAPbouiCOM.StaticText lbldesde;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Button btnJson;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText txtcoti;
        #endregion

        private void btncancelar_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //cerramos el form
            //crearTabla();
            Application.SBO_Application.Forms.ActiveForm.Close();
            
        }

        private static void crearTabla()
        {
            SAPbobsCOM.IUserTablesMD userTables = (SAPbobsCOM.IUserTablesMD)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            int retVal = 0;
            try
            {
                if (!userTables.GetByKey(""))
                {
                    userTables.TableName = "TABLAPRUEBA";
                    userTables.TableDescription = "TABLAPRUEBA";
                    userTables.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterData;
                    retVal = userTables.Add();
                    if (retVal != 0)
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return;
                    }
                }
            }
            catch (Exception e)
            {

            }
            userTables = null;
        }

        private void btnexport_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            //agarramos las variables y hacemos un control
            string v_fechaDesde = txtdesde.Value;
            string v_OP = txtOP.Value;
            string v_cotiMonto = txtcoti.Value;
            //verificamos que los campos no queden vacíos
            if (string.IsNullOrEmpty(v_fechaDesde))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("La fecha no puede quedar vacía!!", 1,"OK");
                return;
            }
            if (string.IsNullOrEmpty(v_OP))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe ingresar número de OP!!", 1, "OK");
                return;
            }

            //armamos el query
            string query = "SELECT "+
                           //atributos
                           "T0.\"DocDate\" AS \"fecha creacion\", "+
                           //informado
                           "CASE WHEN T2.\"U_iNatRec\"=1 THEN 'CONTRIBUYENTE' ELSE 'NO CONTRIBUYENTE' END AS \"situacion\", "+
                           "T1.\"CardName\" AS \"nombre\", "+
                           "T2.\"LicTradNum\" AS \"ruc\", "+
                           "'' AS \"domicilio\", "+
                           "'' AS \"direccion\", "+
                           "'' AS \"correo\", "+
                           "'' AS \"tipo identificacion\", "+
                           "'' AS \"identificacion\", "+
                           "'' AS \"pais\", "+
                           "'' AS \"telefono\", "+
                           //transaccion
                           "T1.\"NumAtCard\" AS \"num comprobante\", "+
                           "CASE WHEN T1.\"GroupNum\"='-1' THEN 'CONTADO' ELSE 'CREDITO' END AS \"condicion\", "+
                           "CASE WHEN T1.\"GroupNum\"='-1' THEN '0' ELSE T1.\"Installmnt\" END AS \"cuota\", "+
                           "'1' AS \"tipo comprobante\", "+
                           "T1.\"DocDate\" AS \"fecha\", "+
                           "T1.\"U_TIMB\" AS \"timbrado\", "+
                           //detalle
                           "'1' AS \"cantidad\", "+
                           "CASE WHEN T3.\"TaxCode\"='IVA_10' THEN 10 ELSE 5 END AS \"tasa aplica\", "+
                           "CASE WHEN T1.\"DocCur\"='GS' THEN T1.\"DocTotal\" ELSE T1.\"DocTotalFC\" END AS \"precio\", " +
                           "T1.\"Comments\", "+
                           //retencion
                           "T0.\"DocDate\" AS \"fecha ret\", "+
                           "CASE WHEN T1.\"DocCur\"='GS' THEN 'PYG' ELSE 'USD' END AS \"moneda\", " +
                           "CASE WHEN T4.\"U_RetReta\" IS NULL OR T4.\"U_RetReta\"=0 THEN 'false' ELSE 'true' END AS \"retencionRenta\", " +
                           "CASE WHEN T4.\"U_RetReta\" IS NULL OR T4.\"U_RetReta\"=0 THEN '' ELSE 'RENTA_EMPRESARIAL_REGISTRADO.1' END AS \"conceptoRenta\", " +
                           "CASE WHEN T4.\"U_RetIva\" IS NULL OR T4.\"U_RetIva\"=0 THEN 'false' ELSE 'true' END AS \"retencionIva\", " +
                           "CASE WHEN T4.\"U_RetIva\" IS NULL OR T4.\"U_RetIva\"=0 THEN '' ELSE 'IVA.1' END AS \"conceptoiva\", " +
                           "CASE WHEN T4.\"U_RetReta\" IS NULL THEN '0' ELSE '0,4' END AS \"rentPorcentaje\", " +
                           "'0' AS \"rentaCabezasBase\", "+
                           "'0' AS \"rentaCabezasCantidad\", "+
                           "'0' AS \"rentaToneladasBase\", "+
                           "'0' AS \"rentaToneladasCantidad\", "+
                           "CASE WHEN T3.\"TaxCode\"='IVA_5' THEN 30 ELSE 0 END AS \"ivaPorcentaje5\", "+
                           "CASE WHEN T3.\"TaxCode\"='IVA_10' THEN 70 ELSE 0 END AS \"ivaPorcentaje10\" "+
                           "FROM OVPM T0 "+
                           "INNER JOIN OPCH T1 ON T0.\"DocNum\"=T1.\"ReceiptNum\" "+
                           "INNER JOIN OCRD T2 ON T1.\"CardCode\"=T2.\"CardCode\" "+
                           "INNER JOIN PCH1 T3 ON T1.\"DocEntry\"=T3.\"DocEntry\" "+
                           "LEFT JOIN \"@RET_CALCULO\" T4 ON T1.\"DocNum\"=T4.\"U_DocNum\" "+
                           "WHERE T0.\"DocNum\"='"+ v_OP + "' AND T3.\"TaxCode\"<>'IVA_EXE' "+
                           "GROUP BY T0.\"DocDate\", T2.\"U_iNatRec\", T1.\"CardName\" , T2.\"LicTradNum\", T1.\"NumAtCard\" , T1.\"GroupNum\", T1.\"Installmnt\", T1.\"DocDate\" , T1.\"U_TIMB\" , T3.\"TaxCode\","+
                           "T1.\"DocTotal\", T1.\"Comments\", T0.\"DocDate\", T4.\"U_RetReta\",T4.\"U_RetIva\",T1.\"DocTotalFC\",T1.\"DocCur\" ";

            //consultamos las facturas de la OP
            string v_json = null;
            string v_json2 = null;
            //string v_json = "[{\"atributos\":{\"fechaCreacion\":'',\"fechaHoraCreacion\":''}," +
            //              "\"informado\":{\"situacion\":'',\"nombre\":'',\"ruc\":'',\"dv\":'',\"domicilio\":'',\"direccion\":'',\"correoElectronico\":'',\"tipoIdentificacion\":'',\"identificacion\":'',\"pais\":'',\"telefono\":''}," +
            //              "\"transaccion\":{\"numeroComprobanteVenta\":'',\"condicionCompra\":'',\"cuotas\":'',\"tipoComprobante\":'',\"fecha\":'',\"numeroTimbrado\":''}," +
            //              "\"detalle\":[{\"cantidad\":'',\"tasaAplica\":'',\"precioUnitario\":'',\"descripcion\":''}]," +
            //              "\"retencion\":{\"fecha\":'',\"moneda\":'',\"retencionRenta\":'',\"conceptoRenta\":'',\"retencionIva\":'',\"conceptoIva\":'',\"rentaPorcentaje\":0,\"rentaCabezasBase\":0,\"rentaCabezasCantidad\":0,\"rentaToneladasBase\":0,\"rentaToneladasCantidad\":0,\"ivaPorcentaje5\":'',\"ivaPorcentaje10\":''}}";

            
            SAPbobsCOM.Recordset oFacturas;
            oFacturas = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oFacturas.DoQuery(query);
            int cant = oFacturas.RecordCount;
            int ini = 1;
            //creamos el datatable y sus columnas
            DataTable dt = new DataTable("DT");
            dt.Columns.Add("fechacreacion");
            dt.Columns.Add("situacion");
            dt.Columns.Add("nombre");
            dt.Columns.Add("ruc");
            dt.Columns.Add("comprobanteventa");
            dt.Columns.Add("condicion");
            dt.Columns.Add("cuota");
            dt.Columns.Add("tipocomprobante");
            dt.Columns.Add("fecha");
            dt.Columns.Add("timbrado");
            dt.Columns.Add("cantidad");
            dt.Columns.Add("tasaplica");
            dt.Columns.Add("prciounitario");
            dt.Columns.Add("descripcion");
            dt.Columns.Add("fecharet");
            dt.Columns.Add("moneda");
            dt.Columns.Add("retencionrenta");
            dt.Columns.Add("conceptorenta");
            dt.Columns.Add("retencioniva");
            dt.Columns.Add("conceptoiva");
            dt.Columns.Add("porcentajerenta");
            dt.Columns.Add("ivaporcentaje5");
            dt.Columns.Add("ivaporcentaje10");
            //recorremos el query
            while (!oFacturas.EoF)
            {
                string v_fechacreacion = oFacturas.Fields.Item(0).Value.ToString();
                string v_situacion = oFacturas.Fields.Item(1).Value.ToString();
                string v_nombre = oFacturas.Fields.Item(2).Value.ToString();
                string v_ruc = oFacturas.Fields.Item(3).Value.ToString();
                string v_numcomprobanteventa = oFacturas.Fields.Item(11).Value.ToString();
                string v_condicion = oFacturas.Fields.Item(12).Value.ToString();
                string v_cuota = oFacturas.Fields.Item(13).Value.ToString();
                string v_tipocomprobante = oFacturas.Fields.Item(14).Value.ToString();
                string v_fecha = oFacturas.Fields.Item(15).Value.ToString();
                string v_timbrado = oFacturas.Fields.Item(16).Value.ToString();
                string v_cantidad = oFacturas.Fields.Item(17).Value.ToString();
                string v_tasaAplica = oFacturas.Fields.Item(18).Value.ToString();
                string v_precioUnitario = oFacturas.Fields.Item(19).Value.ToString();
                string v_descripcion = oFacturas.Fields.Item(20).Value.ToString();
                string v_fecharet = oFacturas.Fields.Item(21).Value.ToString();
                string v_moneda = oFacturas.Fields.Item(22).Value.ToString();
                string v_retencionRenta = oFacturas.Fields.Item(23).Value.ToString();
                string v_conceptorenta = oFacturas.Fields.Item(24).Value.ToString();
                string v_retencioniva = oFacturas.Fields.Item(25).Value.ToString();
                string v_conceptoiva = oFacturas.Fields.Item(26).Value.ToString();
                string v_porcentajeRenta = oFacturas.Fields.Item(27).Value.ToString();
                string v_ivaporcentaje5 = oFacturas.Fields.Item(32).Value.ToString();
                string v_ivaporcentaje10 = oFacturas.Fields.Item(33).Value.ToString();
                //cargamos el datatable
                dt.Rows.Add(v_fechacreacion, v_situacion, v_nombre, v_ruc, v_numcomprobanteventa, v_condicion, v_cuota, v_tipocomprobante, v_fecha, v_timbrado, v_cantidad, v_tasaAplica, v_precioUnitario,
                            v_descripcion, v_fecharet, v_moneda, v_retencionRenta, v_conceptorenta, v_retencioniva, v_conceptoiva, v_porcentajeRenta, v_ivaporcentaje5, v_ivaporcentaje10);
               
                oFacturas.MoveNext();              
            }
            dynamic objectJson;
            if (!string.IsNullOrEmpty(v_cotiMonto))
            {
                 objectJson = ConvertDataTableToArrayUSD(dt,v_cotiMonto);
            }
            else
            {
                objectJson = ConvertDataTableToArray(dt);
            }
           
            //grabamos el json en el escritorio
            var jsontowrite = JsonConvert.SerializeObject(objectJson, Newtonsoft.Json.Formatting.Indented);
            //creamos una carpeta en el escritorio para guardar el excel
            string carpeta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string CarpEscr = carpeta + "\\Retenciones-Tesaka";
            if (!Directory.Exists(CarpEscr))
            {
                Directory.CreateDirectory(CarpEscr);
            }
            string path = CarpEscr + "\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            using (var writer = new StreamWriter(path))
            {
                writer.Write(jsontowrite);
            }
            txtOP.Value = "";
            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Json exportado con éxito!!", 1, "OK");

        }

        //funcion para crear el json GS
        private static object[] ConvertDataTableToArray(DataTable dataTable)
        {
            List <object> dataArray = new List<object>();
            foreach (DataRow row in dataTable.Rows)
            {
                
                //variable a convertir en el formato correcto
                //formato de la fecha
                DateTime v_fecha = DateTime.Parse(row[0].ToString());
                DateTime v_fechaDoc = DateTime.Parse(row[8].ToString());
                //formato del ruc
                string v_ruc = row[3].ToString();
                int index = v_ruc.IndexOf("-");
                int longuitud = v_ruc.Length;
                string ruc = v_ruc.Remove(index,longuitud-index);
                string dv = v_ruc.Remove(0,index+1);
                //datos para la retencion IVA
                string retRetna = row[16].ToString();
                string retIva = row[18].ToString();

                if (retIva.Equals("true"))
                {
                    //creamos el json
                    var dataObject = new
                    {
                        atributos = new Dictionary<string, string>
                    {
                        {"fechaCreacion",v_fecha.ToString("yyyy-MM-dd")},
                        {"fechaHoraCreacion",v_fecha.ToString("yyyy-MM-dd hh:mm:ss")}
                    },
                        informado = new Dictionary<string, string>
                    {
                        {"situacion",row[1].ToString() },
                        {"nombre",row[2].ToString() },
                        {"ruc",ruc },
                        {"dv",dv },
                        {"domicilio","" },
                        {"direccion","" },
                        {"correoElectronico","" },
                        {"tipoIdentificacion","" },
                        {"identificacion","" },
                        {"pais","" },
                        {"telefono","" }
                    },
                        transaccion = new Dictionary<string, object>
                    {
                        {"condicionCompra",row[5].ToString() },
                        {"numeroComprobanteVenta",row[4].ToString() },
                        {"cuotas",int.Parse(row[6].ToString()) },
                        {"tipoComprobante",int.Parse(row[7].ToString()) },
                        {"fecha",v_fechaDoc.ToString("yyyy-MM-dd") },
                        {"numeroTimbrado",row[9].ToString() }
                    },
                        detalle = new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object>
                        {
                            { "cantidad", int.Parse(row[10].ToString()) },
                            { "tasaAplica", row[11].ToString() },
                            { "precioUnitario", double.Parse(row[12].ToString()) },
                            { "descripcion", row[13].ToString() }
                        }
                     },
                        retencion = new Dictionary<string, object>
                    {
                        { "fecha",v_fecha.ToString("yyyy-MM-dd") },
                        { "moneda",row[15].ToString() },
                        { "retencionRenta",bool.Parse("false") },
                        { "conceptoRenta","" },
                        { "retencionIva",bool.Parse(row[18].ToString()) },
                        { "conceptoIva",row[19].ToString() },
                        { "rentaPorcentaje",0 },
                        { "rentaCabezasBase",0 },
                        { "rentaCabezasCantidad",0 },
                        { "rentaToneladasBase",0},
                        { "rentaToneladasCantidad",0 },
                        { "ivaPorcentaje5",int.Parse(row[21].ToString()) },
                        { "ivaPorcentaje10",int.Parse(row[22].ToString()) }
                    }
                    };
                    dataArray.Add(dataObject);
                }
                if (retRetna.Equals("true"))
                {
                    //creamos el json
                    var dataObject = new
                    {
                        atributos = new Dictionary<string, string>
                    {
                        {"fechaCreacion",v_fecha.ToString("yyyy-MM-dd")},
                        {"fechaHoraCreacion",v_fecha.ToString("yyyy-MM-dd hh:mm:ss")}
                    },
                        informado = new Dictionary<string, string>
                    {
                        {"situacion",row[1].ToString() },
                        {"nombre",row[2].ToString() },
                        {"ruc",ruc },
                        {"dv",dv },
                        {"domicilio","" },
                        {"direccion","" },
                        {"correoElectronico","" },
                        {"tipoIdentificacion","" },
                        {"identificacion","" },
                        {"pais","" },
                        {"telefono","" }
                    },
                        transaccion = new Dictionary<string, object>
                    {
                        {"condicionCompra",row[5].ToString() },
                        {"numeroComprobanteVenta",row[4].ToString() },
                        {"cuotas",int.Parse(row[6].ToString()) },
                        {"tipoComprobante",int.Parse(row[7].ToString()) },
                        {"fecha",v_fecha.ToString("yyyy-MM-dd") },
                        {"numeroTimbrado",row[9].ToString() }
                    },
                        detalle = new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object>
                        {
                            { "cantidad", int.Parse(row[10].ToString()) },
                            { "tasaAplica", row[11].ToString() },
                            { "precioUnitario", double.Parse(row[12].ToString()) },
                            { "descripcion", row[13].ToString() }
                        }
                     },
                        retencion = new Dictionary<string, object>
                    {
                        { "fecha",v_fecha.ToString("yyyy-MM-dd") },
                        { "moneda",row[15].ToString() },
                        { "retencionRenta",bool.Parse(row[16].ToString()) },
                        { "conceptoRenta",row[17].ToString() },
                        { "retencionIva",bool.Parse("false") },
                        { "conceptoIva","" },
                        { "rentaPorcentaje",double.Parse(row[20].ToString()) },
                        { "rentaCabezasBase",0 },
                        { "rentaCabezasCantidad",0 },
                        { "rentaToneladasBase",0},
                        { "rentaToneladasCantidad",0 },
                        { "ivaPorcentaje5",0},
                        { "ivaPorcentaje10",0 }
                    }
                    };
                    dataArray.Add(dataObject);
                }
                
            }
            return dataArray.ToArray();            
        }

        //funcion para crear el json USD
        private static object[] ConvertDataTableToArrayUSD(DataTable dataTable,string coti)
        {
            
            List<object> dataArray = new List<object>();
            foreach (DataRow row in dataTable.Rows)
            {

                //variable a convertir en el formato correcto
                //formato de la fecha
                DateTime v_fecha = DateTime.Parse(row[0].ToString());
                //formato del ruc
                string v_ruc = row[3].ToString();
                int index = v_ruc.IndexOf("-");
                int longuitud = v_ruc.Length;
                string ruc = v_ruc.Remove(index, longuitud - index);
                string dv = v_ruc.Remove(0, index + 1);
                //datos para la retencion IVA
                string retRetna = row[16].ToString();
                string retIva = row[18].ToString();

                if (retIva.Equals("true"))
                {
                    //creamos el json
                    var dataObject = new
                    {
                        atributos = new Dictionary<string, string>
                    {
                        {"fechaCreacion",v_fecha.ToString("yyyy-MM-dd")},
                        {"fechaHoraCreacion",v_fecha.ToString("yyyy-MM-dd hh:mm:ss")}
                    },
                        informado = new Dictionary<string, string>
                    {
                        {"situacion",row[1].ToString() },
                        {"nombre",row[2].ToString() },
                        {"ruc",ruc },
                        {"dv",dv },
                        {"domicilio","" },
                        {"direccion","" },
                        {"correoElectronico","" },
                        {"tipoIdentificacion","" },
                        {"identificacion","" },
                        {"pais","" },
                        {"telefono","" }
                    },
                        transaccion = new Dictionary<string, object>
                    {
                        {"condicionCompra",row[5].ToString() },
                        {"numeroComprobanteVenta",row[4].ToString() },
                        {"cuotas",int.Parse(row[6].ToString()) },
                        {"tipoComprobante",int.Parse(row[7].ToString()) },
                        {"fecha",v_fecha.ToString("yyyy-MM-dd") },
                        {"numeroTimbrado",row[9].ToString() }
                    },
                        detalle = new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object>
                        {
                            { "cantidad", int.Parse(row[10].ToString()) },
                            { "tasaAplica", row[11].ToString() },
                            { "precioUnitario", double.Parse(row[12].ToString()) },
                            { "descripcion", row[13].ToString() }
                        }
                     },
                        retencion = new Dictionary<string, object>
                    {
                        { "fecha",v_fecha.ToString("yyyy-MM-dd") },
                        { "moneda",row[15].ToString() },
                        { "tipoCambio",int.Parse(coti) },
                        { "retencionRenta",bool.Parse("false") },
                        { "conceptoRenta","" },
                        { "retencionIva",bool.Parse(row[18].ToString()) },
                        { "conceptoIva",row[19].ToString() },
                        { "rentaPorcentaje",0 },
                        { "rentaCabezasBase",0 },
                        { "rentaCabezasCantidad",0 },
                        { "rentaToneladasBase",0},
                        { "rentaToneladasCantidad",0 },
                        { "ivaPorcentaje5",int.Parse(row[21].ToString()) },
                        { "ivaPorcentaje10",int.Parse(row[22].ToString()) }
                    }
                    };
                    dataArray.Add(dataObject);
                }
                if (retRetna.Equals("true"))
                {
                    //creamos el json
                    var dataObject = new
                    {
                        atributos = new Dictionary<string, string>
                    {
                        {"fechaCreacion",v_fecha.ToString("yyyy-MM-dd")},
                        {"fechaHoraCreacion",v_fecha.ToString("yyyy-MM-dd hh:mm:ss")}
                    },
                        informado = new Dictionary<string, string>
                    {
                        {"situacion",row[1].ToString() },
                        {"nombre",row[2].ToString() },
                        {"ruc",ruc },
                        {"dv",dv },
                        {"domicilio","" },
                        {"direccion","" },
                        {"correoElectronico","" },
                        {"tipoIdentificacion","" },
                        {"identificacion","" },
                        {"pais","" },
                        {"telefono","" }
                    },
                        transaccion = new Dictionary<string, object>
                    {
                        {"condicionCompra",row[5].ToString() },
                        {"numeroComprobanteVenta",row[4].ToString() },
                        {"cuotas",int.Parse(row[6].ToString()) },
                        {"tipoComprobante",int.Parse(row[7].ToString()) },
                        {"fecha",v_fecha.ToString("yyyy-MM-dd") },
                        {"numeroTimbrado",row[9].ToString() }
                    },
                        detalle = new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object>
                        {
                            { "cantidad", int.Parse(row[10].ToString()) },
                            { "tasaAplica", row[11].ToString() },
                            { "precioUnitario", double.Parse(row[12].ToString()) },
                            { "descripcion", row[13].ToString() }
                        }
                     },
                        retencion = new Dictionary<string, object>
                    {
                        { "fecha",v_fecha.ToString("yyyy-MM-dd") },
                        { "moneda",row[15].ToString() },
                        { "tipoCambio",int.Parse(coti) },
                        { "retencionRenta",bool.Parse(row[16].ToString()) },
                        { "conceptoRenta",row[17].ToString() },
                        { "retencionIva",bool.Parse("false") },
                        { "conceptoIva","" },
                        { "rentaPorcentaje",double.Parse(row[20].ToString()) },
                        { "rentaCabezasBase",0 },
                        { "rentaCabezasCantidad",0 },
                        { "rentaToneladasBase",0},
                        { "rentaToneladasCantidad",0 },
                        { "ivaPorcentaje5",0},
                        { "ivaPorcentaje10",0 }
                    }
                    };
                    dataArray.Add(dataObject);
                }

            }
            return dataArray.ToArray();
        }


    }
}