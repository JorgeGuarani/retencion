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
using SAPbobsCOM;
using System.Threading;
using System.Windows.Forms;
using System.DirectoryServices.AccountManagement;

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
            //this.oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
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
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_8").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_12").Specific));
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

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
            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
            
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
            //try
            //{
            //    proceso_exportar();
            //}
            //catch (Exception e)
            //{
            //    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.ToString(), 1, "OK");
            //}

            //throw new System.NotImplementedException();
            //agarramos las variables y hacemos un control
            //string v_fechaDesde = txtdesde.Value;
            //string v_OP = txtOP.Value;
            //string v_cotiMonto = txtcoti.Value;
            ////verificamos que los campos no queden vacíos
            //if (string.IsNullOrEmpty(v_fechaDesde))
            //{
            //    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("La fecha no puede quedar vacía!!", 1,"OK");
            //    return;
            //}
            //if (string.IsNullOrEmpty(v_OP))
            //{
            //    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe ingresar número de OP!!", 1, "OK");
            //    return;
            //}

            ////armamos el query
            //string query = "SELECT "+
            //               //atributos
            //               "T0.\"DocDate\" AS \"fecha creacion\", "+
            //               //informado
            //               "CASE WHEN T2.\"U_iNatRec\"=1 THEN 'CONTRIBUYENTE' ELSE 'NO CONTRIBUYENTE' END AS \"situacion\", "+
            //               "T1.\"CardName\" AS \"nombre\", "+
            //               "T2.\"LicTradNum\" AS \"ruc\", "+
            //               "'' AS \"domicilio\", "+
            //               "'' AS \"direccion\", "+
            //               "'' AS \"correo\", "+
            //               "'' AS \"tipo identificacion\", "+
            //               "'' AS \"identificacion\", "+
            //               "'' AS \"pais\", "+
            //               "'' AS \"telefono\", "+
            //               //transaccion
            //               "T1.\"NumAtCard\" AS \"num comprobante\", "+
            //               "CASE WHEN T1.\"GroupNum\"='-1' THEN 'CONTADO' ELSE 'CREDITO' END AS \"condicion\", "+
            //               "CASE WHEN T1.\"GroupNum\"='-1' THEN '0' ELSE T1.\"Installmnt\" END AS \"cuota\", "+
            //               "'1' AS \"tipo comprobante\", "+
            //               "T1.\"DocDate\" AS \"fecha\", "+
            //               "T1.\"U_TIMB\" AS \"timbrado\", "+
            //               //detalle
            //               "'1' AS \"cantidad\", "+
            //               "CASE WHEN T3.\"TaxCode\"='IVA_10' THEN 10 ELSE 5 END AS \"tasa aplica\", " +
            //               "CASE WHEN T1.\"DocCur\"='GS' THEN T1.\"DocTotal\" ELSE T1.\"DocTotalFC\" END AS \"precio\", " +
            //               "T1.\"Comments\", " +
            //               //retencion
            //               "T0.\"DocDate\" AS \"fecha ret\", "+
            //               "CASE WHEN T1.\"DocCur\"='GS' THEN 'PYG' ELSE 'USD' END AS \"moneda\", " +
            //               "CASE WHEN (T4.\"U_RetReta\" IS NULL OR T4.\"U_RetReta\"=0) THEN 'false' ELSE 'true' END AS \"retencionRenta\", " +
            //               "CASE WHEN T4.\"U_RetReta\" IS NULL OR T4.\"U_RetReta\"=0 THEN '' ELSE 'RENTA_EMPRESARIAL_REGISTRADO.1' END AS \"conceptoRenta\", " +
            //               "CASE WHEN (T4.\"U_RetIva\" IS NULL OR T4.\"U_RetIva\"=0) THEN 'false' ELSE 'true' END AS \"retencionIva\", " +
            //               "CASE WHEN T4.\"U_RetIva\" IS NULL OR T4.\"U_RetIva\"=0 THEN '' ELSE 'IVA.1' END AS \"conceptoiva\", " +
            //               "CASE WHEN T4.\"U_RetReta\" IS NULL THEN '0' ELSE '0.4' END AS \"rentPorcentaje\", " +
            //               "'0' AS \"rentaCabezasBase\", "+
            //               "'0' AS \"rentaCabezasCantidad\", "+
            //               "'0' AS \"rentaToneladasBase\", "+
            //               "'0' AS \"rentaToneladasCantidad\", "+
            //               "CASE WHEN T3.\"TaxCode\"='IVA_5' THEN 30 ELSE 0 END AS \"ivaPorcentaje5\", "+
            //               "CASE WHEN T3.\"TaxCode\"='IVA_10' THEN 70 ELSE 0 END AS \"ivaPorcentaje10\" "+
            //               "FROM OVPM T0 "+
            //               "INNER JOIN OPCH T1 ON T0.\"DocEntry\"=T1.\"ReceiptNum\" "+
            //               "INNER JOIN OCRD T2 ON T1.\"CardCode\"=T2.\"CardCode\" "+
            //               "INNER JOIN PCH1 T3 ON T1.\"DocEntry\"=T3.\"DocEntry\" "+
            //               "LEFT JOIN \"@RET_CALCULO\" T4 ON T1.\"DocNum\"=T4.\"U_DocNum\" "+
            //               "WHERE T0.\"DocNum\"='"+ v_OP + "' AND T3.\"TaxCode\"<>'IVA_EXE' "+
            //               "GROUP BY T0.\"DocDate\", T2.\"U_iNatRec\", T1.\"CardName\" , T2.\"LicTradNum\", T1.\"NumAtCard\" , T1.\"GroupNum\", T1.\"Installmnt\", T1.\"DocDate\" , T1.\"U_TIMB\" , T3.\"TaxCode\","+
            //               "T1.\"DocTotal\", T1.\"Comments\", T0.\"DocDate\", T4.\"U_RetReta\",T4.\"U_RetIva\",T1.\"DocTotalFC\",T1.\"DocCur\" ";

            ////consultamos las facturas de la OP
            //string v_json = null;
            //string v_json2 = null;
            ////string v_json = "[{\"atributos\":{\"fechaCreacion\":'',\"fechaHoraCreacion\":''}," +
            ////              "\"informado\":{\"situacion\":'',\"nombre\":'',\"ruc\":'',\"dv\":'',\"domicilio\":'',\"direccion\":'',\"correoElectronico\":'',\"tipoIdentificacion\":'',\"identificacion\":'',\"pais\":'',\"telefono\":''}," +
            ////              "\"transaccion\":{\"numeroComprobanteVenta\":'',\"condicionCompra\":'',\"cuotas\":'',\"tipoComprobante\":'',\"fecha\":'',\"numeroTimbrado\":''}," +
            ////              "\"detalle\":[{\"cantidad\":'',\"tasaAplica\":'',\"precioUnitario\":'',\"descripcion\":''}]," +
            ////              "\"retencion\":{\"fecha\":'',\"moneda\":'',\"retencionRenta\":'',\"conceptoRenta\":'',\"retencionIva\":'',\"conceptoIva\":'',\"rentaPorcentaje\":0,\"rentaCabezasBase\":0,\"rentaCabezasCantidad\":0,\"rentaToneladasBase\":0,\"rentaToneladasCantidad\":0,\"ivaPorcentaje5\":'',\"ivaPorcentaje10\":''}}";


            //SAPbobsCOM.Recordset oFacturas;
            //oFacturas = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //oFacturas.DoQuery(query);
            //int cant = oFacturas.RecordCount;
            //int ini = 1;
            ////creamos el datatable y sus columnas
            //DataTable dt = new DataTable("DT");
            //dt.Columns.Add("fechacreacion");
            //dt.Columns.Add("situacion");
            //dt.Columns.Add("nombre");
            //dt.Columns.Add("ruc");
            //dt.Columns.Add("comprobanteventa");
            //dt.Columns.Add("condicion");
            //dt.Columns.Add("cuota");
            //dt.Columns.Add("tipocomprobante");
            //dt.Columns.Add("fecha");
            //dt.Columns.Add("timbrado");
            //dt.Columns.Add("cantidad");
            //dt.Columns.Add("tasaplica");
            //dt.Columns.Add("prciounitario");
            //dt.Columns.Add("descripcion");
            //dt.Columns.Add("fecharet");
            //dt.Columns.Add("moneda");
            //dt.Columns.Add("retencionrenta");
            //dt.Columns.Add("conceptorenta");
            //dt.Columns.Add("retencioniva");
            //dt.Columns.Add("conceptoiva");
            //dt.Columns.Add("porcentajerenta");
            //dt.Columns.Add("ivaporcentaje5");
            //dt.Columns.Add("ivaporcentaje10");
            ////recorremos el query
            //while (!oFacturas.EoF)
            //{
            //    string v_fechacreacion = oFacturas.Fields.Item(0).Value.ToString();
            //    string v_situacion = oFacturas.Fields.Item(1).Value.ToString();
            //    string v_nombre = oFacturas.Fields.Item(2).Value.ToString();
            //    string v_ruc = oFacturas.Fields.Item(3).Value.ToString();
            //    string v_numcomprobanteventa = oFacturas.Fields.Item(11).Value.ToString();
            //    string v_condicion = oFacturas.Fields.Item(12).Value.ToString();
            //    string v_cuota = oFacturas.Fields.Item(13).Value.ToString();
            //    string v_tipocomprobante = oFacturas.Fields.Item(14).Value.ToString();
            //    string v_fecha = oFacturas.Fields.Item(15).Value.ToString();
            //    string v_timbrado = oFacturas.Fields.Item(16).Value.ToString();
            //    string v_cantidad = oFacturas.Fields.Item(17).Value.ToString();
            //    string v_tasaAplica = oFacturas.Fields.Item(18).Value.ToString();
            //    string v_precioUnitario = oFacturas.Fields.Item(19).Value.ToString();
            //    string v_descripcion = oFacturas.Fields.Item(20).Value.ToString();
            //    string v_fecharet = oFacturas.Fields.Item(21).Value.ToString();
            //    string v_moneda = oFacturas.Fields.Item(22).Value.ToString();
            //    string v_retencionRenta = oFacturas.Fields.Item(23).Value.ToString();
            //    string v_conceptorenta = oFacturas.Fields.Item(24).Value.ToString();
            //    string v_retencioniva = oFacturas.Fields.Item(25).Value.ToString();
            //    string v_conceptoiva = oFacturas.Fields.Item(26).Value.ToString();
            //    string v_porcentajeRenta = oFacturas.Fields.Item(27).Value.ToString();
            //    string v_ivaporcentaje5 = oFacturas.Fields.Item(32).Value.ToString();
            //    string v_ivaporcentaje10 = oFacturas.Fields.Item(33).Value.ToString();
            //    //cargamos el datatable
            //    dt.Rows.Add(v_fechacreacion, v_situacion, v_nombre, v_ruc, v_numcomprobanteventa, v_condicion, v_cuota, v_tipocomprobante, v_fecha, v_timbrado, v_cantidad, v_tasaAplica, v_precioUnitario,
            //                v_descripcion, v_fecharet, v_moneda, v_retencionRenta, v_conceptorenta, v_retencioniva, v_conceptoiva, v_porcentajeRenta, v_ivaporcentaje5, v_ivaporcentaje10);

            //    oFacturas.MoveNext();              
            //}
            //dynamic objectJson;
            //if (!string.IsNullOrEmpty(v_cotiMonto))
            //{
            //     objectJson = ConvertDataTableToArrayUSD(dt,v_cotiMonto);
            //}
            //else
            //{
            //    objectJson = ConvertDataTableToArray(dt);
            //}

            //grabamos el json en el escritorio
            //var jsontowrite = JsonConvert.SerializeObject(objectJson, Newtonsoft.Json.Formatting.Indented);
            ////creamos una carpeta en el escritorio para guardar el excel
            //string carpeta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string CarpEscr = carpeta + "\\Retenciones-Tesaka";
            //if (!Directory.Exists(CarpEscr))
            //{
            //    Directory.CreateDirectory(CarpEscr);
            //}
            ////prueba
            //string v_userAD = System.DirectoryServices.AccountManagement.UserPrincipal.Current.SamAccountName;
            ////string path = CarpEscr + "\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            //string path = null;
            ////if (SAPbouiCOM.Framework.Application.SBO_Application.ClientType == SAPbouiCOM.BoClientType.ct_Desktop)
            ////{
            ////    path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Retenciones-Tesaka\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            ////    using (var writer = new StreamWriter(path))
            ////    {
            ////        writer.Write(jsontowrite);
            ////    }
            ////}
            ////else
            ////{
            ////    path = "C:\\Users\\" + v_userAD + "\\Desktop\\Retenciones-Tesaka\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            ////    SAPbouiCOM.Framework.Application.SBO_Application.SendFileToBrowser(path);
            ////}           
            //txtOP.Value = "";
            //path = "C:\\Users\\" + v_userAD + "\\Desktop\\Retenciones-Tesaka\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            //using (var writer = new StreamWriter(path))
            //{
            //    writer.Write(jsontowrite);
            //}
            ////SAPbouiCOM.Framework.Application.SBO_Application.SendFileToBrowser(path);
            //SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Json exportado con éxito!!", 1, "OK");

            //PRUEBA PARA EXPORTAR
            //throw new System.NotImplementedException();
            //agarramos las variables y hacemos un control
            string v_fechaDesde = txtdesde.Value;
            string v_OP = txtOP.Value;
            string v_cotiMonto = txtcoti.Value;
            //verificamos que los campos no queden vacíos
            if (string.IsNullOrEmpty(v_fechaDesde))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("La fecha no puede quedar vacía!!", 1, "OK");
                return;
            }
            if (string.IsNullOrEmpty(v_OP))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe ingresar número de OP!!", 1, "OK");
                return;
            }

            //armamos el query
            string query = "SELECT " +
                           //atributos
                           "T0.\"DocDate\" AS \"fecha creacion\", " +
                           //informado
                           "CASE WHEN T2.\"U_iNatRec\"=1 THEN 'CONTRIBUYENTE' ELSE 'NO CONTRIBUYENTE' END AS \"situacion\", " +
                           "T1.\"CardName\" AS \"nombre\", " +
                           "T2.\"LicTradNum\" AS \"ruc\", " +
                           "'' AS \"domicilio\", " +
                           "'' AS \"direccion\", " +
                           "'' AS \"correo\", " +
                           "'' AS \"tipo identificacion\", " +
                           "'' AS \"identificacion\", " +
                           "'' AS \"pais\", " +
                           "'' AS \"telefono\", " +
                           //transaccion
                           "T1.\"NumAtCard\" AS \"num comprobante\", " +
                           "CASE WHEN T1.\"GroupNum\"='-1' THEN 'CONTADO' ELSE 'CREDITO' END AS \"condicion\", " +
                           "CASE WHEN T1.\"GroupNum\"='-1' THEN '0' ELSE T1.\"Installmnt\" END AS \"cuota\", " +
                           "'1' AS \"tipo comprobante\", " +
                           "T1.\"DocDate\" AS \"fecha\", " +
                           "T1.\"U_TIMB\" AS \"timbrado\", " +
                           //detalle
                           "'1' AS \"cantidad\", " +
                           "CASE WHEN T3.\"TaxCode\"='IVA_10' THEN 10 WHEN T3.\"TaxCode\"='IVA_5' THEN 5 ELSE 0 END AS \"tasa aplica\", " +
                           "T3.\"PriceAfVAT\" * CASE WHEN T3.\"Quantity\" = 0 THEN 1 ELSE T3.\"Quantity\" END AS \"precio\", " +
                           "T1.\"Comments\", " +
                           //retencion
                           "T0.\"DocDate\" AS \"fecha ret\", " +
                           "CASE WHEN T1.\"DocCur\"='GS' THEN 'PYG' ELSE 'USD' END AS \"moneda\", " +
                           "CASE WHEN (T4.\"U_RetReta\" IS NULL OR T4.\"U_RetReta\"=0) THEN 'false' ELSE 'true' END AS \"retencionRenta\", " +
                           "CASE WHEN T4.\"U_RetReta\" IS NULL OR T4.\"U_RetReta\"=0 THEN '' ELSE 'RENTA_EMPRESARIAL_REGISTRADO.1' END AS \"conceptoRenta\", " +
                           "CASE WHEN (T4.\"U_RetIva\" IS NULL OR T4.\"U_RetIva\"=0) THEN 'false' ELSE 'true' END AS \"retencionIva\", " +
                           "CASE WHEN T4.\"U_RetIva\" IS NULL OR T4.\"U_RetIva\"=0 THEN '' ELSE 'IVA.1' END AS \"conceptoiva\", " +
                           "CASE WHEN T4.\"U_RetReta\" IS NULL  THEN '0' WHEN T3.\"TaxCode\"='IVA_EXE' THEN '0' ELSE '0.4' END AS \"rentPorcentaje\", " +
                           "'0' AS \"rentaCabezasBase\", " +
                           "'0' AS \"rentaCabezasCantidad\", " +
                           "'0' AS \"rentaToneladasBase\", " +
                           "'0' AS \"rentaToneladasCantidad\", " +
                           "CASE WHEN T3.\"TaxCode\"='IVA_5' THEN 30 ELSE 0 END AS \"ivaPorcentaje5\", " +
                           "CASE WHEN T3.\"TaxCode\"='IVA_10' THEN 70 ELSE 0 END AS \"ivaPorcentaje10\" " +
                           "FROM OVPM T0 " +
                           "INNER JOIN OPCH T1 ON T0.\"DocEntry\"=T1.\"ReceiptNum\" " +
                           "INNER JOIN OCRD T2 ON T1.\"CardCode\"=T2.\"CardCode\" " +
                           "INNER JOIN PCH1 T3 ON T1.\"DocEntry\"=T3.\"DocEntry\" " +
                           "LEFT JOIN \"@RET_CALCULO\" T4 ON T1.\"DocNum\"=T4.\"U_DocNum\" " +
                           "WHERE T0.\"DocNum\"='" + v_OP + "' "; //+
                           //"GROUP BY T0.\"DocDate\", T2.\"U_iNatRec\", T1.\"CardName\" , T2.\"LicTradNum\", T1.\"NumAtCard\" , T1.\"GroupNum\", T1.\"Installmnt\", T1.\"DocDate\" , T1.\"U_TIMB\" , T3.\"TaxCode\"," +
                           //"T1.\"DocTotal\", T1.\"Comments\", T0.\"DocDate\", T4.\"U_RetReta\",T4.\"U_RetIva\",T1.\"DocTotalFC\",T1.\"DocCur\",T3.\"PriceAfVAT\",T3.\"Quantity\"  ";

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
                objectJson = ConvertDataTableToArrayUSDNew(dt, v_cotiMonto, v_OP);
                
            }
            else
            {
                v_cotiMonto = "0";
                objectJson = ConvertDataTableToArrayNew(dt, v_cotiMonto, v_OP);
            }

            //grabamos el json en el escritorio
            var jsontowrite = JsonConvert.SerializeObject(objectJson, Newtonsoft.Json.Formatting.Indented);

            //creamos una carpeta en el escritorio para guardar el excel
            //string carpeta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string CarpEscr = carpeta + "\\Retenciones-Tesaka";
            //if (!Directory.Exists(CarpEscr))
            //{
            //    Directory.CreateDirectory(CarpEscr);
            //}
            //prueba
            string v_userAD = System.DirectoryServices.AccountManagement.UserPrincipal.Current.SamAccountName;
            //string path = CarpEscr + "\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            string path = null;
            //if (SAPbouiCOM.Framework.Application.SBO_Application.ClientType == SAPbouiCOM.BoClientType.ct_Desktop)
            //{
            //    path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Retenciones-Tesaka\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            //    using (var writer = new StreamWriter(path))
            //    {
            //        writer.Write(jsontowrite);
            //    }
            //}
            //else
            //{
            //    path = "C:\\Users\\" + v_userAD + "\\Desktop\\Retenciones-Tesaka\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            //    SAPbouiCOM.Framework.Application.SBO_Application.SendFileToBrowser(path);
            //}           
            txtOP.Value = "";
            path = "C:\\Users\\" + v_userAD + "\\Desktop\\Retenciones-Tesaka\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            using (var writer = new StreamWriter(path))
            {
                writer.Write(jsontowrite);
            }
            //SAPbouiCOM.Framework.Application.SBO_Application.SendFileToBrowser(path);
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
                DateTime v_fechaDoc = DateTime.Parse(row[8].ToString());
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

        private SAPbouiCOM.Button Button0;

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string v_codi = txtOP.Value;
            string v_url = EditText0.Value;
            try
            {
                leerTXTversion10(v_url, v_codi);
            }
            catch(Exception e)
            {

            }           
        }

        //funcion para leer el json
        private static void leerTXTversion10(string url, string cod)
        {
            //listamos las facturas del pago
            SAPbobsCOM.Recordset oFacturas;
            oFacturas = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
            oFacturas.DoQuery("SELECT T1.\"DocNum\",T1.\"NumAtCard\",T1.\"DocDate\",T0.\"DocEntry\" FROM OVPM T0 INNER JOIN OPCH T1 ON T0.\"DocEntry\" = T1.\"ReceiptNum\" WHERE T0.\"DocNum\" = '" + cod + "'");
            while (!oFacturas.EoF)
            {
                string v_docnum = oFacturas.Fields.Item(0).Value.ToString();
                string v_nroPago = oFacturas.Fields.Item(3).Value.ToString();
                string v_nrofac = oFacturas.Fields.Item(1).Value.ToString();
                string v_fechadoc = oFacturas.Fields.Item(2).Value.ToString();
                DateTime fecha_v = DateTime.Parse(v_fechadoc);
                string fechadoc = fecha_v.ToString("yyyyMMdd");
                int v_invoiceid = 0;
                //consultamos el numero de factura
                SAPbobsCOM.Recordset oConsulta;
                oConsulta = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                //oConsulta.DoQuery("SELECT \"DocEntry\",\"NumAtCard\" FROM OPCH WHERE \"DocNum\"='" + v_docnum + "' ");
                oConsulta.DoQuery("SELECT \"DocEntry\" FROM OPCH WHERE \"NumAtCard\"='" + v_nrofac + "' AND \"DocDate\"='" + fechadoc + "' AND \"CANCELED\"='N'");
                //string v_factura = oConsulta.Fields.Item(1).Value.ToString();
                string v_DocEntry = oConsulta.Fields.Item(0).Value.ToString();
                //agarramos el JSON
                string v_json = File.ReadAllText(url);
                JArray jsonArray = JArray.Parse(v_json);
                //recoremos el json
                foreach (JObject jsonOperaciones in jsonArray.Children<JObject>())
                {
                    //extraemos numero de factura
                    //string v_datos = jsonOperaciones["datos"].ToString();
                    //JObject oDetDatos = JObject.Parse(v_datos);
                    string v_transaccion = jsonOperaciones["transaccion"].ToString();
                    JObject oTransDatos = JObject.Parse(v_transaccion);
                    string v_comprobante = oTransDatos["numeroComprobanteVenta"].ToString();
                    if (v_nrofac.Contains(v_comprobante))
                    {
                        //extraemos numero de retencion
                        string v_recepcion = jsonOperaciones["recepcion"].ToString();
                        JObject oRecepcion = JObject.Parse(v_recepcion);
                        string v_retencionNro = oRecepcion["numeroComprobante"].ToString();
                        //instanciamos el objeto
                        SAPbobsCOM.Payments oPagos;
                        oPagos = (SAPbobsCOM.Payments)Menu.sbo.GetBusinessObject(BoObjectTypes.oVendorPayments);
                        if (oPagos.GetByKey(int.Parse(v_nroPago)))
                        {
                            //oPagos.Invoices.DocEntry = int.Parse("314950");
                            oPagos.Invoices.SetCurrentLine(v_invoiceid);
                            oPagos.Invoices.UserFields.Fields.Item("U_NroRet").Value = v_retencionNro;
                            int up = oPagos.Update();
                            string pp = Menu.sbo.GetLastErrorDescription();
                            if (up != 0)
                            {
                                //System.Windows.Forms.MessageBox.Show(Menu.sbo.GetLastErrorDescription());
                                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(Menu.sbo.GetLastErrorDescription(), 1, "OK");
                            }
                            v_invoiceid++;
                            break;
                        }
                    }
                }
                oFacturas.MoveNext();
            }
            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Número de retención actualizado con éxito!!", 1, "OK");

        }

        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText0;

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();

        }

        private SAPbouiCOM.Button Button1;

        private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Thread t = new Thread(() =>
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();

                DialogResult dr = openFileDialog.ShowDialog();
                if (dr == DialogResult.OK)
                {
                    string fileName = openFileDialog.FileName;
                    //FILE.Value = fileName;
                    EditText0.Value = fileName;
                    //SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(fileName);
                }
            });          // Kick off a new thread
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();

        }

        //nuevo proceso para exportar
        private void proceso_exportar()
        {
            //throw new System.NotImplementedException();
            //agarramos las variables y hacemos un control
            string v_fechaDesde = txtdesde.Value;
            string v_OP = txtOP.Value;
            string v_cotiMonto = txtcoti.Value;
            //verificamos que los campos no queden vacíos
            if (string.IsNullOrEmpty(v_fechaDesde))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("La fecha no puede quedar vacía!!", 1, "OK");
                return;
            }
            if (string.IsNullOrEmpty(v_OP))
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe ingresar número de OP!!", 1, "OK");
                return;
            }

            //armamos el query
            string query = "SELECT " +
                           //atributos
                           "T0.\"DocDate\" AS \"fecha creacion\", " +
                           //informado
                           "CASE WHEN T2.\"U_iNatRec\"=1 THEN 'CONTRIBUYENTE' ELSE 'NO CONTRIBUYENTE' END AS \"situacion\", " +
                           "T1.\"CardName\" AS \"nombre\", " +
                           "T2.\"LicTradNum\" AS \"ruc\", " +
                           "'' AS \"domicilio\", " +
                           "'' AS \"direccion\", " +
                           "'' AS \"correo\", " +
                           "'' AS \"tipo identificacion\", " +
                           "'' AS \"identificacion\", " +
                           "'' AS \"pais\", " +
                           "'' AS \"telefono\", " +
                           //transaccion
                           "T1.\"NumAtCard\" AS \"num comprobante\", " +
                           "CASE WHEN T1.\"GroupNum\"='-1' THEN 'CONTADO' ELSE 'CREDITO' END AS \"condicion\", " +
                           "CASE WHEN T1.\"GroupNum\"='-1' THEN '0' ELSE T1.\"Installmnt\" END AS \"cuota\", " +
                           "'1' AS \"tipo comprobante\", " +
                           "T1.\"DocDate\" AS \"fecha\", " +
                           "T1.\"U_TIMB\" AS \"timbrado\", " +
                           //detalle
                           "'1' AS \"cantidad\", " +
                           "CASE WHEN T3.\"TaxCode\"='IVA_10' THEN 10 WHEN T3.\"TaxCode\"='IVA_5' THEN 5 ELSE 0 END AS \"tasa aplica\", " +
                           "T3.\"PriceAfVAT\" * CASE WHEN T3.\"Quantity\" = 0 THEN 1 ELSE T3.\"Quantity\" END AS \"precio\", " +
                           "T1.\"Comments\", " +
                           //retencion
                           "T0.\"DocDate\" AS \"fecha ret\", " +
                           "CASE WHEN T1.\"DocCur\"='GS' THEN 'PYG' ELSE 'USD' END AS \"moneda\", " +
                           "CASE WHEN (T4.\"U_RetReta\" IS NULL OR T4.\"U_RetReta\"=0) THEN 'false' ELSE 'true' END AS \"retencionRenta\", " +
                           "CASE WHEN T4.\"U_RetReta\" IS NULL OR T4.\"U_RetReta\"=0 THEN '' ELSE 'RENTA_EMPRESARIAL_REGISTRADO.1' END AS \"conceptoRenta\", " +
                           "CASE WHEN (T4.\"U_RetIva\" IS NULL OR T4.\"U_RetIva\"=0) THEN 'false' ELSE 'true' END AS \"retencionIva\", " +
                           "CASE WHEN T4.\"U_RetIva\" IS NULL OR T4.\"U_RetIva\"=0 THEN '' ELSE 'IVA.1' END AS \"conceptoiva\", " +
                           "CASE WHEN T4.\"U_RetReta\" IS NULL  THEN '0' WHEN T3.\"TaxCode\"='IVA_EXE' THEN '0' ELSE '0.4' END AS \"rentPorcentaje\", " +
                           "'0' AS \"rentaCabezasBase\", " +
                           "'0' AS \"rentaCabezasCantidad\", " +
                           "'0' AS \"rentaToneladasBase\", " +
                           "'0' AS \"rentaToneladasCantidad\", " +
                           "CASE WHEN T3.\"TaxCode\"='IVA_5' THEN 30 ELSE 0 END AS \"ivaPorcentaje5\", " +
                           "CASE WHEN T3.\"TaxCode\"='IVA_10' THEN 70 ELSE 0 END AS \"ivaPorcentaje10\" " +
                           "FROM OVPM T0 " +
                           "INNER JOIN OPCH T1 ON T0.\"DocEntry\"=T1.\"ReceiptNum\" " +
                           "INNER JOIN OCRD T2 ON T1.\"CardCode\"=T2.\"CardCode\" " +
                           "INNER JOIN PCH1 T3 ON T1.\"DocEntry\"=T3.\"DocEntry\" " +
                           "LEFT JOIN \"@RET_CALCULO\" T4 ON T1.\"DocNum\"=T4.\"U_DocNum\" " +
                           "WHERE T0.\"DocNum\"='" + v_OP + "' " +
                           "GROUP BY T0.\"DocDate\", T2.\"U_iNatRec\", T1.\"CardName\" , T2.\"LicTradNum\", T1.\"NumAtCard\" , T1.\"GroupNum\", T1.\"Installmnt\", T1.\"DocDate\" , T1.\"U_TIMB\" , T3.\"TaxCode\"," +
                           "T1.\"DocTotal\", T1.\"Comments\", T0.\"DocDate\", T4.\"U_RetReta\",T4.\"U_RetIva\",T1.\"DocTotalFC\",T1.\"DocCur\",T3.\"PriceAfVAT\",T3.\"Quantity\"  ";

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
            if (string.IsNullOrEmpty(v_cotiMonto))
            {
                //objectJson = ConvertDataTableToArrayUSD(dt, v_cotiMonto);
                v_cotiMonto = "0";
            }
            //else
            //{
                
            //}
            objectJson = ConvertDataTableToArrayNew(dt, v_cotiMonto, v_OP);
            //grabamos el json en el escritorio
            var jsontowrite = JsonConvert.SerializeObject(objectJson, Newtonsoft.Json.Formatting.Indented);
            //creamos una carpeta en el escritorio para guardar el excel
            //string carpeta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string CarpEscr = carpeta + "\\Retenciones-Tesaka";
            //if (!Directory.Exists(CarpEscr))
            //{
            //    Directory.CreateDirectory(CarpEscr);
            //}
            //prueba
            string v_userAD = System.DirectoryServices.AccountManagement.UserPrincipal.Current.SamAccountName;
            //string path = CarpEscr + "\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            string path = null;
            //if (SAPbouiCOM.Framework.Application.SBO_Application.ClientType == SAPbouiCOM.BoClientType.ct_Desktop)
            //{
            //    path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Retenciones-Tesaka\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            //    using (var writer = new StreamWriter(path))
            //    {
            //        writer.Write(jsontowrite);
            //    }
            //}
            //else
            //{
            //    path = "C:\\Users\\" + v_userAD + "\\Desktop\\Retenciones-Tesaka\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            //    SAPbouiCOM.Framework.Application.SBO_Application.SendFileToBrowser(path);
            //}           
            txtOP.Value = "";
            path = "C:\\Users\\" + v_userAD + "\\Desktop\\Retenciones-Tesaka\\Pagos-TESAKA-FRIGORIFICO GUARANI SACI.txt";
            using (var writer = new StreamWriter(path))
            {
                writer.Write(jsontowrite);
            }
            //SAPbouiCOM.Framework.Application.SBO_Application.SendFileToBrowser(path);
            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Json exportado con éxito!!", 1, "OK");
        }

        //nuevo proceso para el formato json
        private static object[] ConvertDataTableToArrayNew(DataTable dataTable,string coti, string docu)
        {            
            string v_aux = null;
            string v_doc = null;
            int v_cont = 0;
            int v_5 = 0;
            int v_10 = 0;
            List<object> dataArray = new List<object>();
            
            //cargamos el detalle
            SAPbobsCOM.Recordset oConsulta;
            oConsulta = (SAPbobsCOM.Recordset)addOnRetencion.Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
            oConsulta.DoQuery("SELECT T1.\"NumAtCard\" FROM OVPM T0 INNER JOIN OPCH T1 ON T0.\"DocEntry\"=T1.\"ReceiptNum\" WHERE T0.\"DocNum\"='"+ docu + "' ");
            while (!oConsulta.EoF)
            {
                jsonTesaka tesaka = new jsonTesaka();
                tesaka.detalle = new List<detalle>();
                v_doc = oConsulta.Fields.Item(0).Value.ToString();
                //variable a convertir en el formato correcto              
                //formato de la fecha
                foreach (DataRow row in dataTable.Rows)
                {
                    foreach (DataRow row2 in dataTable.Rows)
                    {
                        if (v_doc == row2[4].ToString())
                        {
                            int porc5 = int.Parse(row2[21].ToString());
                            int porc10 = int.Parse(row2[22].ToString());

                            if (porc5 != 0)
                            {
                                v_5 = porc5;
                            }
                            if (porc10 != 0)
                            {
                                v_10 = porc10;
                            }
                        }
                    }
                    if (v_doc == row[4].ToString())
                    {                      
                        DateTime v_fecha = DateTime.Parse(row[0].ToString());
                        DateTime v_fechaDoc = DateTime.Parse(row[8].ToString());
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
                            //si es una o mas lineas
                            if (v_cont == 0)
                            {
                                //creamos el json cabecera
                                tesaka.atributos = new atributos()
                                {
                                    fechaCreacion = v_fecha.ToString("yyyy-MM-dd"),
                                    fechaHoraCreacion = v_fecha.ToString("yyyy-MM-dd hh:mm:ss")
                                };
                                tesaka.informado = new informado()
                                {
                                    situacion = row[1].ToString(),
                                    nombre = row[2].ToString(),
                                    ruc = ruc,
                                    dv = dv,
                                    domicilio = "",
                                    direccion = "",
                                    correoElectronico = "",
                                    tipoIdentificacion = "",
                                    identificacion = "",
                                    pais = "",
                                    telefono = ""
                                };
                                tesaka.transaccion = new transaccion()
                                {
                                    condicionCompra = row[5].ToString(),
                                    numeroComprobanteVenta = row[4].ToString(),
                                    cuotas = int.Parse(row[6].ToString()),
                                    tipoComprobante = int.Parse(row[7].ToString()),
                                    fecha = v_fechaDoc.ToString("yyyy-MM-dd"),
                                    numeroTimbrado = row[9].ToString()
                                };
                                tesaka.retencion = new retencion()
                                {
                                    fecha = v_fecha.ToString("yyyy-MM-dd"),
                                    moneda = row[15].ToString(),
                                    retencionRenta = bool.Parse("false"),
                                    conceptoRenta = "",
                                    retencionIva = bool.Parse(row[18].ToString()),
                                    conceptoIva = row[19].ToString(),
                                    rentaPorcentaje = 0,
                                    rentaCabezasBase = 0,
                                    rentaCabezasCantidad = 0,
                                    rentaToneladasBase = 0,
                                    rentaToneladasCantidad = 0,
                                    ivaPorcentaje5 = v_5,//int.Parse(row[21].ToString()),
                                    ivaPorcentaje10 = v_10//int.Parse(row[22].ToString())
                                };
                                //creamos el json detalle                                
                                tesaka.detalle.Add(new detalle()
                                {
                                    cantidad = int.Parse(row[10].ToString()),
                                    tasaAplica = row[11].ToString(),
                                    precioUnitario = double.Parse(row[12].ToString()),
                                    descripcion = row[13].ToString()
                                });
                                //v_cont++;
                            }
                            else
                            {
                                //creamos el json detalle                                
                                tesaka.detalle.Add(new detalle()
                                {
                                    cantidad = int.Parse(row[10].ToString()),
                                    tasaAplica = row[11].ToString(),
                                    precioUnitario = double.Parse(row[12].ToString()),
                                    descripcion = row[13].ToString()
                                });
                            }
                        }
                        if (retRetna.Equals("true"))
                        {
                            //si es una o mas lineas
                            if (v_cont == 0)
                            {
                                //creamos el json cabecera
                                tesaka.atributos = new atributos()
                                {
                                    fechaCreacion = v_fecha.ToString("yyyy-MM-dd"),
                                    fechaHoraCreacion = v_fecha.ToString("yyyy-MM-dd hh:mm:ss")
                                };
                                tesaka.informado = new informado()
                                {
                                    situacion = row[1].ToString(),
                                    nombre = row[2].ToString(),
                                    ruc = ruc,
                                    dv = dv,
                                    domicilio = "",
                                    direccion = "",
                                    correoElectronico = "",
                                    tipoIdentificacion = "",
                                    identificacion = "",
                                    pais = "",
                                    telefono = ""
                                };
                                tesaka.transaccion = new transaccion()
                                {
                                    condicionCompra = row[5].ToString(),
                                    numeroComprobanteVenta = row[4].ToString(),
                                    cuotas = int.Parse(row[6].ToString()),
                                    tipoComprobante = int.Parse(row[7].ToString()),
                                    fecha = v_fechaDoc.ToString("yyyy-MM-dd"),
                                    numeroTimbrado = row[9].ToString()
                                };
                                tesaka.retencion = new retencion()
                                {
                                    fecha = v_fecha.ToString("yyyy-MM-dd"),
                                    moneda = row[15].ToString(),
                                    retencionRenta = bool.Parse(row[16].ToString()),
                                    conceptoRenta = row[17].ToString(),
                                    retencionIva = bool.Parse("false"),
                                    conceptoIva = "",
                                    rentaPorcentaje = double.Parse(row[20].ToString()),
                                    rentaCabezasBase = 0,
                                    rentaCabezasCantidad = 0,
                                    rentaToneladasBase = 0,
                                    rentaToneladasCantidad = 0,
                                    ivaPorcentaje5 = 0,
                                    ivaPorcentaje10 = 0
                                };
                                //creamos el json detalle                                
                                tesaka.detalle.Add(new detalle()
                                {
                                    cantidad = int.Parse(row[10].ToString()),
                                    tasaAplica = row[11].ToString(),
                                    precioUnitario = double.Parse(row[12].ToString()),
                                    descripcion = row[13].ToString()
                                });
                                //v_cont++;
                            }
                            else
                            {
                                //creamos el json detalle                                
                                tesaka.detalle.Add(new detalle()
                                {
                                    cantidad = int.Parse(row[10].ToString()),
                                    tasaAplica = row[11].ToString(),
                                    precioUnitario = double.Parse(row[12].ToString()),
                                    descripcion = row[13].ToString()
                                });
                            }
                        }                        
                    }                  
                }
                dataArray.Add(tesaka);
                oConsulta.MoveNext();
            }
            return dataArray.ToArray();          
        }

        private static object[] ConvertDataTableToArrayUSDNew(DataTable dataTable, string coti, string docu)
        {
            string v_aux = null;
            string v_doc = null;
            int v_cont = 0;
            int v_5 = 0;
            int v_10 = 0;
            List<object> dataArray = new List<object>();

            //cargamos el detalle
            SAPbobsCOM.Recordset oConsulta;
            oConsulta = (SAPbobsCOM.Recordset)addOnRetencion.Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
            oConsulta.DoQuery("SELECT T1.\"NumAtCard\" FROM OVPM T0 INNER JOIN OPCH T1 ON T0.\"DocEntry\"=T1.\"ReceiptNum\" WHERE T0.\"DocNum\"='" + docu + "' ");
            while (!oConsulta.EoF)
            {
                modelo.jsonTesaka tesaka = new modelo.jsonTesaka();
                tesaka.detalle = new List<modelo.detalle>();
                v_doc = oConsulta.Fields.Item(0).Value.ToString();
                //variable a convertir en el formato correcto              
                //formato de la fecha
                foreach (DataRow row in dataTable.Rows)
                {
                    foreach (DataRow row2 in dataTable.Rows)
                    {
                        if (v_doc == row2[4].ToString())
                        {
                            int porc5 = int.Parse(row2[21].ToString());
                            int porc10 = int.Parse(row2[22].ToString());

                            if (porc5 != 0)
                            {
                                v_5 = porc5;
                            }
                            if (porc10 != 0)
                            {
                                v_10 = porc10;
                            }
                        }
                    }
                    if (v_doc == row[4].ToString())
                    {
                        DateTime v_fecha = DateTime.Parse(row[0].ToString());
                        DateTime v_fechaDoc = DateTime.Parse(row[8].ToString());
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
                            //si es una o mas lineas
                            if (v_cont == 0)
                            {
                                //creamos el json cabecera
                                tesaka.atributos = new modelo.atributos()
                                {
                                    fechaCreacion = v_fecha.ToString("yyyy-MM-dd"),
                                    fechaHoraCreacion = v_fecha.ToString("yyyy-MM-dd hh:mm:ss")
                                };
                                tesaka.informado = new modelo.informado()
                                {
                                    situacion = row[1].ToString(),
                                    nombre = row[2].ToString(),
                                    ruc = ruc,
                                    dv = dv,
                                    domicilio = "",
                                    direccion = "",
                                    correoElectronico = "",
                                    tipoIdentificacion = "",
                                    identificacion = "",
                                    pais = "",
                                    telefono = ""
                                };
                                tesaka.transaccion = new modelo.transaccion()
                                {
                                    condicionCompra = row[5].ToString(),
                                    numeroComprobanteVenta = row[4].ToString(),
                                    cuotas = int.Parse(row[6].ToString()),
                                    tipoComprobante = int.Parse(row[7].ToString()),
                                    fecha = v_fechaDoc.ToString("yyyy-MM-dd"),
                                    numeroTimbrado = row[9].ToString()
                                };
                                tesaka.retencion= new modelo.retencion()
                                {
                                    fecha = v_fecha.ToString("yyyy-MM-dd"),
                                    moneda = row[15].ToString(),
                                    tipoCambio = int.Parse(coti),
                                    retencionRenta = bool.Parse("false"),
                                    conceptoRenta = "",
                                    retencionIva = bool.Parse(row[18].ToString()),
                                    conceptoIva = row[19].ToString(),
                                    rentaPorcentaje = 0,
                                    rentaCabezasBase = 0,
                                    rentaCabezasCantidad = 0,
                                    rentaToneladasBase = 0,
                                    rentaToneladasCantidad = 0,
                                    ivaPorcentaje5 = v_5,//int.Parse(row[21].ToString()),
                                    ivaPorcentaje10 = v_10//int.Parse(row[22].ToString())
                                };
                                //creamos el json detalle                                
                                tesaka.detalle.Add(new modelo.detalle()
                                {
                                    cantidad = int.Parse(row[10].ToString()),
                                    tasaAplica = row[11].ToString(),
                                    precioUnitario = double.Parse(row[12].ToString()),
                                    descripcion = row[13].ToString()
                                });
                                //v_cont++;
                            }
                            else
                            {
                                //creamos el json detalle                                
                                tesaka.detalle.Add(new modelo.detalle()
                                {
                                    cantidad = int.Parse(row[10].ToString()),
                                    tasaAplica = row[11].ToString(),
                                    precioUnitario = double.Parse(row[12].ToString()),
                                    descripcion = row[13].ToString()
                                });
                            }
                        }
                        if (retRetna.Equals("true"))
                        {
                            //si es una o mas lineas
                            if (v_cont == 0)
                            {
                                //creamos el json cabecera
                                tesaka.atributos = new modelo.atributos()
                                {
                                    fechaCreacion = v_fecha.ToString("yyyy-MM-dd"),
                                    fechaHoraCreacion = v_fecha.ToString("yyyy-MM-dd hh:mm:ss")
                                };
                                tesaka.informado = new modelo.informado()
                                {
                                    situacion = row[1].ToString(),
                                    nombre = row[2].ToString(),
                                    ruc = ruc,
                                    dv = dv,
                                    domicilio = "",
                                    direccion = "",
                                    correoElectronico = "",
                                    tipoIdentificacion = "",
                                    identificacion = "",
                                    pais = "",
                                    telefono = ""
                                };
                                tesaka.transaccion = new modelo.transaccion()
                                {
                                    condicionCompra = row[5].ToString(),
                                    numeroComprobanteVenta = row[4].ToString(),
                                    cuotas = int.Parse(row[6].ToString()),
                                    tipoComprobante = int.Parse(row[7].ToString()),
                                    fecha = v_fechaDoc.ToString("yyyy-MM-dd"),
                                    numeroTimbrado = row[9].ToString()
                                };
                                tesaka.retencion = new modelo.retencion()
                                {
                                    fecha = v_fecha.ToString("yyyy-MM-dd"),
                                    moneda = row[15].ToString(),
                                    tipoCambio = int.Parse(coti),
                                    retencionRenta = bool.Parse(row[16].ToString()),
                                    conceptoRenta = row[17].ToString(),
                                    retencionIva = bool.Parse("false"),
                                    conceptoIva = "",
                                    rentaPorcentaje = double.Parse(row[20].ToString()),
                                    rentaCabezasBase = 0,
                                    rentaCabezasCantidad = 0,
                                    rentaToneladasBase = 0,
                                    rentaToneladasCantidad = 0,
                                    ivaPorcentaje5 = 0,
                                    ivaPorcentaje10 = 0
                                };
                                //creamos el json detalle                                
                                tesaka.detalle.Add(new modelo.detalle()
                                {
                                    cantidad = int.Parse(row[10].ToString()),
                                    tasaAplica = row[11].ToString(),
                                    precioUnitario = double.Parse(row[12].ToString()),
                                    descripcion = row[13].ToString()
                                });
                                //v_cont++;
                            }
                            else
                            {
                                //creamos el json detalle                                
                                tesaka.detalle.Add(new modelo.detalle()
                                {
                                    cantidad = int.Parse(row[10].ToString()),
                                    tasaAplica = row[11].ToString(),
                                    precioUnitario = double.Parse(row[12].ToString()),
                                    descripcion = row[13].ToString()
                                });
                            }
                        }
                    }
                }
                dataArray.Add(tesaka);
                oConsulta.MoveNext();
            }
            return dataArray.ToArray();
        }


        //MODELOS PARA EL JSON
        #region MODELO PARA EL JSON
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

 
        #endregion



    }
}