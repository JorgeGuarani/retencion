using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using System.Reflection;
using Newtonsoft.Json;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Web.UI.WebControls;
using System.Linq;

namespace addOnRetencion
{
    class Program
    {
        public static string v_docnum = null;
        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbouiCOM.Form formPago = null;
        public static double TotalRetIva = 0;
        public static double TotalRetRenta = 0;
        public static decimal global_coti = 0;
        public static string monedaPago = "GS";
        public static double totalFact = 0;
        public static double totalcuota = 0;


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            try
            {
                //oCompany = Menu.sbo;
                SAPbouiCOM.Framework.Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new SAPbouiCOM.Framework.Application();
                }
                else
                {
                    oApp = new SAPbouiCOM.Framework.Application(args[0]);
                }
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                SAPbouiCOM.Framework.Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SAPbouiCOM.Framework.Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);

                oApp.Run();

                //conectamos a la base de datos
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("AddOn Retenciones conectado!!", SAPbouiCOM.BoMessageTime.bmt_Short, false);

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                #region FACTURA PROVEEDOR
                if (pVal.FormTypeEx == "141")
                {                   
                    //agarramos el ID del form
                    SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);

                    //agarrar el total de la factura proveedor
                    if(pVal.ItemUID=="171" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.EditText ototal = (SAPbouiCOM.EditText)form.Items.Item("29").Specific;
                        SAPbouiCOM.EditText oCuota = (SAPbouiCOM.EditText)form.Items.Item("170").Specific;
                        string v_total = ototal.Value;
                        v_total = string.Concat(v_total.Where(c => Char.IsDigit(c)));
                        totalFact = double.Parse(v_total);
                        //totalcuota = double.Parse(oCuota.Value);
                        //totalcuota = Math.Round((totalFact / totalcuota),0);
                    }
                    
                    //FUNCIONES
                    #region AGARRAR EL DOCNUM
                    if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == true)
                    {
                        //agarramos el docnum del documento
                        SAPbouiCOM.EditText oDocNum = (SAPbouiCOM.EditText)form.Items.Item("8").Specific;
                        v_docnum = oDocNum.Value;
                    }
                    #endregion
                 
                    #region CALCULO AL CREAR EL DOCUMENTO
                    if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemUID == "1" && pVal.FormTypeEx == "141" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.FormMode == 3)
                    {
                        try
                        {
                            string v_iva = null;
                            string v_renta = null;
                            decimal v_total = 0;
                            decimal v_RetRenta = 0;
                            decimal v_RetIva = 0;
                            string v_Moneda = null;
                            string v_DocEntry = null;
                            int v_cuotas = 0;

                            //consultamos el tipo de retención
                            SAPbobsCOM.Recordset otipoRet;
                            otipoRet = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                            otipoRet.DoQuery("SELECT \"U_RetIVA\",\"U_RetRenta\",CASE WHEN \"DocCur\"='GS' THEN \"DocTotal\" ELSE \"DocTotalFC\" END,\"DocCur\",\"DocEntry\",\"Installmnt\" FROM OPCH WHERE \"DocNum\"='" + v_docnum + "' ");
                            while (!otipoRet.EoF)
                            {
                                v_iva = otipoRet.Fields.Item(0).Value.ToString();
                                v_renta = otipoRet.Fields.Item(1).Value.ToString();
                                v_total = decimal.Parse(otipoRet.Fields.Item(2).Value.ToString());
                                v_Moneda = otipoRet.Fields.Item(3).Value.ToString();
                                v_DocEntry = otipoRet.Fields.Item(4).Value.ToString();
                                v_cuotas = int.Parse(otipoRet.Fields.Item(5).Value.ToString());
                                otipoRet.MoveNext();
                            }

                            //impuesto
                            SAPbobsCOM.Recordset oImpuesto;
                            oImpuesto = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                            oImpuesto.DoQuery("SELECT \"TaxCode\" FROM PCH1 WHERE \"DocEntry\"='"+ v_DocEntry + "' ");
                            string v_taxcode = oImpuesto.Fields.Item(0).Value.ToString();

                            //calculamos la retencion
                            if (v_iva.Equals("SI"))
                            {
                                if (v_Moneda.Equals("USD"))
                                {
                                    if (v_taxcode.Equals("IVA_5"))
                                    {
                                        v_RetIva = decimal.Round(((v_total / 21) * 30) / 100, 2);
                                    }
                                    if (v_taxcode.Equals("IVA_10"))
                                    {
                                        v_RetIva = decimal.Round(((v_total / 11) * 70) / 100, 2);
                                    }

                                }
                                else
                                {
                                    if (v_taxcode.Equals("IVA_5"))
                                    {
                                        v_RetIva = Math.Round(((v_total / 21) * 30) / 100);
                                    }
                                    if (v_taxcode.Equals("IVA_10"))
                                    {
                                        v_RetIva = Math.Round(((v_total / 11) * 70) / 100);
                                    }
                                }

                            }
                            //en caso de ser solo IVA
                            if (v_renta.Equals("SI"))
                            {
                                if (v_Moneda.Equals("USD"))
                                {
                                    v_RetRenta = decimal.Round((v_total / decimal.Parse("1,05")) * decimal.Parse("0,004"), 2);
                                }
                                else
                                {
                                    v_RetRenta = Math.Round((v_total / decimal.Parse("1,05")) * decimal.Parse("0,004"));
                                }
                            }

                            //guardamos en la tabla de calculo de retención
                            SAPbobsCOM.Recordset codMax;
                            codMax = (Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                            codMax.DoQuery("select CASE WHEN MAX(\"DocEntry\")=0 THEN 1 ELSE  MAX(\"DocEntry\") END from \"@RET_CALCULO\" ");
                            int MaxCod = int.Parse(codMax.Fields.Item(0).Value.ToString()) + 1;


                            string v_NumCuota = null;
                            decimal v_montoCuota = 0;
                            //consultamos si es por cuotas
                            if (v_cuotas > 1)
                            {
                                SAPbobsCOM.Recordset oCuotas;
                                oCuotas = (Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                                oCuotas.DoQuery("SELECT \"InstlmntID\",\"InsTotal\" FROM PCH6 WHERE \"DocEntry\"='" + v_DocEntry + "' ");
                                //recorremos
                                while (!oCuotas.EoF)
                                {
                                    string v_cuota = " de " + v_cuotas;
                                    v_NumCuota = oCuotas.Fields.Item(0).Value.ToString();
                                    v_cuota = v_NumCuota + v_cuota;
                                    v_montoCuota = decimal.Parse(oCuotas.Fields.Item(1).Value.ToString());
                                    //calculamos la retencion
                                    if (v_iva.Equals("SI"))
                                    {
                                        if (v_Moneda.Equals("USD"))
                                        {
                                            v_RetIva = decimal.Round(((v_montoCuota / 21) * 30) / 100, 2);
                                        }
                                        else
                                        {
                                            v_RetIva = Math.Round(((v_montoCuota / 21) * 30) / 100);
                                        }

                                    }
                                    //en caso de ser solo IVA
                                    if (v_renta.Equals("SI"))
                                    {
                                        if (v_Moneda.Equals("USD"))
                                        {
                                            v_RetRenta = decimal.Round((v_montoCuota / decimal.Parse("1,05")) * decimal.Parse("0,004"), 2);
                                        }
                                        else
                                        {
                                            v_RetRenta = Math.Round((v_montoCuota / decimal.Parse("1,05")) * decimal.Parse("0,004"));
                                        }
                                    }

                                    SAPbobsCOM.Recordset oGrabar;
                                    oGrabar = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    oGrabar.DoQuery("INSERT INTO \"@RET_CALCULO\" (\"Code\",\"DocEntry\",\"U_DocNum\",\"U_Estado\",\"U_RetIva\",\"U_RetReta\",\"U_Plazo\",\"U_Moneda\") VALUES ('" + MaxCod + "','" + MaxCod + "','" + v_docnum + "','P','" + v_RetIva.ToString().Replace(",", ".") + "', '" + v_RetRenta.ToString().Replace(",", ".") + "', '" + v_cuota + "','" + v_Moneda + "') ");

                                    MaxCod++;
                                    oCuotas.MoveNext();
                                }
                            }
                            else
                            {
                                SAPbobsCOM.Recordset oGrabar;
                                oGrabar = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                                oGrabar.DoQuery("INSERT INTO \"@RET_CALCULO\" (\"Code\",\"DocEntry\",\"U_DocNum\",\"U_Estado\",\"U_RetIva\",\"U_RetReta\",\"U_Plazo\",\"U_Moneda\") VALUES ('" + MaxCod + "','" + MaxCod + "','" + v_docnum + "','P','" + v_RetIva.ToString().Replace(",", ".") + "', '" + v_RetRenta.ToString().Replace(",", ".") + "','1 de 1','" + v_Moneda + "') ");
                            }
                        }
                        catch(Exception e)
                        {
                           // System.Windows.Forms.MessageBox.Show(e.Message);
                            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.ToString(), 1, "OK");
                        }                       

                    }
                    #endregion

                    
                }
                #endregion

                #region FORM PAGOS
                if (pVal.FormTypeEx == "426")
                {
                    //agarramos el ID del form
                    SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);
                    formPago = form;

                    #region CREAR BOTON
                    //agregar boton para exportar txt
                    if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);
                        Item oItem;
                        SAPbouiCOM.Button oButton;
                        oItem = oForm.Items.Add("btnJson", BoFormItemTypes.it_BUTTON);
                        //Inicializando el objeto boton con la referencia del objeto item
                        oButton = (SAPbouiCOM.Button)oItem.Specific;
                        //Agregando propiedades al boton
                        oButton.Caption = "Exportar Json";
                        //agregando posicio del boton
                        oItem.Top = oForm.Height - (oItem.Height + 10);
                        oItem.Left = (oItem.Width + 20) + 63;
                    }

                    //agregar boton para recalcular
                    if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);
                        Item oItem;
                        SAPbouiCOM.Button oButton;
                        oItem = oForm.Items.Add("btnRecal", BoFormItemTypes.it_BUTTON);
                        //Inicializando el objeto boton con la referencia del objeto item
                        oButton = (SAPbouiCOM.Button)oItem.Specific;
                        //Agregando propiedades al boton
                        oButton.Caption = "Recalcular";
                        //agregando posicio del boton
                        oItem.Top = oForm.Height - (oItem.Height + 10);
                        oItem.Left = (oItem.Width + 20) + 134;
                    }

                    //agregar boton para importar
                    if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);
                        Item oItem;
                        SAPbouiCOM.Button oButton;
                        oItem = oForm.Items.Add("btnImp", BoFormItemTypes.it_BUTTON);
                        //Inicializando el objeto boton con la referencia del objeto item
                        oButton = (SAPbouiCOM.Button)oItem.Specific;
                        //Agregando propiedades al boton
                        oButton.Caption = "Importar Json";
                        //agregando posicio del boton
                        oItem.Top = oForm.Height - (oItem.Height + 10);
                        oItem.Left = (oItem.Width + 20) + 204;
                    }

                    //para exportar el Json
                    if (pVal.ItemUID=="btnJson" && pVal.BeforeAction==false && pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                    {
                        //agarramos las variables del pago
                        EditText oDocNum = (SAPbouiCOM.EditText)form.Items.Item("3").Specific;
                        string v_DocNum = oDocNum.Value;
                        EditText oFecha = (SAPbouiCOM.EditText)form.Items.Item("10").Specific;
                        EditText oCoti = (SAPbouiCOM.EditText)form.Items.Item("41").Specific;
                        string v_coti = oCoti.Value;
                        string v_fecha = oFecha.Value;
                        string v_cotiMonto = null;
                        if (v_coti.Equals("USD"))
                        {
                            EditText oCotizacion = (SAPbouiCOM.EditText)form.Items.Item("21").Specific;
                            v_cotiMonto = oCotizacion.Value;
                            double cotimonto_v = Math.Round(double.Parse(v_cotiMonto.Replace(".",",")));
                            v_cotiMonto = cotimonto_v.ToString();
                        }
                        //abrimos el form para exportar
                        SAPbouiCOM.Framework.Application.SBO_Application.ActivateMenuItem("addOnRetencion.Form1");
                        SAPbouiCOM.Form formJson = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                        //agarramos los campos y mandamos las variables del pago
                        EditText oFechaJson = (EditText)formJson.Items.Item("Item_3").Specific;
                        EditText oDocJson = (EditText)formJson.Items.Item("Item_4").Specific;
                        SAPbouiCOM.Button oBtnJson = (SAPbouiCOM.Button)formJson.Items.Item("Item_5").Specific;
                        SAPbouiCOM.Button oBtnCancel = (SAPbouiCOM.Button)formJson.Items.Item("Item_6").Specific;
                        EditText ocotiJson = (EditText)formJson.Items.Item("Item_7").Specific;
                        oFechaJson.Value = v_fecha;
                        oDocJson.Value = v_DocNum;
                        ocotiJson.Value = v_cotiMonto;
                        oBtnJson.Item.Click();
                        oBtnCancel.Item.Click();
                    }

                    //para recalcular los valores de las retenciones
                    if(pVal.ItemUID=="btnRecal" && pVal.BeforeAction==false && pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)formPago.Items.Item("20").Specific;
                        int v_can = oMatrix.RowCount;
                        int v_fila = 1;
                        while (v_fila <= v_can)
                        {
                            SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("10000127").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oDocNum = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oIVA = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorIva").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oRENTA = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorRenta").Cells.Item(v_fila).Specific;
                            bool v_check = oCheck.Checked;
                            string v_DocNum = null;
                            if (v_check == true)
                            {
                                v_DocNum = oDocNum.Value;
                                try
                                {
                                    recalcular(v_DocNum);
                                }
                                catch (Exception e)
                                {
                                   // System.Windows.Forms.MessageBox.Show(e.Message);
                                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.ToString(), 1, "OK");
                                }

                                SAPbobsCOM.Recordset oDatos;
                                oDatos = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                                oDatos.DoQuery("SELECT \"U_RetIva\",\"U_RetReta\" FROM \"@RET_CALCULO\" WHERE \"U_DocNum\"='"+ v_DocNum + "' ");
                                string v_iva = oDatos.Fields.Item(0).Value.ToString();
                                string v_renta = oDatos.Fields.Item(1).Value.ToString();
                                double iva_v = double.Parse(v_iva.Replace(",", "."));
                                double renta_v = double.Parse(v_renta.Replace(",", "."));
                                oIVA.Value = v_iva.Replace(",", ".");
                                oRENTA.Value = v_renta.Replace(",", ".");
                            }

                            v_fila++;
                        }
                    }

                    //para importar el json
                    if (pVal.ItemUID == "btnImp" && pVal.BeforeAction == false && pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                    {
                        //buscamos el archivo txt
                        Thread t = new Thread(() =>
                        {
                            OpenFileDialog openFileDialog = new OpenFileDialog();
                            DialogResult dr = openFileDialog.ShowDialog();
                            if (dr == DialogResult.OK)
                            {
                                string fileName = openFileDialog.FileName;
                                leerTXT(fileName);
                                // FILE.Value = fileName;
                                //SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(fileName);
                            }
                        });          // Kick off a new thread
                        t.IsBackground = true;
                        t.SetApartmentState(ApartmentState.STA);
                        t.Start();
                    }
                        #endregion

                    #region GRILLA
                        if (pVal.ItemUID == "20" && pVal.EventType==BoEventTypes.et_CLICK && pVal.BeforeAction==false)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item("20").Specific;

                        //consultamos la cotizacion del día
                        string v_fecha = DateTime.Now.ToString("yyyyMMdd");
                        SAPbobsCOM.Recordset oCotizacion;
                        oCotizacion = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oCotizacion.DoQuery("SELECT \"Rate\" FROM ORTT WHERE \"Currency\"='USD' AND \"RateDate\"='"+ v_fecha + "'");
                        decimal v_coti = decimal.Parse(oCotizacion.Fields.Item(0).Value.ToString());
                        int  v_cotizacion = 0;
                        string[] coma = v_coti.ToString().Split(',');
                        int c = 0;
                        foreach(string valor in coma)
                        {
                            if (c == 0)
                            {
                                v_cotizacion = int.Parse(valor);
                            }

                            if (c == 1)
                            {
                                int valordecimal = int.Parse(valor);
                                if (valordecimal > 50)
                                {
                                    v_cotizacion = v_cotizacion + 1;
                                }
                               
                            }
                            c++;
                        }
                        global_coti = v_coti;
                        //cantidad de fila de la matriz
                        int v_filaCant = oMatrix.RowCount;
                        int v_fila = 1;
                        //recorremos la matrix
                            SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("10000127").Cells.Item(pVal.Row).Specific;
                            bool v_check = oCheck.Checked;
                            if(v_check == true)
                            {
                                SAPbouiCOM.EditText oDocNum = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific; 
                                string v_DocNum = oDocNum.Value;
                                SAPbouiCOM.EditText oPlazo = (SAPbouiCOM.EditText)oMatrix.Columns.Item("71").Cells.Item(pVal.Row).Specific;
                                string v_Plazo = oPlazo.Value;
                                SAPbouiCOM.EditText oRetIva = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorIva").Cells.Item(pVal.Row).Specific;
                                SAPbouiCOM.EditText oRetRenta = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorRenta").Cells.Item(pVal.Row).Specific;

                                SAPbobsCOM.Recordset oDatosRet;
                                oDatosRet = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                                oDatosRet.DoQuery("SELECT \"U_DocNum\",\"U_RetIva\",\"U_RetReta\",\"U_Plazo\",\"U_Moneda\" FROM \"@RET_CALCULO\" WHERE \"U_DocNum\"='" + v_DocNum + "' AND \"U_Plazo\"='"+v_Plazo+"'  ");
                                while (!oDatosRet.EoF)
                                {
                                    string v_Moneda = oDatosRet.Fields.Item(4).Value.ToString();
                                    //consultamos el tipo de moneda
                                    if (v_Moneda.Equals("GS"))
                                    {
                                        if (monedaPago.Equals("GS"))
                                        {
                                            oRetIva.Value = oDatosRet.Fields.Item(1).Value.ToString();
                                            oRetRenta.Value = oDatosRet.Fields.Item(2).Value.ToString();
                                        }
                                        else
                                        {
                                            decimal v_ivaRet = decimal.Parse(oDatosRet.Fields.Item(1).Value.ToString());
                                            decimal v_rentaRet = decimal.Parse(oDatosRet.Fields.Item(2).Value.ToString());

                                            oRetIva.Value = (decimal.Round(v_ivaRet / v_cotizacion, 0)).ToString().Replace(",", ".");
                                            oRetRenta.Value = (decimal.Round(v_rentaRet / v_cotizacion, 0)).ToString().Replace(",", ".");
                                        }                                       
                                    }
                                    else
                                    {
                                        if (monedaPago.Equals("GS"))
                                         {
                                            decimal v_ivaRet = decimal.Parse(oDatosRet.Fields.Item(1).Value.ToString());
                                            decimal v_rentaRet = decimal.Parse(oDatosRet.Fields.Item(2).Value.ToString());

                                            oRetIva.Value = (decimal.Round(v_ivaRet * v_cotizacion, 0)).ToString().Replace(",", ".");
                                            oRetRenta.Value = (decimal.Round(v_rentaRet * v_cotizacion, 0)).ToString().Replace(",", ".");
                                        }
                                        else
                                        {
                                            oRetIva.Value = oDatosRet.Fields.Item(1).Value.ToString().Replace(",", ".");
                                            oRetRenta.Value = oDatosRet.Fields.Item(2).Value.ToString().Replace(",", ".");
                                        }                                      
                                    }
                                   
                                    oDatosRet.MoveNext();
                                }
                            }                       
                    }
                    #endregion

                    #region EXPORTAR JSON
                    if(pVal.ItemUID=="btnJson" && pVal.BeforeAction==false && pVal.EventType == BoEventTypes.et_CLICK)
                    {
                        SAPbouiCOM.EditText oDoc = (SAPbouiCOM.EditText)form.Items.Item("3").Specific;
                        string v_doc = oDoc.Value;
                    }
                    #endregion

                }
                #endregion

                #region MEDIO DE PAGO
                if (pVal.FormTypeEx == "196")
                {
                    //agarramos el ID del form
                    SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);

                    #region MONEDA DE PAGO
                    if (pVal.ItemUID == "8" && pVal.EventType==BoEventTypes.et_COMBO_SELECT && pVal.BeforeAction==false)
                    {
                        SAPbouiCOM.ComboBox oMonepago = (SAPbouiCOM.ComboBox)form.Items.Item("8").Specific;
                        monedaPago = oMonepago.Selected.Value;
                    }
                    #endregion

                    #region RETENCION MEDIO DE PAGO
                    if (pVal.ItemUID == "112" && pVal.EventType==BoEventTypes.et_COMBO_SELECT && pVal.BeforeAction==false)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item("112").Specific;
                        SAPbouiCOM.ComboBox oSeleccion = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("41").Cells.Item(1).Specific;
                        SAPbouiCOM.EditText oCuenta = (SAPbouiCOM.EditText)oMatrix.Columns.Item("67").Cells.Item(1).Specific;
                        SAPbouiCOM.EditText oMonto = (SAPbouiCOM.EditText)oMatrix.Columns.Item("46").Cells.Item(1).Specific;
                        string v_TipoRet= oSeleccion.Selected.Value.ToString();
                        //si es retencion IVA
                        if (v_TipoRet.Equals("2"))
                        {
                            oCuenta.Value = "2.01.04.001.002";
                            oMonto.Value = TotalRetIva.ToString();
                        }
                        //si es retencion renta
                        if (v_TipoRet.Equals("7"))
                        {
                            oCuenta.Value = "2.01.04.001.003";
                            oMonto.Value = TotalRetRenta.ToString();
                        }

                    }
                    #endregion

                    #region CHEQUES
                    if (pVal.ItemUID == "3" && pVal.EventType==BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction==false)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)formPago.Items.Item("20").Specific;
                        SAPbouiCOM.Matrix oMatrixMP = (SAPbouiCOM.Matrix)form.Items.Item("28").Specific;
                        SAPbouiCOM.ComboBox oTipoMoneda = (SAPbouiCOM.ComboBox)form.Items.Item("8").Specific;
                        SAPbouiCOM.EditText oCoti = (SAPbouiCOM.EditText)form.Items.Item("95").Specific;
                        double v_coti = decimal.ToDouble(global_coti); 
                        string v_tipoMoneda = oTipoMoneda.Selected.Value;
                        string v_tipoMonedaV = oTipoMoneda.Selected.Description;
                        int v_rows = oMatrix.RowCount;
                        int v_fila = 1;
                        int v_filaCheque = 1;
                        //instanciamos el array
                        List<object> dataArray = new List<object>();
                        while (v_fila <= v_rows)
                        {
                            //variables de la tabla de pagos
                            SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("10000127").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oValorIva = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorIva").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oValorRenta = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorRenta").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oTotalLinea = (SAPbouiCOM.EditText)oMatrix.Columns.Item("24").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.ComboBox oTipoPago = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_TipoPago").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oVto = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_fecha_rete").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oProveedor = (SAPbouiCOM.EditText)oMatrix.Columns.Item("40").Cells.Item(v_fila).Specific;
                            //variable de la tabla de medios de pago
                            //SAPbouiCOM.EditText oImportaMP = (SAPbouiCOM.EditText)oMatrixMP.Columns.Item("7").Cells.Item(v_filaCheque).Specific;
                            //SAPbouiCOM.EditText oCliMP = (SAPbouiCOM.EditText)oMatrixMP.Columns.Item("10").Cells.Item(v_filaCheque).Specific;
                            string v_proveedor = oProveedor.Value;
                            string v_fecha = oVto.Value;
                            string v_tipoPago = oTipoPago.Value.ToString();
                            bool v_check = oCheck.Checked;
                            //consultamos si esta checkeado
                            if (v_check == true)
                            {
                                //si la opcion de pago es cheque
                                if (v_tipoPago.Equals("Cheque"))
                                {
                                    //el tipo de moneda del documento
                                     if (v_tipoMoneda.Equals("GS") || v_tipoMonedaV.Contains("GS"))
                                    {
                                        string v_totalLinea = Regex.Replace(oTotalLinea.Value, @"([A-Z])", string.Empty);
                                        string v_totalLineaMone = Regex.Replace(oTotalLinea.Value, @"([.,0-9])", string.Empty);
                                        double v_totalNew = 0;
                                        if (v_totalLineaMone.Contains("USD"))
                                        {
                                            v_totalNew = Math.Round(double.Parse(v_totalLinea) * v_coti);
                                        }
                                        else
                                        {
                                            v_totalNew = double.Parse(v_totalLinea);
                                        }
                                        string v_retIva = oValorIva.Value.Replace(".", ",");
                                        string v_retRenta = oValorRenta.Value.Replace(".", ",");
                                        //realizamos el calculo
                                        double v_newValor = v_totalNew - (double.Parse(v_retIva) + double.Parse(v_retRenta));
                                        //oImportaMP.Value = v_newValor.ToString();
                                        arrayCheque(dataArray, v_fecha, v_newValor, v_proveedor);
                                        v_filaCheque++;
                                        v_newValor = 0;
                                    }
                                    if (v_tipoMoneda.Equals("USD"))
                                    {
                                        string v_totalLinea = Regex.Replace(oTotalLinea.Value, @"([A-Z])", string.Empty);
                                        double d_totalLinea = double.Parse(v_totalLinea);
                                        string v_totalLineaMone = Regex.Replace(oTotalLinea.Value, @"([.,0-9])", string.Empty);
                                        double d_retIva = 0;
                                        double d_retRenta = 0;
                                        if (v_totalLineaMone.Contains("USD"))
                                        {
                                            string v_retIva = oValorIva.Value.Replace(".", ",");
                                            string v_retRenta = oValorRenta.Value.Replace(".", ",");

                                            d_retIva = double.Parse(v_retIva);
                                            d_retRenta = double.Parse(v_retRenta);
                                        }
                                        else
                                        {
                                            string v_retIva = oValorIva.Value.Replace(".", ",");
                                            d_retIva = Math.Round((double.Parse(v_retIva) / v_coti), 2);
                                            string v_retRenta = oValorRenta.Value.Replace(".", ",");
                                            d_retRenta = Math.Round((double.Parse(v_retRenta) / v_coti), 2);
                                        }
                                        
                                        //realizamos el calculo
                                        double v_newValor = d_totalLinea - (d_retIva + d_retRenta);
                                        //oImportaMP.Value = v_newValor.ToString();
                                        arrayCheque(dataArray, v_fecha, v_newValor, v_proveedor);
                                        v_filaCheque++;
                                        v_newValor = 0;
                                    }
                                    
                                }
                            }
                            v_fila++;
                        }
                        //cargamos los cheques
                        cargaCheques(dataArray, form);

                    }
                    #endregion

                    #region TRANSFERENCIA
                    if (pVal.ItemUID == "4" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)formPago.Items.Item("20").Specific;
                        //varibale de la seccion de transferencia
                        SAPbouiCOM.EditText oTransferencia = (SAPbouiCOM.EditText)form.Items.Item("34").Specific;
                        SAPbouiCOM.ComboBox oTipoMoneda = (SAPbouiCOM.ComboBox)form.Items.Item("8").Specific;
                        SAPbouiCOM.EditText oFechaTrans = (SAPbouiCOM.EditText)form.Items.Item("44").Specific;
                        double v_coti = decimal.ToDouble(global_coti);
                        string v_tipoMoneda = oTipoMoneda.Selected.Value;
                        string v_tipoMonedaV = oTipoMoneda.Selected.Description;
                        double v_newValor = 0;
                        string v_frete = null;
                        int v_rows = oMatrix.RowCount;
                        int v_fila = 1;
                        while (v_fila <= v_rows)
                        {
                            //variables de la tabla de pagos
                            SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("10000127").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oValorIva = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorIva").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oValorRenta = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorRenta").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oTotalLinea = (SAPbouiCOM.EditText)oMatrix.Columns.Item("24").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.ComboBox oTipoPago = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_TipoPago").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oFtrans = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_fecha_rete").Cells.Item(v_fila).Specific;
                            
                            string v_tipoPago = oTipoPago.Value.ToString();
                            bool v_check = oCheck.Checked;
                            //consultamos si esta checkeado
                            if (v_check == true)
                            {
                                v_frete = oFtrans.Value;
                                //si la opcion de pago es cheque
                                if (v_tipoPago.Equals("Transferencia"))
                                {
                                    double d_retIva = 0;
                                    double d_retRenta = 0;
                                    //el tipo de moneda del documento
                                    if (v_tipoMoneda.Equals("GS") || v_tipoMonedaV.Contains("GS"))
                                    {
                                        string v_totalLinea = Regex.Replace(oTotalLinea.Value, @"([A-Z])", string.Empty);
                                        string v_totalLineaMone = Regex.Replace(oTotalLinea.Value, @"([.,0-9])", string.Empty);
                                        double v_totalNew = 0;
                                        if (v_totalLineaMone.Contains("USD"))
                                        {
                                            v_totalNew = Math.Round(double.Parse(v_totalLinea) * v_coti);
                                        }
                                        else
                                        {
                                            v_totalNew = double.Parse(v_totalLinea);
                                        }
                                        string v_retIva = oValorIva.Value.Replace(".", ",");
                                        string v_retRenta = oValorRenta.Value.Replace(".", ",");
                                        //realizamos el calculo
                                        v_newValor = v_newValor +  (v_totalNew - (double.Parse(v_retIva) + double.Parse(v_retRenta)));
                                    }
                                    if (v_tipoMoneda.Equals("USD"))
                                    {
                                        string v_totalLinea = Regex.Replace(oTotalLinea.Value, @"([A-Z])", string.Empty);
                                        double d_totalLinea = double.Parse(v_totalLinea);
                                        string v_totalLineaMone = Regex.Replace(oTotalLinea.Value, @"([.,0-9])", string.Empty);
                                        if (v_totalLineaMone.Contains("USD"))
                                        {
                                            string v_retIva = oValorIva.Value.Replace(".", ",");
                                            d_retIva = double.Parse(v_retIva);
                                            string v_retRenta = oValorRenta.Value.Replace(".", ",");
                                            d_retRenta = double.Parse(v_retRenta);
                                        }
                                        else
                                        {
                                            string v_retIva = oValorIva.Value.Replace(".", ",");
                                            d_retIva = Math.Round((double.Parse(v_retIva) / v_coti), 2);
                                            string v_retRenta = oValorRenta.Value.Replace(".", ",");
                                            d_retRenta = Math.Round((double.Parse(v_retRenta) / v_coti), 2);
                                        }
                                        
                                        //realizamos el calculo
                                        v_newValor = v_newValor + (d_totalLinea - (d_retIva + d_retRenta));
                                    }
                                }
                            }

                                v_fila++;
                        }
                        oTransferencia.Value = v_newValor.ToString();
                        oFechaTrans.Value = v_frete;


                    }
                    #endregion

                    #region EFECTIVO
                    if (pVal.ItemUID == "6" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)formPago.Items.Item("20").Specific;
                        //varibale de la seccion de transferencia
                        SAPbouiCOM.EditText oEfectivo = (SAPbouiCOM.EditText)form.Items.Item("38").Specific;
                        SAPbouiCOM.ComboBox oTipoMoneda = (SAPbouiCOM.ComboBox)form.Items.Item("8").Specific;
                        double v_coti = decimal.ToDouble(global_coti);
                        string v_tipoMoneda = oTipoMoneda.Selected.Value;
                        string v_tipoMonedaV = oTipoMoneda.Selected.Description;
                        double v_newValor = 0;
                        int v_rows = oMatrix.RowCount;
                        int v_fila = 1;
                        while (v_fila <= v_rows)
                        {
                            //variables de la tabla de pagos
                            SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("10000127").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oValorIva = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorIva").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oValorRenta = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorRenta").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oTotalLinea = (SAPbouiCOM.EditText)oMatrix.Columns.Item("24").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.ComboBox oTipoPago = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_TipoPago").Cells.Item(v_fila).Specific;
                            string v_tipoPago = oTipoPago.Value.ToString();
                            bool v_check = oCheck.Checked;
                            //consultamos si esta checkeado
                            if (v_check == true)
                            {
                                //si la opcion de pago es cheque
                                if (v_tipoPago.Equals("Efectivo"))
                                {
                                    //el tipo de moneda del documento
                                    if (v_tipoMoneda.Equals("GS") || v_tipoMonedaV.Contains("GS"))
                                    {
                                        string v_totalLinea = Regex.Replace(oTotalLinea.Value, @"([A-Z])", string.Empty);
                                        string v_totalLineaMone = Regex.Replace(oTotalLinea.Value, @"([.,0-9])", string.Empty);
                                        double v_totalNew = 0;
                                        if (v_totalLineaMone.Contains("USD"))
                                        {
                                            v_totalNew = Math.Round(double.Parse(v_totalLinea) * v_coti);
                                        }
                                        else
                                        {
                                            v_totalNew = double.Parse(v_totalLinea);
                                        }
                                        string v_retIva = oValorIva.Value.Replace(".", ",");
                                        string v_retRenta = oValorRenta.Value.Replace(".", ",");
                                        //realizamos el calculo
                                        v_newValor = v_newValor + (v_totalNew - (double.Parse(v_retIva) + double.Parse(v_retRenta)));
                                    }
                                    if (v_tipoMoneda.Equals("USD"))
                                    {
                                        string v_totalLinea = Regex.Replace(oTotalLinea.Value, @"([A-Z])", string.Empty);
                                        double d_totalLinea = double.Parse(v_totalLinea);
                                        string v_totalLineaMone = Regex.Replace(oTotalLinea.Value, @"([.,0-9])", string.Empty);
                                        double d_retIva = 0;
                                        double d_retRenta = 0;
                                        if (v_totalLineaMone.Contains("USD"))
                                        {
                                            string v_retIva = oValorIva.Value.Replace(".", ",");
                                            string v_retRenta = oValorRenta.Value.Replace(".", ",");

                                            d_retIva = double.Parse(v_retIva);
                                            d_retRenta = double.Parse(v_retRenta);
                                        }
                                        else
                                        {
                                            string v_retIva = oValorIva.Value.Replace(".", ",");
                                            d_retIva = Math.Round((double.Parse(v_retIva) / v_coti), 2);
                                            string v_retRenta = oValorRenta.Value.Replace(".", ",");
                                            d_retRenta = Math.Round((double.Parse(v_retRenta) / v_coti), 2);
                                        }
                                       
                                        //realizamos el calculo
                                        v_newValor = v_newValor + (d_totalLinea - (d_retIva + d_retRenta));
                                    }
                                }
                            }

                            v_fila++;
                        }
                        oEfectivo.Value = v_newValor.ToString();


                    }
                    #endregion
                }
                #endregion

                #region TOTAL RETENCIONES
                if (pVal.FormTypeEx == "426")
                {
                    if (pVal.ItemUID == "234000001" && pVal.EventType == BoEventTypes.et_CLICK && pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)formPago.Items.Item("20").Specific;
                        int v_rows = oMatrix.RowCount;
                        int v_fila = 1;
                        TotalRetIva = 0;
                        TotalRetRenta = 0;
                        while (v_fila <= v_rows)
                        {
                            SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("10000127").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oValorIva = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorIva").Cells.Item(v_fila).Specific;
                            SAPbouiCOM.EditText oValorRenta = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_RetValorRenta").Cells.Item(v_fila).Specific;

                            bool v_check = oCheck.Checked;
                            if (v_check == true)
                            {
                                double v_totalIva = double.Parse(oValorIva.Value.Replace(".", ","));
                                double v_totalRenta = double.Parse(oValorRenta.Value.Replace(".", ","));

                                TotalRetIva = TotalRetIva + v_totalIva;
                                TotalRetRenta = TotalRetRenta + v_totalRenta;
                            }
                            v_fila++;
                        }
                    }
                }
                #endregion

                #region FORM DE PLAZOS
                if (pVal.FormTypeEx == "65021")
                {
                    //agarramos el ID del form
                    SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);

                    //agregar boton de recalcular
                    if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);
                        Item oItem;
                        SAPbouiCOM.Button oButton;
                        oItem = oForm.Items.Add("btnRecal", BoFormItemTypes.it_BUTTON);
                        //Inicializando el objeto boton con la referencia del objeto item
                        oButton = (SAPbouiCOM.Button)oItem.Specific;
                        //Agregando propiedades al boton
                        oButton.Caption = "Recalcular";
                        //agregando posicio del boton
                        oItem.Top = oForm.Height - (oItem.Height + 5);
                        oItem.Left = (oItem.Width + 20) + 60;
                    }

                    if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE && pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item("3").Specific;
                        SAPbouiCOM.EditText oCuotaActual = (SAPbouiCOM.EditText)oMatrix.Columns.Item("19").Cells.Item(1).Specific;
                        totalcuota = double.Parse(string.Concat(oCuotaActual.Value.ToString().Where(c => Char.IsDigit(c))));
                    }

                    //evento de recalcular cuotas
                    if (pVal.ItemUID=="btnRecal" && pVal.EventType==BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item("3").Specific;
                        int v_filas = oMatrix.RowCount;
                        int cont = 1;
                        double v_cuotaAux = 0;
                        //recorremos la matrix
                        while(cont <= v_filas)
                        {
                            SAPbouiCOM.EditText oCuotaActual = (SAPbouiCOM.EditText)oMatrix.Columns.Item("19").Cells.Item(cont).Specific;
                            string v_cuotaACtual = oCuotaActual.Value;
                            v_cuotaACtual = string.Concat(v_cuotaACtual.Where(c => Char.IsDigit(c)));
                            double v_cuota = double.Parse(v_cuotaACtual);
                            if(v_cuota != totalcuota)
                            {
                                v_cuotaAux = v_cuotaAux + v_cuota;
                            }
                            cont++;
                        }
                        //calculamos las nuevas cuotas
                        double v_valorNuevo = totalFact - v_cuotaAux;

                        cont = 1;
                        int conAux = 0;
                        while (cont <= v_filas)
                        {
                            SAPbouiCOM.EditText oCuotaActual = (SAPbouiCOM.EditText)oMatrix.Columns.Item("19").Cells.Item(cont).Specific;
                            string v_cuotaACtual = oCuotaActual.Value;
                            v_cuotaACtual = string.Concat(v_cuotaACtual.Where(c => Char.IsDigit(c)));
                            double v_cuota = double.Parse(v_cuotaACtual);
                            if (v_cuota == totalcuota)
                            {
                                conAux++;
                            }
                            cont++;
                        }

                        double v_cuotaNueva = v_valorNuevo / conAux;

                        //cargamos las nuevas cuotas
                        cont = 1;
                        while (cont <= v_filas)
                        {
                            SAPbouiCOM.EditText oCuotaActual = (SAPbouiCOM.EditText)oMatrix.Columns.Item("19").Cells.Item(cont).Specific;
                            string v_cuotaACtual = oCuotaActual.Value;
                            v_cuotaACtual = string.Concat(v_cuotaACtual.Where(c => Char.IsDigit(c)));
                            double v_cuota = double.Parse(v_cuotaACtual);
                            if (v_cuota == totalcuota)
                            {
                                oCuotaActual.Value = v_cuotaNueva.ToString(); ;
                            }
                            cont++;
                        }


                    }
                                      
                    

                }
                #endregion


            }
            catch (Exception ex)
            {

            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }

        private static void conexion()
        {
            oCompany = (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
        }


        //funcion para construir el array
        private static void arrayCheque(List<object> dataArray,string fecha, double monto,string proveedor)
        {
            if (dataArray.Count > 0)
            {
                object json = dataArray;
                var jsontowrite = JsonConvert.SerializeObject(json, Newtonsoft.Json.Formatting.Indented);
                var jsondata = JsonConvert.DeserializeObject<dynamic>(jsontowrite);
                int v_filaArray = 0;
                bool v_proceso = false;
                foreach (var i in jsondata) 
                {                
                    var v_cheque = i.cheque;
                    var v_fecha = v_cheque.fecha;
                    string fecha_v = Convert.ToString(v_fecha);
                    var v_monto = v_cheque.monto;
                    string monto_v = Convert.ToString(v_monto);
                    double montoV = double.Parse(monto_v);
                    var v_prov = v_cheque.proveedor;
                    string prov_v = Convert.ToString(v_prov);
                    //si el proveedor y la fecha son el mismo, sumamos el valor
                    if(prov_v == proveedor && fecha_v == fecha)
                    {
                        monto = monto + montoV;
                        dataArray.RemoveAt(v_filaArray);
                        if (v_filaArray > 0)
                        {
                            v_filaArray--;
                        }
                        //creamos el objeto para el json
                        var dataObject = new
                        {
                            cheque = new Dictionary<string, object>
                            {
                                {"fecha",fecha },
                                {"monto",monto },
                                {"proveedor",proveedor }
                            }
                        };
                        dataArray.Add(dataObject);
                        v_proceso = true;
                    }
                    v_filaArray++;
                }
                if (v_proceso == false)
                {
                    //creamos el objeto para el json
                    var dataObject = new
                    {
                        cheque = new Dictionary<string, object>
                            {
                                {"fecha",fecha },
                                {"monto",monto },
                                {"proveedor",proveedor }
                            }
                    };
                    dataArray.Add(dataObject);
                }
            }
            else
            {
                //creamos el objeto para el json
                var dataObject = new
                {
                    cheque = new Dictionary<string, object>
                {
                    {"fecha",fecha },
                    {"monto",monto },
                    {"proveedor",proveedor }
                }
                };
                dataArray.Add(dataObject);
            }
            

           
        }

        //funcion para cargar los cheques
        private static void cargaCheques(List<object> dataArray, SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrixMP = (SAPbouiCOM.Matrix)oForm.Items.Item("28").Specific;
            object json = dataArray;
            var jsontowrite = JsonConvert.SerializeObject(json, Newtonsoft.Json.Formatting.Indented);
            var jsondata = JsonConvert.DeserializeObject<dynamic>(jsontowrite);
            //recorremos el array
            int v_filaCheque = 1;
            foreach (var i in jsondata)
            {
                SAPbouiCOM.EditText oImportaMP = (SAPbouiCOM.EditText)oMatrixMP.Columns.Item("7").Cells.Item(v_filaCheque).Specific;
                SAPbouiCOM.EditText oFecha = (SAPbouiCOM.EditText)oMatrixMP.Columns.Item("1").Cells.Item(v_filaCheque).Specific;
                var v_cheque = i.cheque;              
                var v_monto = v_cheque.monto;
                var v_fecha = v_cheque.fecha;
                string monto_v = Convert.ToString(v_monto);
                string fecha_v = Convert.ToString(v_fecha);
                oImportaMP.Value = monto_v;
                oFecha.Value = fecha_v;

                v_filaCheque++;
            }
            
        }

        //funcion para recalcular
        private static void recalcular(string docnum)
        {
            try
            {
                //eliminamos si es que ya existe
                SAPbobsCOM.Recordset oDelete;
                oDelete = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                oDelete.DoQuery("DELETE FROM \"@RET_CALCULO\" WHERE \"U_DocNum\"='"+docnum+"' ");

                //definimos variables
                string v_iva = null;
                string v_renta = null;
                decimal v_total = 0;
                decimal v_RetRenta = 0;
                decimal v_RetIva = 0;
                string v_Moneda = null;
                string v_DocEntry = null;
                int v_cuotas = 0;

                //consultamos el tipo de retención
                SAPbobsCOM.Recordset otipoRet;
                otipoRet = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                otipoRet.DoQuery("SELECT \"U_RetIVA\",\"U_RetRenta\",CASE WHEN \"DocCur\"='GS' THEN \"DocTotal\" ELSE \"DocTotalFC\" END,\"DocCur\",\"DocEntry\",\"Installmnt\" FROM OPCH WHERE \"DocNum\"='" + docnum + "' ");
                while (!otipoRet.EoF)
                {
                    v_iva = otipoRet.Fields.Item(0).Value.ToString();
                    v_renta = otipoRet.Fields.Item(1).Value.ToString();
                    v_total = decimal.Parse(otipoRet.Fields.Item(2).Value.ToString());
                    v_Moneda = otipoRet.Fields.Item(3).Value.ToString();
                    v_DocEntry = otipoRet.Fields.Item(4).Value.ToString();
                    v_cuotas = int.Parse(otipoRet.Fields.Item(5).Value.ToString());
                    otipoRet.MoveNext();
                }
                //impuesto
                SAPbobsCOM.Recordset oImpuesto;
                oImpuesto = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                oImpuesto.DoQuery("SELECT \"TaxCode\" FROM PCH1 WHERE \"DocEntry\"='" + v_DocEntry + "' ");
                string v_taxcode = oImpuesto.Fields.Item(0).Value.ToString();

                //calculamos la retencion
                if (v_iva.Equals("SI"))
                {
                    if (v_Moneda.Equals("USD"))
                    {
                        if (v_taxcode.Equals("IVA_5"))
                        {
                            v_RetIva = decimal.Round(((v_total / 21) * 30) / 100, 2);
                        }
                        if (v_taxcode.Equals("IVA_10"))
                        {
                            v_RetIva = decimal.Round(((v_total / 11) * 70) / 100, 2);
                        }

                    }
                    else
                    {
                        if (v_taxcode.Equals("IVA_5"))
                        {
                            v_RetIva = Math.Round(((v_total / 21) * 30) / 100);
                        }
                        if (v_taxcode.Equals("IVA_10"))
                        {
                            v_RetIva = Math.Round(((v_total / 11) * 70) / 100);
                        }
                    }

                }
                //en caso de ser solo IVA
                if (v_renta.Equals("SI"))
                {
                    if (v_Moneda.Equals("USD"))
                    {
                        v_RetRenta = decimal.Round((v_total / decimal.Parse("1,05")) * decimal.Parse("0,004"), 2);
                    }
                    else
                    {
                        v_RetRenta = Math.Round((v_total / decimal.Parse("1,05")) * decimal.Parse("0,004"));
                    }
                }

                //guardamos en la tabla de calculo de retención
                SAPbobsCOM.Recordset codMax;
                codMax = (Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                codMax.DoQuery("select CASE WHEN MAX(\"DocEntry\")=0 THEN 1 ELSE  MAX(\"DocEntry\") END from \"@RET_CALCULO\" ");
                int MaxCod = int.Parse(codMax.Fields.Item(0).Value.ToString()) + 1;


                string v_NumCuota = null;
                decimal v_montoCuota = 0;
                //consultamos si es por cuotas
                if (v_cuotas > 1)
                {
                    SAPbobsCOM.Recordset oCuotas;
                    oCuotas = (Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oCuotas.DoQuery("SELECT \"InstlmntID\",\"InsTotal\" FROM PCH6 WHERE \"DocEntry\"='" + v_DocEntry + "' ");
                    //recorremos
                    while (!oCuotas.EoF)
                    {
                        string v_cuota = " de " + v_cuotas;
                        v_NumCuota = oCuotas.Fields.Item(0).Value.ToString();
                        v_cuota = v_NumCuota + v_cuota;
                        v_montoCuota = decimal.Parse(oCuotas.Fields.Item(1).Value.ToString());
                        //calculamos la retencion
                        if (v_iva.Equals("SI"))
                        {
                            if (v_Moneda.Equals("USD"))
                            {
                                if (v_taxcode.Equals("IVA_5"))
                                {
                                    v_RetIva = decimal.Round(((v_montoCuota / 21) * 30) / 100, 2);
                                }
                                if (v_taxcode.Equals("IVA_10"))
                                {
                                    v_RetIva = decimal.Round(((v_montoCuota / 11) * 70) / 100, 2);
                                }

                            }
                            else
                            {
                                if (v_taxcode.Equals("IVA_5"))
                                {
                                    v_RetIva = Math.Round(((v_montoCuota / 21) * 30) / 100);
                                }
                                if (v_taxcode.Equals("IVA_10"))
                                {
                                    v_RetIva = Math.Round(((v_montoCuota / 11) * 70) / 100);
                                }
                            }

                        }
                        //en caso de ser solo IVA
                        if (v_renta.Equals("SI"))
                        {
                            if (v_Moneda.Equals("USD"))
                            {
                                v_RetRenta = decimal.Round((v_montoCuota / decimal.Parse("1,05")) * decimal.Parse("0,004"), 2);
                            }
                            else
                            {
                                v_RetRenta = Math.Round((v_montoCuota / decimal.Parse("1,05")) * decimal.Parse("0,004"));
                            }
                        }

                        SAPbobsCOM.Recordset oGrabar;
                        oGrabar = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oGrabar.DoQuery("INSERT INTO \"@RET_CALCULO\" (\"Code\",\"DocEntry\",\"U_DocNum\",\"U_Estado\",\"U_RetIva\",\"U_RetReta\",\"U_Plazo\",\"U_Moneda\") VALUES ('" + MaxCod + "','" + MaxCod + "','" + docnum + "','P','" + v_RetIva.ToString().Replace(",", ".") + "', '" + v_RetRenta.ToString().Replace(",", ".") + "', '" + v_cuota + "','" + v_Moneda + "') ");

                        MaxCod++;
                        oCuotas.MoveNext();
                    }
                }
                else
                {
                    SAPbobsCOM.Recordset oGrabar;
                    oGrabar = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oGrabar.DoQuery("INSERT INTO \"@RET_CALCULO\" (\"Code\",\"DocEntry\",\"U_DocNum\",\"U_Estado\",\"U_RetIva\",\"U_RetReta\",\"U_Plazo\",\"U_Moneda\") VALUES ('" + MaxCod + "','" + MaxCod + "','" + docnum + "','P','" + v_RetIva.ToString().Replace(",", ".") + "', '" + v_RetRenta.ToString().Replace(",", ".") + "','1 de 1','" + v_Moneda + "') ");
                }
            }
            catch (Exception e)
            {
                //System.Windows.Forms.MessageBox.Show(e.Message);
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.ToString(), 1, "OK");
            }
        }

        //funcion para leer el json
        private static void leerTXT(string url)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)formPago.Items.Item("20").Specific;
            int v_can = oMatrix.RowCount;
            int v_fila = 1;
            int v_invoiceid = 0;
            while (v_fila <= v_can)
            {
                //obtenemos el docnum de la factura
                SAPbouiCOM.EditText oDocNum = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(v_fila).Specific;
                SAPbouiCOM.EditText oNroFac = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_nrofactura").Cells.Item(v_fila).Specific;
                SAPbouiCOM.EditText oFechaDoc = (SAPbouiCOM.EditText)oMatrix.Columns.Item("21").Cells.Item(v_fila).Specific;
                SAPbouiCOM.EditText oNroPago = (SAPbouiCOM.EditText)formPago.Items.Item("3").Specific;
                string v_docnum = oDocNum.Value;
                string v_nroPago = oNroPago.Value;
                string v_nrofac = oNroFac.Value;
                string v_fechadoc = oFechaDoc.Value;
                //consultamos el numero de factura
                SAPbobsCOM.Recordset oConsulta;
                oConsulta = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                //oConsulta.DoQuery("SELECT \"DocEntry\",\"NumAtCard\" FROM OPCH WHERE \"DocNum\"='" + v_docnum + "' ");
                oConsulta.DoQuery("SELECT \"DocEntry\" FROM OPCH WHERE \"NumAtCard\"='"+ v_nrofac + "' AND \"DocDate\"='"+ v_fechadoc + "' AND \"CANCELED\"='N'");
                //string v_factura = oConsulta.Fields.Item(1).Value.ToString();
                string v_DocEntry = oConsulta.Fields.Item(0).Value.ToString();                              
                //agarramos el JSON
                string v_json = File.ReadAllText(url);
                JArray jsonArray = JArray.Parse(v_json);
                //recoremos el json
                foreach (JObject jsonOperaciones in jsonArray.Children<JObject>())
                {
                    //extraemos numero de factura
                    string v_datos = jsonOperaciones["datos"].ToString();
                    JObject oDetDatos = JObject.Parse(v_datos);
                    string v_transaccion = oDetDatos["transaccion"].ToString();
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


                v_fila++;
            }
            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Número de retención actualizado con éxito!!", 1, "OK");

        }
       

    }
}
