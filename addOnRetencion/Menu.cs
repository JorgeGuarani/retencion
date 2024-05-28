using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM.Framework;

namespace addOnRetencion
{
    class Menu
    {
        public static SAPbobsCOM.Company sbo = null;
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;
            sbo =  (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43537"); // moudles'
           

            //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            //oCreationPackage.UniqueID = "addOnRetencion";
            //oCreationPackage.String = "addOnRetencion";
            //oCreationPackage.Enabled = true;
            //oCreationPackage.Position = -1;

            //oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                if (oMenus.Exists("addOnRetencion.Form1"))
                {
                    oMenus.RemoveEx("addOnRetencion.Form1");
                }
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception e)
            {

            }
            //prueba
            oMenus = oMenuItem.SubMenus;
            // Create s sub menu
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            oCreationPackage.UniqueID = "addOnRetencion.Form1";
            oCreationPackage.String = "Exportar a txt";
            oCreationPackage.Position = 3;
            oMenus.AddEx(oCreationPackage);
            //try
            //{
            //    // Get the menu collection of the newly added pop-up item
            //    oMenuItem = Application.SBO_Application.Menus.Item("addOnRetencion");
            //    oMenus = oMenuItem.SubMenus;

            //    // Create s sub menu
            //    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            //    oCreationPackage.UniqueID = "addOnRetencion.Form1";
            //    oCreationPackage.String = "Form1";
            //    oMenus.AddEx(oCreationPackage);
            //}
            //catch (Exception er)
            //{ //  Menu already exists
            //    Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            //}
            Application.SBO_Application.SetStatusBarMessage("AddOn de retenciones conectado", SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "addOnRetencion.Form1")
                {
                    Form1 export = new Form1();
                    export.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
