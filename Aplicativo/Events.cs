using Aplicativo.Forms;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Aplicativo
{
    class Events
    {

        //Atributos da classe
        private SAPbobsCOM.Company company;
        private SAPbouiCOM.Application SBO_Application;   
        
        formPN Pn;
        formEst est;

        public Events()
        {

            try
            {
                SetApplication();           
                CompanyConnection();        

                SBO_Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);  
                SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent); 
                SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent); 
                SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);  
                SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);   
                SBO_Application.PrintEvent += new _IApplicationEvents_PrintEventEventHandler(SBO_Application_PrintEvent);   

                CriarMenus();
            }
            catch 
            {
                System.Environment.Exit(0); 
            }
        }
        
        
        private void SetApplication()       
        {
            SAPbouiCOM.SboGuiApi SboGuiApi = new SAPbouiCOM.SboGuiApi();    

            string sConnectionString = System.Convert.ToString("0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056");

            SboGuiApi.Connect(sConnectionString);                
            SBO_Application = SboGuiApi.GetApplication(-1);      
        }
        
        
        private void CompanyConnection()    
        {
            try
            {
                company = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();       
            }
            catch 
            {
                SBO_Application.StatusBar.SetText(company.GetLastErrorDescription(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        
        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent menuEvent, out bool BubbleEvent)   
                                                                                                           
        {
            BubbleEvent = true;     

            try
            {
                if (!menuEvent.BeforeAction)
                {
                    switch (menuEvent.MenuUID)
                    {

                       case "mn_Pn":
                            Pn = new formPN(SBO_Application, company);
                            Pn.ShowForm();
                            break;

                        case "mn_TransEst":
                           est = new formEst(SBO_Application, company);
                            est.ShowForm();
                            break;
                    }
                }


            }
            catch (Exception e)
            {
                SBO_Application.StatusBar.SetText("Erro UIEvents: " + e.Message.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        
        
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormTypeEx == "Pn")
            {
                Pn.itemEventPn(pVal, out BubbleEvent);
            }

            if (pVal.FormTypeEx == "est")
            {
                est.itemEventEstoq(pVal, out BubbleEvent);
            }


        }
        
        
        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo info, out bool BubbleEvent)
        {

            BubbleEvent = true;
        }
        
        
        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {

        }
        
        
        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
        
        
        private void SBO_Application_PrintEvent(ref SAPbouiCOM.PrintEventInfo info, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        void CriarMenus()
        {
            try
            {
                if (SBO_Application.Menus.Exists("mn_A"))
                    SBO_Application.Menus.Item("43520").SubMenus.Remove(SBO_Application.Menus.Item("mn_A"));
                

                if (!SBO_Application.Menus.Exists("mn_A"))
                {
                    MenuItem oMenuItem = SBO_Application.Menus.Item("43520");
                    Menus oMenus = oMenuItem.SubMenus;
                    MenuCreationParams oCreationPackage = (MenuCreationParams)SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = BoMenuType.mt_POPUP;
                    oCreationPackage.UniqueID = "mn_A";
                    oCreationPackage.String = "Addon Teste";
                    oCreationPackage.Position = 17;
                    oMenus.AddEx(oCreationPackage);
                }

                if (!SBO_Application.Menus.Exists("mn_Pn"))
                {
                    MenuItem oMenuItem = SBO_Application.Menus.Item("mn_A");
                    Menus oMenus = oMenuItem.SubMenus;
                    MenuCreationParams oCreationPackage = (MenuCreationParams)SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "mn_Pn";
                    oCreationPackage.String = "Parceiro de Negócio";
                    oCreationPackage.Position = 1;
                    oMenus.AddEx(oCreationPackage);
                }

                if (!SBO_Application.Menus.Exists("mn_TransEst"))
                {
                    MenuItem oMenuItem = SBO_Application.Menus.Item("mn_A");
                    Menus oMenus = oMenuItem.SubMenus;
                    MenuCreationParams oCreationPackage = (MenuCreationParams)SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "mn_TransEst";
                    oCreationPackage.String = "Transferência de Estoque";
                    oCreationPackage.Position = 2;
                    oMenus.AddEx(oCreationPackage);
                }


            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
