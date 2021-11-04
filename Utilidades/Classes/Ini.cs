using System;
using SAPbouiCOM;
using System.Reflection;

namespace Utilidades.Classes
{
    class Ini
    {
        private static SAPbobsCOM.Company oCompany;
        private static SAPbouiCOM.Application oApp;
        private static SboGuiApi oGuiApi;
        private static Assembly asm;

        private static clsProyectos frmProyectos;
        private static clsBaseMaestra frmBaseMaestra;

        private Ini()
        {
            ConectaSap();
            Filtros();
            Eventos();
            //MenusApp();
        }

        private void ConectaSap()
        {
            string Connectionstr = string.Empty;
            asm = GetType().Assembly;
            if (!String.IsNullOrEmpty(Environment.GetCommandLineArgs().GetValue(1).ToString()))
            {
                Connectionstr = Environment.GetCommandLineArgs().GetValue(1).ToString();
                oGuiApi = new SboGuiApi();
                oGuiApi.Connect(Connectionstr);
                oApp = oGuiApi.GetApplication();
                oApp.SetStatusBarMessage("Estamos conectando el add-On " + asm.GetName().Name + " por favor espere unos segundos.",BoMessageTime.bmt_Short,false);
                oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            }
            else
            {
                Connectionstr = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                oGuiApi = new SboGuiApi();
                oGuiApi.Connect(Connectionstr);
                oApp = oGuiApi.GetApplication();
                oApp.SetStatusBarMessage("Estamos conectando el add-On " + asm.GetName().Name + " por favor espere unos segundos.");
                oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            }
        }

        private void Filtros()
        {
            SAPbouiCOM.EventFilters oFilters = null;
            oFilters = oApp.GetFilter();

            if (oFilters == null)
            {
                oFilters = new SAPbouiCOM.EventFilters();
                oApp.SetFilter(oFilters);
            }

            SAPbouiCOM.EventFilter oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            oFilter.AddEx("711");
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            oFilter.AddEx("711");
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            oFilter.AddEx("UDO_FT_RNTK_MASTERB");

            oApp.SetFilter(oFilters);
        }

        private void Eventos()
        {
            oApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            oApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        }

        internal static bool Run()
        {
            try
            {
                _ = new Ini();
                oApp.SetStatusBarMessage("Add-On " + asm.GetName().Name + " conectado exitosamente",BoMessageTime.bmt_Short,false);
                return true;
            }catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Execpción al conectar add-On Rentek utilidades, \n" + ex.Message + "\n" + ex.StackTrace);
                return false;
            }
           
        }

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (!pVal.BeforeAction)
                {
                    switch (pVal.FormTypeEx)
                    {
                        case"711":
                            switch (pVal.EventType)
                            {
                                case BoEventTypes.et_FORM_LOAD:
                                    if (frmProyectos == null)
                                    {
                                        frmProyectos = new clsProyectos(oApp, oCompany, pVal);
                                    }
                                    frmProyectos.LoadControls(pVal);
                                    break;
                                case BoEventTypes.et_CLICK:
                                    if (pVal.ItemUID == "BtnBscr")
                                    {
                                        if (frmProyectos == null)
                                        {
                                            frmProyectos = new clsProyectos(oApp, oCompany, pVal);
                                        }
                                        frmProyectos.LoadData(pVal);
                                        break;
                                    }
                                    break;
                            }
                            break;
                        case "UDO_FT_RNTK_MASTERB":
                            switch (pVal.EventType)
                            {
                                case BoEventTypes.et_COMBO_SELECT:
                                    if (pVal.ItemUID == "3_U_G") // matriz real
                                    {
                                        if (pVal.ColUID.Equals("C_3_2")) // combo tipo real
                                        {
                                            ComboBox comboTipoReal = null;
                                            Form formBaseMaestra = null;
                                            Matrix MatrizReal = null;

                                            if (frmBaseMaestra == null)
                                            {
                                                frmBaseMaestra = new clsBaseMaestra(oApp, oCompany, pVal);
                                            }

                                            formBaseMaestra = (Form)oApp.Forms.Item(pVal.FormUID);
                                            MatrizReal = (Matrix)formBaseMaestra.Items.Item("3_U_G").Specific;
                                            comboTipoReal = (ComboBox)MatrizReal.Columns.Item("C_3_2").Cells.Item(pVal.Row).Specific;
                                            if (comboTipoReal.Selected != null)
                                            {
                                                frmBaseMaestra.AddOptionCombobox(ref MatrizReal, comboTipoReal.Selected.Value, pVal, ref formBaseMaestra);
                                            }


                                            if (comboTipoReal != null)
                                            {
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(comboTipoReal);
                                            }
                                            if (formBaseMaestra != null)
                                            {
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(formBaseMaestra);
                                            }
                                            if (MatrizReal != null)
                                            {
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(MatrizReal);
                                            }
                                        }                                        
                                        break;
                                    }
                                    break;
                            }
                            break;
                        
                    }

                }

            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage("Error SBO_Application_ItemEvent: " + "\n" + pVal.EventType.ToString() + "\n" + pVal.ItemUID + "\n" + ex.Message.ToString() + " " + ex.StackTrace.ToString(), BoMessageTime.bmt_Short, true);
                BubbleEvent = false;
            }
        }

        private void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case BoAppEventTypes.aet_ShutDown:
                        oCompany.Disconnect();
                        oCompany = null;
                        System.Windows.Forms.Application.Exit();
                        break;
                    case BoAppEventTypes.aet_ServerTerminition:
                        oCompany.Disconnect();
                        oCompany = null;
                        System.Windows.Forms.Application.Exit();
                        break;
                    case BoAppEventTypes.aet_CompanyChanged:
                        oCompany.Disconnect();
                        oCompany = null;
                        System.Windows.Forms.Application.Exit();
                        break;
                    case BoAppEventTypes.aet_FontChanged:
                        oCompany.Disconnect();
                        oCompany = null;
                        System.Windows.Forms.Application.Exit();
                        break;
                    case BoAppEventTypes.aet_LanguageChanged:
                        oCompany.Disconnect();
                        oCompany = null;
                        System.Windows.Forms.Application.Exit();
                        break;
                }


            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage("Error: " + ex.ToString() + " " + ex.StackTrace.ToString());

            }
        }
    
    }
}
