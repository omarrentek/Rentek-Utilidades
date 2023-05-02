using SAPbouiCOM;
using System;

namespace Utilidades.Classes
{
    internal class clsBaseMaestra
    {
        private Application oApp;
        private SAPbobsCOM.Company oCompany;
        private ItemEvent pVal;

        public clsBaseMaestra(Application oApp, SAPbobsCOM.Company oCompany, ItemEvent pVal)
        {
            this.oApp = oApp;
            this.oCompany = oCompany;
            this.pVal = pVal;
        }

        internal void AddOptionCombobox(ref Matrix matrizReal, string valueSelected, ItemEvent pVal, ref Form formBaseMaestra)
        {
            ComboBox combosubtipoReal = null;
            //Column subtipoReal = null;
            string[] mueble = new string[] {"Contratos","Equipos","Vehículos","Fuente de pago","Fideiocomiso de garantías o de admin","Cesiones" }; 
            string[] inmueble = new string[] {"Vivienda","Local comercial","Rural"};
            try
            {
                //subtipoReal = (Column)matrizReal.Columns.Item("C_3_3");
                //subtipoReal.DisplayDesc = true;
                combosubtipoReal = (ComboBox)matrizReal.Columns.Item("C_3_9").Cells.Item(pVal.Row).Specific;
                //combosubtipoReal.Item.DisplayDesc = true;

                formBaseMaestra.Freeze(true);

                if (combosubtipoReal.ValidValues.Count > 1)
                {
                    int contValcombo = combosubtipoReal.ValidValues.Count;
                    for (int i = 0; i < contValcombo; i++)
                    {
                        combosubtipoReal.ValidValues.Remove(0, BoSearchKey.psk_Index); //Elimina si hay subtipos en el campo
                    }
                }

                if (valueSelected.Equals("Mueble"))
                {
                    for(int i = 0;i <= mueble.Length -1; i++)
                    {
                        combosubtipoReal.ValidValues.Add(mueble[i], mueble[i]);
                    }
                }else if (valueSelected.Equals("Inmueble"))
                {
                    for (int i = 0; i <= inmueble.Length - 1; i++)
                    {
                        combosubtipoReal.ValidValues.Add(inmueble[i], inmueble[i]);
                    }
                }

               
            }
            catch (Exception ex)
            {
                formBaseMaestra.Freeze(false);
                oApp.MessageBox("Error AddOptionCombobox Add-on Utilidades: " + "\n" + pVal.EventType.ToString() + "\n" + pVal.ItemUID + "\n" + ex.Message.ToString() + " " + ex.StackTrace.ToString());
            }
            finally
            {
                formBaseMaestra.Freeze(false);
                if (combosubtipoReal != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(combosubtipoReal);
                }

            }
            




        }

    }
}