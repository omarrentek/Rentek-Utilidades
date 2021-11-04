using SAPbouiCOM;
using System;
using System.Text;
using System.Reflection;

namespace Utilidades.Classes
{
    internal class clsProyectos
    {
        private Application oApp;
        private SAPbobsCOM.Company oCompany;
        private ItemEvent pVal;
        private Form oForm;

        public clsProyectos(Application oApp, SAPbobsCOM.Company oCompany, ItemEvent pVal)
        {
            this.oApp = oApp;
            this.oCompany = oCompany;
            this.pVal = pVal;
        }

        internal void LoadControls(ItemEvent pVal)
        {
            SAPbouiCOM.Item oItemRef = null, oItem = null;
            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.EditText oEditText = null;
            try
            {
                oForm = (SAPbouiCOM.Form)oApp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                oItemRef = (SAPbouiCOM.Item)oForm.Items.Item("2");
                oItem = oForm.Items.Add("BtnBscr", BoFormItemTypes.it_BUTTON);
                oItem.Visible = true;
                oItem.Top = oItemRef.Top;
                oItem.Left = oItemRef.Left + 100;
                oItem.Width = oItemRef.Width;
                oItem.Height = oItemRef.Height;
                oItem.AffectsFormMode = false;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Buscar";

                oItemRef = (SAPbouiCOM.Item)oForm.Items.Item("BtnBscr");
                oItem = oForm.Items.Add("EdtBscar", BoFormItemTypes.it_EDIT);
                oItem.Visible = true;
                oItem.Top = oItemRef.Top;
                oItem.Left = oItemRef.Left + oItemRef.Width + 10;
                oItem.Width = oItemRef.Width + 100;
                oItem.Height = oItemRef.Height - 4;
                oItem.AffectsFormMode = false;
                oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            }
            catch (Exception ex)
            {
                oApp.MessageBox("Error LoadControls: " + ex.ToString() + " " + ex.StackTrace.ToString());
            }
            finally
            {

                if (oItem != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                }
                if (oItemRef != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItemRef);
                }
                if (oEditText != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText);
                }
                if (oButton != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oButton);
                }
            }
        }

        internal void LoadData(ItemEvent pVal)
        {
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.Button oButton = null;

            string queryOprj = null;
            StringBuilder oStringBuilder;
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {

                oForm = oApp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                oForm.Freeze(true);
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("EdtBscar").Specific;
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                oButton = (SAPbouiCOM.Button)oForm.Items.Item("BtnBscr").Specific;


                queryOprj = new System.IO.StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}.SQL.OPRJ.sql")).ReadToEnd();

                oStringBuilder = new StringBuilder();
                queryOprj = oStringBuilder.AppendFormat(queryOprj, oEditText.Value.ToString()).ToString();
                rs.DoQuery(queryOprj);
                rs.MoveFirst();


                if (rs != null)
                {
                    if (rs.RecordCount > 0)
                    {
                        oMatrix.Columns.Item("U_CodigoSN").Cells.Item(Int32.Parse(rs.Fields.Item(0).Value.ToString())).Click(BoCellClickType.ct_Regular);
                        oForm.Mode = BoFormMode.fm_OK_MODE;
                        oForm.Freeze(false);
                        oButton.Item.Refresh();
                        oForm.Update();
                    }
                }

            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                oApp.SetStatusBarMessage("Error loadData: " + ex.ToString() + " " + ex.StackTrace.ToString(), BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                if (oButton != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oButton);
                }
                if (oMatrix != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix);
                }
                if (oEditText != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText);
                }
            }
        }

    }
}