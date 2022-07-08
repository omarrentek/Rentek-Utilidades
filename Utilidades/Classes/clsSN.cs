using SAPbouiCOM;
using System;

namespace Utilidades.Classes
{
    internal class clsSN
    {
        private Application oApp;
        private SAPbobsCOM.Company oCompany;
        private Form oForm;
        private Item oItem;
        private Item oRefItem;
        private Folder oFolder;
        private StaticText oLabel;
        private EditText oEditText;
        private CheckBox oCheckBox;
        private ComboBox oComboBox;
        private Button oButton;
        private SAPbobsCOM.Recordset oRecordset;



        public clsSN(Application oApp, SAPbobsCOM.Company oCompany)
        {
            this.oApp = oApp;
            this.oCompany = oCompany;
        }

        internal void AddControls(ItemEvent pVal)
        {
            try
            {
                oForm = (Form)oApp.Forms.Item(pVal.FormUID);

                oForm.DataSources.UserDataSources.Add("FldUDS1", BoDataType.dt_SHORT_TEXT, 10);


                oRefItem = oForm.Items.Item("9");
                oItem = oForm.Items.Add("FldFctOrd", BoFormItemTypes.it_FOLDER);
                oItem.Top = oRefItem.Top;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.AffectsFormMode = false;
                oItem.Enabled = true;
                oItem.LinkTo = "9";
                oItem.AffectsFormMode = false;
                oFolder = (SAPbouiCOM.Folder)oItem.Specific;
                oFolder.Caption = "Facturación y/o Órdenes de compra";
                oFolder.DataBind.SetBound(true, "", "FldUDS1");
                oFolder.GroupWith("9");

                oRefItem = oForm.Items.Item("78");
                oItem = oForm.Items.Add("DtsEnc", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width + 70;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = 30;
                oItem.ToPane = 30;
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Datos del encargado de la vinculación";
                oLabel.Item.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_BOLD;

                /*Nombre encargado de vinculación*/
                oRefItem = oForm.Items.Item("75");
                oItem = oForm.Items.Add("EdNmEnc", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left - 70;
                oItem.Width = oRefItem.Width + 30;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = 30;
                oItem.ToPane = 30;
                oItem.LinkTo = "DtsEnc";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_NomEncV");
                //
                oRefItem = oForm.Items.Item("DtsEnc");
                oItem = oForm.Items.Add("NomEncS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width - 140;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdNmEnc";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Nombre";

                /*Cargo encargado de vinculación*/
                oRefItem = oForm.Items.Item("EdNmEnc");
                oItem = oForm.Items.Add("EdCrEnc", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdNmEnc";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_CrgEncV");
                //
                oRefItem = oForm.Items.Item("NomEncS");
                oItem = oForm.Items.Add("CrgEncS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdCrEnc";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Cargo";

                /*Email encargado de vinculación*/
                oRefItem = oForm.Items.Item("EdCrEnc");
                oItem = oForm.Items.Add("EdEmEnc", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdCrEnc";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_EmlEncV");
                //
                oRefItem = oForm.Items.Item("CrgEncS");
                oItem = oForm.Items.Add("EmEncS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdEmEnc";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Email";

                /*Telefono encargado de vinculación*/
                oRefItem = oForm.Items.Item("EdEmEnc");
                oItem = oForm.Items.Add("EdTlEnc", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdEmEnc";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_TelEncV");
                //
                oRefItem = oForm.Items.Item("EmEncS");
                oItem = oForm.Items.Add("TelEncS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdTlEnc";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Teléfono";


                /****************************************** Datos area de compras ***********************************************/

                oRefItem = oForm.Items.Item("DtsEnc");
                oItem = oForm.Items.Add("DtsEncC", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 95;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.Enabled = true;
                oItem.LinkTo = "DtsEnc";
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Datos del área de compras";
                oLabel.Item.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_BOLD;

                /*Nombre encargado de vinculación*/
                oRefItem = oForm.Items.Item("EdNmEnc");
                oItem = oForm.Items.Add("EdNmEncC", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 95;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = 30;
                oItem.ToPane = 30;
                oItem.LinkTo = "DtsEncC";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_NomEncC");
                //
                oRefItem = oForm.Items.Item("NomEncS");
                oItem = oForm.Items.Add("NomEncC", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 95;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdNmEncC";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Nombre";

                /*Cargo encargado de vinculación*/
                oRefItem = oForm.Items.Item("EdCrEnc");
                oItem = oForm.Items.Add("EdCrEncC", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 95;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdNmEncC";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_CrgEncC");
                //
                oRefItem = oForm.Items.Item("CrgEncS");
                oItem = oForm.Items.Add("CrgEncC", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 95;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdCrEncC";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Cargo";

                /*Email encargado de vinculación*/
                oRefItem = oForm.Items.Item("EdEmEnc");
                oItem = oForm.Items.Add("EdEmEncC", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 95;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "CrgEncC";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_EmlEncC");
                //
                oRefItem = oForm.Items.Item("EmEncS");
                oItem = oForm.Items.Add("EmEncC", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 95;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdEmEncC";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Email";

                /*Telefono encargado de vinculación*/
                oRefItem = oForm.Items.Item("EdTlEnc");
                oItem = oForm.Items.Add("EdTlEncC", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 95;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EmEncC";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_TelEncC");
                //
                oRefItem = oForm.Items.Item("TelEncS");
                oItem = oForm.Items.Add("TelEncC", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 95;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EdTlEncC";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Teléfono";

                /******************* Requerimientos *****************************/

                oRefItem = oForm.Items.Item("DtsEnc");
                oItem = oForm.Items.Add("DtsRqrm", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top;
                oItem.Left = oRefItem.Left + 300;
                oItem.Width = oRefItem.Width + 70;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.Enabled = true;
                oItem.LinkTo = "DtsEnc";
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Requerimientos para radicación de facturación";
                oLabel.Item.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_BOLD;

                //
                oRefItem = oForm.Items.Item("DtsRqrm");
                oItem = oForm.Items.Add("RqFact", BoFormItemTypes.it_CHECK_BOX);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width + 50;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "DtsRqrm";
                oItem.Enabled = true;
                oItem.AffectsFormMode = true;
                oCheckBox = (CheckBox)oItem.Specific;
                oCheckBox.Caption = "¿Requiere emitir Orden de Compra para recibir factura?";
                oCheckBox.DataBind.SetBound(true, "OCRD", "U_EmOrdCm");

                /*Nombre encargado de radicacion facturación*/
                oRefItem = oForm.Items.Item("RqFact");
                oItem = oForm.Items.Add("ENmOrCm", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left + 100;
                oItem.Width = oRefItem.Width - 120;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "RqFact";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_NmOrCm");
                oEditText.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                //
                oRefItem = oForm.Items.Item("RqFact");
                oItem = oForm.Items.Add("NmOrCm", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width - 220;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "ENmOrCm";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Nombre";
                oLabel.Item.Visible = false;

                /*Email encargado de radicacion facturación*/
                oRefItem = oForm.Items.Item("ENmOrCm");
                oItem = oForm.Items.Add("EmlOrCm", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "ENmOrCm";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_EmlOrCm");
                oEditText.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                //
                oRefItem = oForm.Items.Item("NmOrCm");
                oItem = oForm.Items.Add("EmlOrCmS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EmlOrCm";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Email";

                /*Contacto encargado de radicacion facturación*/
                oRefItem = oForm.Items.Item("EmlOrCm");
                oItem = oForm.Items.Add("CntOrdCm", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EmlOrCm";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_CntOrdCm");
                oEditText.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                //
                oRefItem = oForm.Items.Item("EmlOrCmS");
                oItem = oForm.Items.Add("CntOrdCmS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "CntOrdCm";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Contacto";

                /*Detalle de radicacion facturación*/
                oRefItem = oForm.Items.Item("CntOrdCmS");
                oItem = oForm.Items.Add("DtlOrdCmS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width + 250;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "CntOrdCmS";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Detalle su proceso para solicitud y envío de Ordenes de Compra";
                //
                oRefItem = oForm.Items.Item("DtlOrdCmS");
                oItem = oForm.Items.Add("DtlOrdCm", BoFormItemTypes.it_EXTEDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width - 50;
                oItem.Height = oRefItem.Height + 100;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "DtlOrdCmS";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_DtlOrdCm");
                oEditText.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);

                /*Periodicidad de solicitud de orden de compra*/
                oRefItem = oForm.Items.Item("RqFact");
                oItem = oForm.Items.Add("SlOrFac", BoFormItemTypes.it_COMBO_BOX);
                oItem.Top = oRefItem.Top;
                oItem.Left = oRefItem.Left + 600;
                oItem.Width = oRefItem.Width - 200;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "RqFact";
                oItem.AffectsFormMode = true;
                oComboBox = (ComboBox)oItem.Specific;
                oComboBox.DataBind.SetBound(true, "OCRD", "U_SlOrFac");
                oComboBox.Item.DisplayDesc = true;
                oComboBox.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                //
                oRefItem = oForm.Items.Item("RqFact");
                oItem = oForm.Items.Add("SlOrFacS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top;
                oItem.Left = oRefItem.Left + 300;
                oItem.Width = oRefItem.Width - 20;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "SlOrFac";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "¿Cada cuanto se debe solicitar Orden de Compra para facturar?";

                /*Otra periodicidad*/
                oRefItem = oForm.Items.Item("SlOrFac");
                oItem = oForm.Items.Add("SlOrOtr", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left - 200;
                oItem.Width = oRefItem.Width + 200;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "EmlOrCm";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_SlOrOtr");
                oEditText.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                //
                oRefItem = oForm.Items.Item("SlOrFacS");
                oItem = oForm.Items.Add("SlOrOtrS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width - 200;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "SlOrOtr";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Otra periodicidad";

                /*Tiempo estimado para envio de orden de compra*/
                oRefItem = oForm.Items.Item("SlOrFac");
                oItem = oForm.Items.Add("Tmpest", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 30;
                oItem.Left = oRefItem.Left + 20;
                oItem.Width = oRefItem.Width - 20;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "SlOrOtr";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_Tmpest");
                //
                oRefItem = oForm.Items.Item("SlOrFacS");
                oItem = oForm.Items.Add("TmpestS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 30;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width + 20;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "Tmpest";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Tiempo estimado de respuesta para envío de la Orden de Compra";

                /*Tipo Tiempo estimado para envio de orden de compra*/
                oRefItem = oForm.Items.Item("Tmpest");
                oItem = oForm.Items.Add("TpTmpEs", BoFormItemTypes.it_COMBO_BOX);
                oItem.Top = oRefItem.Top;
                oItem.Left = oRefItem.Left + 90;
                oItem.Width = oRefItem.Width - 50;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "Tmpest";
                oItem.AffectsFormMode = true;
                oComboBox = (ComboBox)oItem.Specific;
                oComboBox.Item.DisplayDesc = true;
                oComboBox.DataBind.SetBound(true, "OCRD", "U_TpTmpEs");

                /*Documentos adicionales*/
                oRefItem = oForm.Items.Item("TmpestS");
                oItem = oForm.Items.Add("DocAdic", BoFormItemTypes.it_CHECK_BOX);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "TpTmpEs";
                oItem.Enabled = true;
                oItem.AffectsFormMode = true;
                oCheckBox = (CheckBox)oItem.Specific;
                oCheckBox.DataBind.SetBound(true, "OCRD", "U_DocAdic");
                oCheckBox.Caption = "¿Requiere otros documentos adicionales para recibir factura?";
                /*Documentos*/
                oRefItem = oForm.Items.Item("SlOrOtr");
                oItem = oForm.Items.Add("DtlDocAd", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 45;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "SlOrOtr";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_DtlDocAd");
                oEditText.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                //
                oRefItem = oForm.Items.Item("SlOrOtrS");
                oItem = oForm.Items.Add("DtlDocAdS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 45;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "Tmpest";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Documentos";

                /*Periodicidad otros documentos*/
                oRefItem = oForm.Items.Item("DtlDocAd");
                oItem = oForm.Items.Add("PrdDoc", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "DtlDocAd";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_PrdDoc");
                oEditText.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                //
                oRefItem = oForm.Items.Item("DtlDocAdS");
                oItem = oForm.Items.Add("PrdDocS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "PrdDoc";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Periodicidad";

                /*Otros mecanismos*/
                oRefItem = oForm.Items.Item("DocAdic");
                oItem = oForm.Items.Add("PrcdRecp", BoFormItemTypes.it_CHECK_BOX);
                oItem.Top = oRefItem.Top + 45;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width + 30;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "DocAdic";
                oItem.Enabled = true;
                oItem.AffectsFormMode = true;
                oCheckBox = (CheckBox)oItem.Specific;
                oCheckBox.DataBind.SetBound(true, "OCRD", "U_PrcdRecp");
                oCheckBox.Caption = "¿Tiene otros mecanismos o procedimientos para la recepción de facturas?";

                /*Detalle de mecanismos*/
                oRefItem = oForm.Items.Item("PrdDoc");
                oItem = oForm.Items.Add("DtlRecp", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 30;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "DtlDocAd";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_DtlRecp");
                oEditText.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                //
                oRefItem = oForm.Items.Item("PrdDocS");
                oItem = oForm.Items.Add("DtlRecpS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 30;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "DtlRecp";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Detalle";

                /*Días del mes recepción de facturas*/
                oRefItem = oForm.Items.Item("Tmpest");
                oItem = oForm.Items.Add("CrtFact", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 90;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "DtlDocAd";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_CrtFact");
                //
                oRefItem = oForm.Items.Item("TmpestS");
                oItem = oForm.Items.Add("CrtFactS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 90;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "CrtFact";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Fecha de corte para la recepción de facturas del mes";

                /*Fecha envío*/
                oRefItem = oForm.Items.Item("CrtFact");
                oItem = oForm.Items.Add("FchEnv", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left - 200;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "CrtFact";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_FchEnv");
                //
                oRefItem = oForm.Items.Item("CrtFactS");
                oItem = oForm.Items.Add("FchEnvS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width - 200;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "FchEnv";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Fecha envío";

                /*Fecha vinculación*/
                oRefItem = oForm.Items.Item("FchEnv");
                oItem = oForm.Items.Add("Fchvinc", BoFormItemTypes.it_EDIT);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "DtlDocAd";
                oItem.AffectsFormMode = true;
                oEditText = (EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, "OCRD", "U_Fchvinc");
                //
                oRefItem = oForm.Items.Item("FchEnvS");
                oItem = oForm.Items.Add("FchvincS", BoFormItemTypes.it_STATIC);
                oItem.Top = oRefItem.Top + 15;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "Fchvinc";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oLabel = (StaticText)oItem.Specific;
                oLabel.Caption = "Fecha vinculación";

                oRefItem = oForm.Items.Item("FchvincS");
                oItem = oForm.Items.Add("BtnImp", BoFormItemTypes.it_BUTTON);
                oItem.Top = oRefItem.Top + 20;
                oItem.Left = oRefItem.Left;
                oItem.Width = oRefItem.Width - 50;
                oItem.Height = oRefItem.Height + 5;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "Fchvinc";
                oItem.Enabled = true;
                oItem.AffectsFormMode = false;
                oButton = (Button)oItem.Specific;
                oButton.Caption = "Imprimir";

            }
            //catch (Exception ex)
            //{
            //    oApp.MessageBox("Error LoadControls: " + ex.ToString() + "\n" + ex.StackTrace.ToString());
            //}
            finally
            {
                if (oForm != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
                if (oFolder != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oFolder);
                }
                if (oItem != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                }
                if (oRefItem != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRefItem);
                }
                if (oButton != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oButton);
                }
                if (oLabel != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oLabel);
                }
                if (oEditText != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText);
                }
                if (oCheckBox != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCheckBox);
                }
                if (oComboBox != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oComboBox);
                }
            }

        }

        internal void Print(ItemEvent pVal)
        {
            string menu = "vinculacion_proveedor", menuId = string.Empty,cardcode = string.Empty;
            try
            {
                oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery($"SELECT T0.MenuUID FROM [OCMN] T0 where T0.Name LIKE '{menu}' AND T0.Type = 'C'");
                if(oRecordset != null)
                {
                    oRecordset.MoveFirst();
                    menuId = oRecordset.Fields.Item("MenuUID").Value;
                }
                if (!String.IsNullOrEmpty(menuId))
                {
                    oForm = (Form)oApp.Forms.Item(pVal.FormUID);
                    cardcode = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0);
                    oApp.ActivateMenuItem(menuId);
                    oForm = (Form)oApp.Forms.ActiveForm;
                    oEditText = (EditText)oForm.Items.Item("1000003").Specific;
                    oButton = (Button)oForm.Items.Item("1").Specific;
                    oEditText.Value = cardcode;
                    oButton.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
               
            }
            catch (Exception ex)
            {
                oApp.MessageBox("Error print Add-On Utilidades: " + ex.ToString() + "\n" + ex.StackTrace.ToString());

            }
            finally
            {
                if (oForm != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
                if (oButton != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oButton);
                }
                if (oEditText != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText);
                }
            }
        }

        internal void EnableFields(ItemEvent pVal, string item)
        {
            string[] fields = new string[] { };
            try
            {
                oForm = (Form)oApp.Forms.Item(pVal.FormUID);
                if (!item.Equals("SlOrFac"))
                {
                    oCheckBox = (CheckBox)oForm.Items.Item(item).Specific;

                }
                else
                {
                    oComboBox = (ComboBox)oForm.Items.Item(item).Specific;

                }

                switch (item)
                {
                    case "RqFact":
                        oForm.Freeze(true);
                        fields = new string[] { "ENmOrCm", "EmlOrCm", "CntOrdCm", "DtlOrdCm", "SlOrFac" };
                        ActiveFields(fields, oForm, oCheckBox.Checked);
                        break;
                    case "DocAdic":
                        oForm.Freeze(true);
                        fields = new string[] { "DtlDocAd", "PrdDoc" };
                        ActiveFields(fields, oForm, oCheckBox.Checked);
                        break;
                    case "PrcdRecp":
                        oForm.Freeze(true);
                        fields = new string[] { "DtlRecp" };
                        ActiveFields(fields, oForm, oCheckBox.Checked);
                        break;
                    case "SlOrFac":

                        if (oComboBox.Selected != null)
                        {
                            oForm.Freeze(true);
                            fields = new string[] { "SlOrOtr" };
                            ActiveFields(fields, oForm, true, oComboBox.Selected.Value);

                        }
                        break;
                }

            }
            catch (Exception ex)
            {
                oForm.Freeze(false);

                oApp.MessageBox("Error EnableFields Add-On Utilidades: " + ex.ToString() + "\n" + ex.StackTrace.ToString());

            }
            finally
            {
                oForm.Freeze(false);

                if (oCheckBox != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCheckBox);
                }
                if (oForm != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
                if (oComboBox != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oComboBox);
                }
            }
        }

        private void ActiveFields(string[] fields, Form oForm, bool @checked, string optionCombo = null)
        {
            foreach (string field in fields)
            {
                if (!field.Equals("SlOrOtr"))
                {
                    if (!@checked)
                    {
                        oForm.Items.Item(field).SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True);
                    }
                    else
                    {
                        oForm.Items.Item(field).SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                    }
                }
                else
                {
                    if (optionCombo.Equals("OT"))
                    {
                        oForm.Items.Item(field).SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True);

                    }
                    else
                    {
                        oForm.Items.Item(field).SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);

                    }

                }

            }
        }
    }
}