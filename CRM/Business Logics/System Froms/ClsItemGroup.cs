using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using Microsoft.ApplicationBlocks.Data;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using System.Collections;
using System.Threading;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
using Sap.Data.Hana;


namespace CRM
{
    class ClsItemGroup : Base
    {

         #region VARIABLE DECLARATION
        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        private bool flag;
        public SAPbouiCOM.Matrix _Matrix;
        SAPbouiCOM.StaticText _StaticText;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        private static HanaConnection _conn;
        private static HanaDataAdapter _da;
        //private static DataTable _dataTable;
        private static HanaParameter _parm1;
        private static HanaParameter _parm2;
        private static int _errors = 0;
        string strcon, _Sql;
        private SAPbouiCOM.EditText _EditText;
        private SAPbouiCOM.ComboBox _ComboBox;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource;
        private SAPbobsCOM.Recordset _RS;
        private int RowID { get; set; }
        private string _Code { get; set; }
        private string _LineId { get; set; }
        private string _ActivityType = string.Empty;
        string _ItemGroupCode = string.Empty, _ItemGroupName = string.Empty;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsItemGroup()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsItemGroup()
        {
        }
        #endregion

        #region Item Event
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.Before_Action)
            {
                try
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                           

                           
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
            #endregion

            #region After Action
            else
            {
                try
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            if (pVal.ItemUID == "btnDtSetup")
                            {
                                _ItemGroupCode = Form.DataSources.DBDataSources.Item("OITB").GetValue("ItmsGrpCod", 0).ToString().Trim();
                                _ItemGroupName = Form.DataSources.DBDataSources.Item("OITB").GetValue("ItmsGrpNam", 0).ToString();
                                this._Object = new ClsDiscountSetup();
                                Utilities.LoadForm(ref this._Object, Constants.Forms.DiscountSetup);
                                ((ClsDiscountSetup)this._Object).FormDefault(_ItemGroupCode, _ItemGroupName);
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
            #endregion
        }

        #endregion

        #region Form Default
        internal void FormDefault()
        {
            try
            {
                _ExistItem = _Form.Items.Item("2");
                _NewItem = _Form.Items.Add("btnDtSetup", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
                _NewItem.Width = _ExistItem.Width + 50;
                _NewItem.Top = _ExistItem.Top;
                _NewItem.FromPane = _ExistItem.FromPane;
                _NewItem.ToPane = _ExistItem.ToPane;
                _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                _Button.Caption = "Discount Setup";

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion


    }
}
