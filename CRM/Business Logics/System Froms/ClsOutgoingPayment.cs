using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRM
{
    class ClsOutgoingPayment : Base
    {
        #region Variable Declaration

        SAPbouiCOM.EditText _EditText;
        public static string StartDate = string.Empty;
        public static string EndDate = string.Empty;
        public static int ActivtyNo = 0;
        public static int _RowNo = 0;
        SAPbouiCOM.Matrix _Matrix2;
        SAPbouiCOM.OptionBtn _OptionBtn;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RS2 = null;
        public static bool _Flag = false;
        public static string _BookingCancellationNo = string.Empty;
        public static string _SalesOrderNo = string.Empty;
        public static string _SODecEntry = string.Empty;
      

        #endregion

        #region Form Default
        internal void FormDefault(string cardCode,string BookingCancallationNo,string SalesOrderNo,string SODecEntry,double DeductionCharge)
        {
            try
            {
                SAPbouiCOM.Form _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                _BookingCancellationNo = BookingCancallationNo;
                _SalesOrderNo = SalesOrderNo;
                _SODecEntry = SODecEntry;
                Utilities.ExecuteSQL(ref _RS, "Select * From OVPM Where \"U_CancNo\" = '" + BookingCancallationNo + "'");
                if (_RS.RecordCount > 0)
                {
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("U_CancNo").Specific).Value = _BookingCancellationNo.ToString();
                    _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    _OptionBtn = (SAPbouiCOM.OptionBtn)_Form.Items.Item("56").Specific;
                    _OptionBtn.Selected = true;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("5").Specific).Value = cardCode.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("U_CancNo").Specific).Value = _BookingCancellationNo.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("U_SONum").Specific).Value = _SalesOrderNo.ToString();
                    ((SAPbouiCOM.EditText)_Form.Items.Item("U_SO_DocNUm").Specific).Value = _SODecEntry.ToString();
                    ((SAPbouiCOM.CheckBox)_Form.Items.Item("37").Specific).Checked = true;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("13").Specific).Value = DeductionCharge.ToString();
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }

       
        #endregion

        #region ITEM EVENT
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            #region BEFORE ACTION ACTION
            if (pVal.Before_Action == true)
            {
                switch (pVal.EventType)
                {

                    #region ITEM PRESS EVENT

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            int RetVal = 0;
                            
                            if (_SODecEntry != "")
                            {
                                SAPbobsCOM.Documents _Order;
                                _Order = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                if (_Order.GetByKey(Convert.ToInt16(_SODecEntry)))
                                {
                                   // RetVal = _Order.Close();
                                    RetVal = _Order.Cancel();
                                    if (RetVal != 0)
                                    {
                                        string _errMsg = string.Empty;
                                        int _errCode = 0;
                                        Utilities.Application.Company.GetLastError(out _errCode, out _errMsg);
                                    }
                                }
                            }
                        }
                        break;
                         
                    #endregion

                }
            }
            #endregion

            #region After ACTION ACTION
            else if (pVal.Before_Action == false)
            {

                switch (pVal.EventType)
                {
                    //case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                    //    if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    //    {
                    //        int RetVal = 0;
                    //        string ErrCode;
                    //        string ErrMsg;

                    //        SAPbobsCOM.Documents _Order;
                    //        _Order = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                    //        if (_Order.GetByKey(Convert.ToInt16(_SODecEntry)))
                    //        {
                    //            RetVal = _Order.Close();
                    //            if (RetVal != 0)
                    //            {
                    //                string _errMsg = string.Empty;
                    //                int _errCode = 0;
                    //                Utilities.Application.Company.GetLastError(out _errCode, out _errMsg);
                    //            }
                    //        }
                    //    }
                    //    break;
                } 
            }

            #endregion

        }
        #endregion
    }
}
