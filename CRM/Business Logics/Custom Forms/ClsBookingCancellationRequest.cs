using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;

namespace CRM
{
    class ClsBookingCancellationRequest : Base
    {
        #region Variable declaration

        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RS2 = null, _RS3 = null, _RSCheck = null;
        private SAPbouiCOM.Matrix _Matrix, _Matrix1, _Matrix2, _Matrix3;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource, _ChildDBDataSource1, _ChildDBDataSource2, _ChildDBDataSource3;
        SAPbouiCOM.ComboBox _ComboBox;
        public string _CardCode = string.Empty;
        public string _DocEntry = string.Empty;
        public static string _SDocNum = string.Empty;
        public static string _SDocEntry = string.Empty;
        private int RowToDelete;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsBookingCancellationRequest()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsBookingCancellationRequest()
        {
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
                        if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            Utilities.ExecuteSQL(ref _RSCheck, "Select \"DocStatus\" From \"ORDR\" Where \"DocStatus\" = 'C' AND \"DocNum\" = '" + _SDocNum + "' AND \"DocEntry\" = '" + _SDocEntry + "'");
                            if(_RSCheck.RecordCount > 0)
                            {
                                Utilities.Message("Can not make the cancel request for closed Sales Order", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            //bool approved = false;
                            //for (int i = 1; i <= _Matrix3.RowCount; i++)
                            //{
                            //    if (((SAPbouiCOM.ComboBox)_Matrix3.Columns.Item("V_3").Cells.Item(i).Specific).Selected.Value != "Approved")
                            //    {
                            //        Utilities.Message("Without approved decision can not open booking cancellation screen....", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //        BubbleEvent = false;
                            //        return;
                            //    }

                            //}

                        }
                        if(pVal.ItemUID == "46" && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            Utilities.Message("Please add Cancellation request first", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }
                        if (pVal.ItemUID == "25")
                        {
                            _Form.PaneLevel = 1;
                        }
                        else if (pVal.ItemUID == "26")
                        {
                            _Form.PaneLevel = 2;
                        }
                        else if (pVal.ItemUID == "31")
                        {
                            _Form.PaneLevel = 3;
                        }
                        else if (pVal.ItemUID == "35")
                        {
                            _Form.PaneLevel = 4;
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
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "46")
                        {

                            this._Object = new ClsBookingCancellation();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.BookingCancellation);
                            ((ClsBookingCancellation)_Object).FormDefault(_Form.DataSources.DBDataSources.Item("@BOOKINGCRM").GetValue("U_DocNum", 0).ToString(), _Form.DataSources.DBDataSources.Item("@BOOKINGCRM").GetValue("U_CustCode", 0).ToString(), _Form.DataSources.DBDataSources.Item("@BOOKINGCRM").GetValue("U_SODocNum", 0).ToString(),_DocEntry);

                        }
                        break;


                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        if (pVal.ItemUID == "36" && pVal.ColUID == "V_3")
                        {
                            if (((SAPbouiCOM.ComboBox)_Matrix3.Columns.Item("V_3").Cells.Item(pVal.Row).Specific).Selected.Value != "")
                            {
                                ((SAPbouiCOM.ComboBox)_Matrix3.Columns.Item("V_4").Cells.Item(pVal.Row).Specific).Select("Answered", BoSearchKey.psk_ByDescription);

                            }
                            else
                            {
                                ((SAPbouiCOM.ComboBox)_Matrix3.Columns.Item("V_4").Cells.Item(pVal.Row).Specific).Select("Pending", BoSearchKey.psk_ByDescription);
                            }
                            bool approved = false;
                            for (int i = 1; i <= _Matrix3.RowCount; i++)
                            {
                                if (((SAPbouiCOM.ComboBox)_Matrix3.Columns.Item("V_3").Cells.Item(i).Specific).Selected.Value == "Approved")
                                {
                                    approved = true;
                                }
                                else
                                {
                                    approved = false;
                                    return;
                                }
                            }
                            if (approved == true)
                            {
                                Utilities.ExecuteSQL(ref _RS, "update \"@BOOKINGCRM\" SET \"U_Status\" = 'Approved' Where \"U_SODocNum\" = '" + _Form.DataSources.DBDataSources.Item("@BOOKINGCRM").GetValue("U_SODocNum", 0).ToString() + "'");
                            }
                        }
                        break;

                   case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                    if(pVal.ItemUID == "32" && pVal.ColUID == "V_4" && pVal.InnerEvent == false && pVal.Action_Success == true)
                    {
                        _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                        _Matrix2 = (SAPbouiCOM.Matrix)Form.Items.Item("32").Specific;

                        if (((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_11").Cells.Item(pVal.Row).Specific).Value == "")
                        {
                            int RowNo = 0;
                            bool flag = true;
                            RowNo = pVal.Row;
                            Utilities.Application.SBO_Application.ActivateMenuItem("2563");
                            this._Object = new ClsActivty();
                            ((ClsActivty)this._Object).FormDefault(_CardCode, RowNo, flag);
                        }
                        else
                        {
                            int RowNo = 0;
                            bool flag = false;
                            RowNo = pVal.Row;
                           string _ActivtyNo = ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_11").Cells.Item(pVal.Row).Specific).Value.ToString();
                            Utilities.Application.SBO_Application.ActivateMenuItem("2563");
                            this._Object = new ClsActivty();
                            ((ClsActivty)this._Object).FormDefaultExit(_CardCode, RowNo, flag, _ActivtyNo);
                            //_Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                            //_Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            //((SAPbouiCOM.EditText)_Form.Items.Item("5").Specific).Value = _ActivtyNo;
                            //_Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }

                    }
                    break;

                }
            }

            #endregion

        }

        #endregion

        #region EVENT :Menu Event

        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            #region Menu_Event Before Action
            if (pVal.BeforeAction == true)
            {
                switch (pVal.MenuUID)
                {
                    case Constants.System_Menus.mnu_DELETE_ROW:
                        //_DBDataSource.RemoveRecord(this.RowToDelete - 1);
                        //if (_Matrix.VisualRowCount > 0)
                        //{
                        //    if (Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        //    {
                        //        Utilities.ShowErrorMessage("Please update the value first.");
                        //        BubbleEvent = false;
                        //        return;
                        //    }
                        //    else
                        //    {
                        //        _Matrix.DeleteRow(_Matrix.VisualRowCount);
                        //        _Matrix.FlushToDataSource();
                        //        _Matrix.LoadFromDataSource();
                        //        if (_Visual.ContainsKey(this.RowToDelete))
                        //            _Visual.Remove(this.RowToDelete);

                        //    }

                        //}
                        break;

                }
            }
            #endregion

            #region Menu_Event After Action
            else
            {
                switch (pVal.MenuUID)
                {
                    case Constants.System_Menus.mnu_ADD:
                                                  
                        break;

                    case Constants.System_Menus.mnu_FIRST:
                       
                        break;

                    case Constants.System_Menus.mnu_NEXT:
                        
                        break;

                    case Constants.System_Menus.mnu_PREVIOUS:
                        
                        break;

                    case Constants.System_Menus.mnu_LAST:
                       
                        break;

                }

            }
            #endregion
        }

        #endregion

        #region FormData_Event
        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {

            #region Before Action
            if (BusinessObjectInfo.BeforeAction)
            {
                try
                {
                    switch (BusinessObjectInfo.EventType)
                    {

                        #region et_FORM_DATA_ADD Before action
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                           
                            break;
                        #endregion

                        #region et_FORM_DATA_UPDATE Before action
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //if (BusinessObjectInfo.ActionSuccess == true)
                            //{
                            //    if (_DBDataSource.GetValue("U_StageCode", _DBDataSource.Size - 1).Trim() == "" && _DBDataSource.GetValue("U_ParamCode", _DBDataSource.Size - 1).Trim() == "")
                            //    {
                            //        _DBDataSource.RemoveRecord(_DBDataSource.Size - 1);
                            //    }
                            //}
                            break;
                        #endregion

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
                    switch (BusinessObjectInfo.EventType)
                    {
                        #region et_FORM_DATA_LOAD After Action

                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            _Matrix2 = (SAPbouiCOM.Matrix)Form.Items.Item("32").Specific;
                            if (_ChildDBDataSource2 != null && _Matrix2 != null)
                            {
                                if (_ChildDBDataSource2.GetValue("U_StartDt", _ChildDBDataSource2.Size - 1).Trim() != "")
                                {
                                    _ChildDBDataSource2.InsertRecord(_ChildDBDataSource2.Size);
                                    _ChildDBDataSource2.SetValue("U_CreateAc", _ChildDBDataSource2.Size - 1, ".");
                                    _Matrix2.LoadFromDataSource();
                                }

                            }
                            for (int i = 2; i < _Matrix2.VisualRowCount; i++)
                            {
                                _Matrix2.CommonSetting.SetRowEditable(i - 1, false);
                            }
                                break;

                        #endregion
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

        #region FormDefault

        public void FormDefault(string DocEntry,string CardCode,string DocNum)
        {
            _Form.Freeze(true);
            _CardCode = CardCode;
            _DocEntry = DocEntry;
            _SDocEntry = DocEntry;
            _SDocNum = DocNum;
            _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
            _Form.DataBrowser.BrowseBy = "10";
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@BOOKINGCRM");
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@BOOKINGCRC1");
            _ChildDBDataSource1 = _Form.DataSources.DBDataSources.Item("@BOOKINGCRC2");
            _ChildDBDataSource2 = _Form.DataSources.DBDataSources.Item("@BOOKINGCRC3");
            _ChildDBDataSource3 = _Form.DataSources.DBDataSources.Item("@BOOKINGCRC4");

            _ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item("130").Specific;
            _ComboBox.ValidValues.Add("Open", "Open");
            _ComboBox.ValidValues.Add("Approved", "Approved");
            _ComboBox.ValidValues.Add("Closed", "Closed");
            _ComboBox.Select("Open", BoSearchKey.psk_ByDescription);

            Utilities.ExecuteSQL(ref _RS, "Select * From \"@BOOKINGCRM\" Where \"U_SODocNum\" = '" + DocNum + "' AND \"U_DocEntry\" = '" + DocEntry + "'");
             if (_RS.RecordCount > 0)
             {
                 _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                 ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = DocNum;
                 ((SAPbouiCOM.EditText)_Form.Items.Item("44").Specific).Value = DocEntry;
                 _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                 _Form.Items.Item("112").Enabled = false;
                 _Form.Items.Item("4").Enabled = false;
                 _Form.Items.Item("6").Enabled = false;
                 _Form.Items.Item("8").Enabled = false;
                 _Form.Items.Item("22").Enabled = false;
                 _Form.Items.Item("24").Enabled = false;
                 _Form.Items.Item("30").Enabled = false;
                 _Form.Items.Item("10").Enabled = false;
                 _Form.Items.Item("12").Enabled = false;
                 _Form.Items.Item("18").Enabled = false;
                 _Form.Items.Item("20").Enabled = false;
                 _Form.Items.Item("40").Enabled = false;
                 _Form.Items.Item("27").Enabled = false;
                 _Form.Items.Item("28").Enabled = false;
                 _Form.Items.Item("130").Enabled = false;
                 _Matrix3 = (SAPbouiCOM.Matrix)Form.Items.Item("36").Specific;
                 _Matrix3.Columns.Item(0).Editable = false;
                 _Matrix3.Columns.Item(1).Editable = false;
                 _Matrix3.Columns.Item(2).Editable = false;
                 _Matrix3.Columns.Item(3).Editable = false;
                 _Matrix3.Columns.Item(4).Editable = false;

                 string _SQL3 = "select \"USER_CODE\",\"U_NAME\",T1.\"Name\" as \"Department\" from OUSR T0 inner join \"OUDP\" T1 ON T1.\"Code\" = T0.\"Department\" where \"U_AuthCanc\" = 'Yes'";

                 Utilities.ExecuteSQL(ref _RS3, _SQL3);
                 for (int i = 0; i < _RS3.RecordCount; i++)
                 {
                     if (((SAPbouiCOM.EditText)_Matrix3.Columns.Item("V_6").Cells.Item(i + 1).Specific).Value != Utilities.Application.Company.UserName)
                     {
                         _Matrix3.CommonSetting.SetRowEditable(i + 1, false);
                     }
                     _RS3.MoveNext();
                 }
             }
             else
             {

                 //_Matrix.AddRow(1, -1);

                 string _SQL = "SELECT T0.\"CardCode\",T0.\"CardName\",T2.\"FirstName\" || ' ' || T2.\"LastName\" as \"Contact Person\",T3.\"PymntGroup\" as \"Payment Terms\", " +
                                " T4.\"Balance\" as \"Advance Payment Amount\",T0.\"DocNum\" as \"SO DocNum\",T0.\"DocEntry\" as \"SO DocEntry\",T0.\"DocDate\" as \"SO Date\",T5.\"SlpName\" as \"Sales Employee\", " +
                                " T7.\"Name\" as \"Opportuntiy Name\",T7.\"OpprId\" as \"Opportunity No\" " +
                                " FROM ORDR T0 " +
                                " INNER JOIN RDR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                                " LEFT OUTER JOIN OCPR T2 ON T0.\"CntctCode\" = T2.\"CntctCode\" " +
                                " LEFT OUTER JOIN OCTG T3 ON T0.\"GroupNum\" = T3.\"GroupNum\" " +
                                " INNER JOIN OCRD T4 ON T4.\"CardCode\" = T0.\"CardCode\" " +
                                " LEFT OUTER JOIN OSLP T5 ON T5.\"SlpCode\" = T0.\"SlpCode\" " +
                                " LEFT OUTER JOIN OPR1 T6 ON T6.\"DocNumber\" = T0.\"DocNum\" AND T6.\"DocId\" = T0.\"DocEntry\" and T6.\"ObjType\" = '17' " +
                                " LEFT OUTER JOIN OOPR T7 ON T7.\"OpprId\" = T6.\"OpprId\" " +
                                " WHERE T0.\"DocEntry\" = '" + DocEntry + "' ";
                 Utilities.ExecuteSQL(ref _RS, _SQL);

                 ((SAPbouiCOM.EditText)_Form.Items.Item("112").Specific).Value = _RS.Fields.Item("CardCode").Value.ToString();
                 ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value = _RS.Fields.Item("CardName").Value.ToString();
                 ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value = _RS.Fields.Item("Contact Person").Value.ToString();
                 ((SAPbouiCOM.EditText)_Form.Items.Item("8").Specific).Value = _RS.Fields.Item("Advance Payment Amount").Value.ToString();
                 ((SAPbouiCOM.EditText)_Form.Items.Item("22").Specific).Value = _RS.Fields.Item("Payment Terms").Value.ToString();
                 ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = Utilities.Application.Company.UserName;
                 ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value = Utilities.getMaxColumnValueNum("@BOOKINGCRM", "U_DocNum");
                 ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = _RS.Fields.Item("SO DocNum").Value.ToString();
                 ((SAPbouiCOM.EditText)_Form.Items.Item("44").Specific).Value = _RS.Fields.Item("SO DocEntry").Value.ToString();
                 ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = Convert.ToDateTime(_RS.Fields.Item("SO Date").Value.ToString()).ToString("yyyyMMdd");
                 ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value = _RS.Fields.Item("Sales Employee").Value.ToString();
                 ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = _RS.Fields.Item("Opportunity No").Value.ToString();
                 ((SAPbouiCOM.EditText)_Form.Items.Item("40").Specific).Value = _RS.Fields.Item("Opportuntiy Name").Value.ToString();

                 _Form.Items.Item("112").Enabled = false;
                 _Form.Items.Item("4").Enabled = false;
                 _Form.Items.Item("6").Enabled = false;
                 _Form.Items.Item("8").Enabled = false;
                 _Form.Items.Item("22").Enabled = false;
                 _Form.Items.Item("24").Enabled = false;
                 _Form.Items.Item("30").Enabled = false;
                 _Form.Items.Item("10").Enabled = false;
                 _Form.Items.Item("12").Enabled = false;
                 _Form.Items.Item("18").Enabled = false;
                 _Form.Items.Item("20").Enabled = false;
                 _Form.Items.Item("43").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                 _Form.Items.Item("40").Enabled = false;
                 _Form.Items.Item("44").Enabled = false;
                 _Form.Items.Item("27").Enabled = false;
                 _Form.Items.Item("28").Enabled = false;
                 _Form.Items.Item("130").Enabled = false;
                 _Matrix3 = (SAPbouiCOM.Matrix)Form.Items.Item("36").Specific;
                 _Matrix3.Columns.Item(0).Editable = false;
                 _Matrix3.Columns.Item(1).Editable = false;
                 _Matrix3.Columns.Item(2).Editable = false;
                 _Matrix3.Columns.Item(3).Editable = false;
                 _Matrix3.Columns.Item(4).Editable = false;


                 string _SQL1 = "select T1.\"ItemCode\",T1.\"Dscription\",T1.\"Quantity\",T1.\"Price\",T1.\"WhsCode\",T1.\"Project\",T1.\"VatGroup\",T1.\"LineTotal\", " +
                                " T3.\"DistNumber\",T3.\"MnfSerial\",T3.\"U_ModelYear\",T3.\"U_RegNo\", " +
                                " T3.\"U_KeyNo\",T3.\"U_BatteryNo\",T3.\"U_Made\" " +
                                " from ORDR T0 " +
                                " INNER JOIN RDR1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                                " LEFT OUTER JOIN SRI1 T2 on T1.\"ItemCode\"=T2.\"ItemCode\"   and (T1.\"DocEntry\"=T2.\"BaseEntry\" and T1.\"ObjType\"=T2.\"BaseType\") " +
                                " LEFT OUTER JOIN OSRN T3 on T3.\"ItemCode\"=T2.\"ItemCode\" and T2.\"SysSerial\"=T3.\"SysNumber\" " +
                                " where T0.\"DocEntry\" = '" + DocEntry + "'";

                 Utilities.ExecuteSQL(ref _RS1, _SQL1);
                 for (int i = 0; i < _RS1.RecordCount; i++)
                 {
                     _Matrix.AddRow(1, -1);
                     _Matrix.FlushToDataSource();
                     _ChildDBDataSource.SetValue("U_ItemCode", i, _RS1.Fields.Item("ItemCode").Value.ToString());
                     _ChildDBDataSource.SetValue("U_ItmDscrp", i, _RS1.Fields.Item("Dscription").Value.ToString());
                     _ChildDBDataSource.SetValue("U_Quantity", i, _RS1.Fields.Item("Quantity").Value.ToString());
                     _ChildDBDataSource.SetValue("U_Price", i, _RS1.Fields.Item("Price").Value.ToString());
                     _ChildDBDataSource.SetValue("U_Whs", i, _RS1.Fields.Item("WhsCode").Value.ToString());
                     _ChildDBDataSource.SetValue("U_Project", i, _RS1.Fields.Item("Project").Value.ToString());
                     _ChildDBDataSource.SetValue("U_TaxCode", i, _RS1.Fields.Item("VatGroup").Value.ToString());
                     _ChildDBDataSource.SetValue("U_Total", i, _RS1.Fields.Item("LineTotal").Value.ToString());
                     _ChildDBDataSource.SetValue("U_ChasisNo", i, _RS1.Fields.Item("DistNumber").Value.ToString());
                     _ChildDBDataSource.SetValue("U_EngineNo", i, _RS1.Fields.Item("MnfSerial").Value.ToString());
                     _ChildDBDataSource.SetValue("U_ModealYr", i, _RS1.Fields.Item("U_ModelYear").Value.ToString());
                     _ChildDBDataSource.SetValue("U_RegNo", i, _RS1.Fields.Item("U_RegNo").Value.ToString());
                     _ChildDBDataSource.SetValue("U_KeyNo", i, _RS1.Fields.Item("U_KeyNo").Value.ToString());
                     _Matrix.LoadFromDataSource();
                     _RS1.MoveNext();
                 }

                 _Matrix1 = (SAPbouiCOM.Matrix)Form.Items.Item("28").Specific;
                 string _SQL2 = "select 'Incoming Payment' as \"Type\",T0.\"DocEntry\",T0.\"DocNum\",T0.\"DocDate\",T0.\"DocTotal\", " +
                                " T1.\"CheckNum\",T1.\"DueDate\" as \"CheckDate\",T2.\"BankName\",T1.\"Branch\",T0.\"CheckSum\",T0.\"TrsfrSum\",T0.\"TrsfrDate\", " +
                                " T0.\"CreditSum\",T0.\"CashSum\",T3.\"U_NAME\" " +
                                " from ORCT T0 " +
                                " left outer join RCT1 T1 ON T0.\"DocEntry\" = T1.\"DocNum\" " +
                                " left outer join ODSC T2 ON T2.\"BankCode\" = T1.\"BankCode\" " +
                                " left outer join OUSR T3 ON T3.\"USERID\" = T0.\"UserSign\" " +
                                " where \"CardCode\" = '" + CardCode + "' " +
                                " UNION ALL " +
                                " select 'Outgoing Payment' as \"Type\",T0.\"DocEntry\",T0.\"DocNum\",T0.\"DocDate\",T0.\"DocTotal\", " +
                                " T1.\"CheckNum\",T1.\"DueDate\" as \"CheckDate\",T2.\"BankName\",T1.\"Branch\",T0.\"CheckSum\",T0.\"TrsfrSum\",T0.\"TrsfrDate\", " +
                                " T0.\"CreditSum\",T0.\"CashSum\",T3.\"U_NAME\" " +
                                " from OVPM T0 " +
                                " left outer join VPM1 T1 ON T0.\"DocEntry\" = T1.\"DocNum\" " +
                                " left outer join ODSC T2 ON T2.\"BankCode\" = T1.\"BankCode\" " +
                                " left outer join OUSR T3 ON T3.\"USERID\" = T0.\"UserSign\" " +
                                " where \"CardCode\" = '" + CardCode + "' " +
                                " UNION ALL " +
                                " SELECT 'Journal Entry' AS \"Type\",T0.\"TransId\" AS \"DocEntry\",T0.\"Number\" AS \"DocNum\",T0.\"RefDate\" AS \"DocDate\", " +
                                " CASE WHEN T1.\"Credit\" > 0 THEN -T1.\"Credit\" ELSE T1.\"Debit\" END AS \"DocTotal\", " +
                                " 0 AS \"CheckNum\",'' AS \"CheckDate\",'' AS \"BankName\",'' AS \"Branch\",0 AS \"CheckSum\",0 AS \"TrsfrSum\",'' AS \"TrsfrDate\", " +
                                " 0 AS \"CreditSum\",0 AS \"CashSum\",T2.\"U_NAME\" " +
                                " from OJDT T0 " +
                                " inner join JDT1 T1 ON T0.\"TransId\" = T1.\"TransId\" " +
                                " left outer join OUSR T2 ON T2.\"USERID\" = T1.\"CreatedBy\" " +
                                " where T0.\"Number\" = T0.\"BaseRef\" AND T1.\"ShortName\" = '" + CardCode + "'";

                 Utilities.ExecuteSQL(ref _RS2, _SQL2);

                 for (int i = 0; i < _RS2.RecordCount; i++)
                 {
                     _Matrix1.AddRow(1, -1);
                     _Matrix1.FlushToDataSource();
                     _ChildDBDataSource1.SetValue("U_DocType", i, _RS2.Fields.Item("Type").Value.ToString());
                     _ChildDBDataSource1.SetValue("U_DocNum", i, _RS2.Fields.Item("DocNum").Value.ToString());
                     _ChildDBDataSource1.SetValue("U_DocDate", i, Convert.ToDateTime(_RS2.Fields.Item("DocDate").Value.ToString()).ToString("yyyyMMdd"));
                     _ChildDBDataSource1.SetValue("U_TotAmoun", i, _RS2.Fields.Item("DocTotal").Value.ToString());
                     _ChildDBDataSource1.SetValue("U_CheqNo", i, _RS2.Fields.Item("CheckNum").Value.ToString());
                     _ChildDBDataSource1.SetValue("U_CheqDt", i, Convert.ToDateTime(_RS2.Fields.Item("CheckDate").Value.ToString()).ToString("yyyyMMdd"));
                     _ChildDBDataSource1.SetValue("U_BankNm", i, _RS2.Fields.Item("BankName").Value.ToString());
                     _ChildDBDataSource1.SetValue("U_CheqAmt", i, _RS2.Fields.Item("CheckSum").Value.ToString());
                     _ChildDBDataSource1.SetValue("U_BnkTrnsA", i, _RS2.Fields.Item("TrsfrSum").Value.ToString());
                     _ChildDBDataSource1.SetValue("U_TransDt", i, Convert.ToDateTime(_RS2.Fields.Item("TrsfrDate").Value.ToString()).ToString("yyyyMMdd"));
                     _ChildDBDataSource1.SetValue("U_CrdCrAmt", i, _RS2.Fields.Item("CreditSum").Value.ToString());
                     _ChildDBDataSource1.SetValue("U_CashAmt", i, _RS2.Fields.Item("CashSum").Value.ToString());
                     _ChildDBDataSource1.SetValue("U_EntrBy", i, _RS2.Fields.Item("U_NAME").Value.ToString());
                     _RS2.MoveNext();
                     _Matrix1.LoadFromDataSource();
                 }

                 _Matrix2 = (SAPbouiCOM.Matrix)Form.Items.Item("32").Specific;
                 _Matrix2.AddRow(1, -1);

                 _Matrix2.FlushToDataSource();
                 _ChildDBDataSource2.SetValue("U_CreateAc", 0, ".");
                 _Matrix2.LoadFromDataSource();

                 _Matrix3 = (SAPbouiCOM.Matrix)Form.Items.Item("36").Specific;
                 _Matrix3.AddRow(1, -1);
                 
                 
                
                 string _SQL3 = "select \"USER_CODE\",\"U_NAME\",T1.\"Name\" as \"Department\" from OUSR T0 inner join \"OUDP\" T1 ON T1.\"Code\" = T0.\"Department\" where \"U_AuthCanc\" = 'Yes'";

                 Utilities.ExecuteSQL(ref _RS3, _SQL3);
                 for (int i = 0; i < _RS3.RecordCount; i++)
                 {
                     if (i > 0)
                     {
                         _Matrix3.AddRow(1, -1);
                     }
                     _ComboBox = (SAPbouiCOM.ComboBox)_Matrix3.Columns.Item("V_4").Cells.Item(i + 1).Specific;
                     _ComboBox.Select("Pending", SAPbouiCOM.BoSearchKey.psk_ByDescription);

                     _Matrix3.FlushToDataSource();
                     _ChildDBDataSource3.SetValue("U_Departm", i, _RS3.Fields.Item("Department").Value.ToString());
                     _ChildDBDataSource3.SetValue("U_UsrCode", i, _RS3.Fields.Item("USER_CODE").Value.ToString());
                     _ChildDBDataSource3.SetValue("U_UsrName", i, _RS3.Fields.Item("U_NAME").Value.ToString());
                  
                     _Matrix3.LoadFromDataSource();
                     if (((SAPbouiCOM.EditText)_Matrix3.Columns.Item("V_5").Cells.Item(i+1).Specific).Value !=  Utilities.Application.Company.UserName)
                     {
                         _Matrix3.CommonSetting.SetRowEditable(i+1, false);
                     }
                     _RS3.MoveNext();
                 }

             }
             _Form.Items.Item("27").Visible = true;
             _Form.Items.Item("28").Visible = true;
             _Form.Items.Item("32").Visible = true;
             _Form.Items.Item("36").Visible = true;
             _Form.Items.Item("25").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
             _Form.Freeze(false);
        }

        #endregion

    }
}
