using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;

namespace CRM
{
    class ClsBookingCancellation : Base
    {
        #region Variable declaration

        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RS2 = null, _RS3 = null, _RS4 = null;
        private SAPbouiCOM.Matrix _Matrix, _Matrix1, _Matrix2, _Matrix3;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource, _ChildDBDataSource1, _ChildDBDataSource2, _ChildDBDataSource3;
        SAPbouiCOM.ComboBox _ComboBox;
        public string _CardCode = string.Empty;
        public string _SODocEntry = string.Empty; 
        private int RowToDelete;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsBookingCancellation()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsBookingCancellation()
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
                        if (pVal.ItemUID == "41")
                        {
                            if (Convert.ToDouble(_Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_RfndAmt", 0).ToString()) > 0)
                            {
                                Utilities.Application.SBO_Application.ActivateMenuItem("2818");
                                this._Object = new ClsOutgoingPayment();
                                ((ClsOutgoingPayment)this._Object).FormDefault(_CardCode, _Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_DocNum", 0).ToString(), _Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_SODocNum", 0).ToString(), _SODocEntry, Convert.ToDouble(_Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_RfndAmt", 0).ToString()));
                            }
                            else 
                            {
                                //Utilities.Message("Not having any Advance Payment Amount", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                Utilities.Message("Refund Amount Must Be Greater Than Zero", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                        }
                        if (pVal.ItemUID == "1" && (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                        {
                            string _DocDate = ((SAPbouiCOM.EditText)_Form.Items.Item("45").Specific).Value;
                            if(_DocDate!="")
                            {
                                Utilities.ExecuteSQL(ref _RS, "update \"@BOOKINGCRM\" SET \"U_Status\" = 'Closed' Where \"U_SODocNum\" = '" + _Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_SODocNum", 0).ToString() + "'");

                                SAPbobsCOM.SalesOpportunities _SalesOpportunity = (SAPbobsCOM.SalesOpportunities)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesOpportunities);
                                if (_SalesOpportunity.GetByKey(Convert.ToInt16(_Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_OpporNo", 0).ToString())))
                                {
                                    int RetVal = 0;
                                    _SalesOpportunity.Status = BoSoOsStatus.sos_Missed;
                                    RetVal = _SalesOpportunity.Update();
                                    if (RetVal != 0)
                                    {
                                        string _errMsg = string.Empty;
                                        int _errCode = 0;
                                        Utilities.Application.Company.GetLastError(out _errCode, out _errMsg);
                                    }
                                }
                               
                            }
                            else
                            {
                                Utilities.Message("Document Date can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                           
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
                    #region et_KEY_DOWN
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "22" && pVal.InnerEvent == false && pVal.Action_Success == true && pVal.CharPressed == 9)
                        {
                            if(Convert.ToDouble(_Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_AdvAmount", 0).ToString())  > 0)
                                ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = (Convert.ToDouble(_Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_AdvAmount", 0).ToString()) - Convert.ToDouble(_Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_DedChrge", 0).ToString())).ToString();
                            else 
                                 ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = (-Convert.ToDouble(_Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_AdvAmount", 0).ToString()) - Convert.ToDouble(_Form.DataSources.DBDataSources.Item("@BOOKINGRM").GetValue("U_DedChrge", 0).ToString())).ToString();
                        }
                        break;
                    #endregion
                }
            }

            #endregion

        }

        #endregion

        #region FormDefault

        public void FormDefault(string DocNum, string CardCode, string DocEntry,string SODocEntry)
        {
            _Form.Freeze(true);
            _CardCode = CardCode;
            _SODocEntry = SODocEntry;
            _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
            _Form.DataBrowser.BrowseBy = "30";
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@BOOKINGRM");
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@BOOKINGCC1");
            _ChildDBDataSource1 = _Form.DataSources.DBDataSources.Item("@BOOKINGCC2");
            _ChildDBDataSource2 = _Form.DataSources.DBDataSources.Item("@BOOKINGCC3");
            _ChildDBDataSource3 = _Form.DataSources.DBDataSources.Item("@BOOKINGCC4");

            //_ComboBox = (SAPbouiCOM.ComboBox)Form.Items.Item("130").Specific;
            //_ComboBox.ValidValues.Add("Open", "Open");
            //_ComboBox.ValidValues.Add("Appved", "Approved");
            //_ComboBox.ValidValues.Add("UnAppved", "UnApproved");
            //_ComboBox.ValidValues.Add("Closed", "Closed");
            //_ComboBox.Select("Open", BoSearchKey.psk_ByDescription);

            Utilities.ExecuteSQL(ref _RS, "Select * From \"@BOOKINGRM\" Where \"U_CancReqNo\" = '" + DocNum + "'");
            if (_RS.RecordCount > 0)
            {
                _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                ((SAPbouiCOM.EditText)_Form.Items.Item("16").Specific).Value = DocNum;
                _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                _Form.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("112").Enabled = false;
                _Form.Items.Item("4").Enabled = false;
                _Form.Items.Item("6").Enabled = false;
                _Form.Items.Item("8").Enabled = false;
                _Form.Items.Item("38").Enabled = false;
                _Form.Items.Item("47").Enabled = false;
                _Form.Items.Item("30").Enabled = false;
                _Form.Items.Item("10").Enabled = false;
                _Form.Items.Item("12").Enabled = false;
                _Form.Items.Item("14").Enabled = false;
                _Form.Items.Item("16").Enabled = false;
                _Form.Items.Item("18").Enabled = false;
                _Form.Items.Item("20").Enabled = false;
                _Form.Items.Item("40").Enabled = false;
                _Form.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("43").Enabled = false;
                _Form.Items.Item("27").Enabled = false;
                _Form.Items.Item("28").Enabled = false;
                _Form.Items.Item("32").Enabled = false;
                _Form.Items.Item("36").Enabled = false;

                //string _SQL3 = "select \"USER_CODE\",\"U_NAME\",T1.\"Name\" as \"Department\" from OUSR T0 inner join \"OUDP\" T1 ON T1.\"Code\" = T0.\"Department\" where \"U_AuthCanc\" = 'Yes'";

                //Utilities.ExecuteSQL(ref _RS3, _SQL3);
                //for (int i = 0; i < _RS3.RecordCount; i++)
                //{
                //    if (((SAPbouiCOM.EditText)_Matrix3.Columns.Item("V_5").Cells.Item(i + 1).Specific).Value != Utilities.Application.Company.UserName)
                //    {
                //        _Matrix3.CommonSetting.SetRowEditable(i + 1, false);
                //    }
                //    _RS3.MoveNext();
                //}
            }
            else
            {

                _Matrix.AddRow(1, -1);

                string _SQL = "select \"U_CustCode\",\"U_CustName\",\"U_ContctPrsn\",\"U_AdvAmount\",\"U_PymntTrm\",\"U_PrepBy\",\"U_Status\",\"U_DocNum\",\"U_DocEntry\",\"U_SODocNum\",\"U_SoDate\",\"U_CancReqDt\",\"U_CancResion\",\"U_BookedBy\",\"U_OpporNo\",\"U_OpporNm\" from \"@BOOKINGCRM\" where \"U_DocNum\" = '" + DocNum + "'";
                Utilities.ExecuteSQL(ref _RS, _SQL);

                ((SAPbouiCOM.EditText)_Form.Items.Item("112").Specific).Value = _RS.Fields.Item("U_CustCode").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value = _RS.Fields.Item("U_CustName").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value = _RS.Fields.Item("U_ContctPrsn").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("8").Specific).Value = _RS.Fields.Item("U_AdvAmount").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("38").Specific).Value = _RS.Fields.Item("U_Status").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("30").Specific).Value = Utilities.getMaxColumnValueNum("@BOOKINGRM", "U_DocNum");
                ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = _RS.Fields.Item("U_SODocNum").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = Convert.ToDateTime(_RS.Fields.Item("U_SoDate").Value.ToString()).ToString("yyyyMMdd");
                ((SAPbouiCOM.EditText)_Form.Items.Item("14").Specific).Value = Convert.ToDateTime(_RS.Fields.Item("U_CancReqDt").Value.ToString()).ToString("yyyyMMdd");
                ((SAPbouiCOM.EditText)_Form.Items.Item("16").Specific).Value = _RS.Fields.Item("U_DocNum").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("18").Specific).Value = _RS.Fields.Item("U_CancResion").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("20").Specific).Value = _RS.Fields.Item("U_OpporNo").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("40").Specific).Value = _RS.Fields.Item("U_OpporNm").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("43").Specific).Value = Utilities.Application.Company.UserName;

                _Form.Items.Item("112").Enabled = false;
                _Form.Items.Item("4").Enabled = false;
                _Form.Items.Item("6").Enabled = false;
                _Form.Items.Item("8").Enabled = false;
                _Form.Items.Item("38").Enabled = false;
                _Form.Items.Item("47").Enabled = false;
                _Form.Items.Item("30").Enabled = false;
                _Form.Items.Item("10").Enabled = false;
                _Form.Items.Item("12").Enabled = false;
                _Form.Items.Item("14").Enabled = false;
                _Form.Items.Item("16").Enabled = false;
                _Form.Items.Item("18").Enabled = false;
                _Form.Items.Item("20").Enabled = false;
                _Form.Items.Item("40").Enabled = false;
                _Form.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("43").Enabled = false;
                _Form.Items.Item("27").Enabled = false;
                _Form.Items.Item("28").Enabled = false;
                _Form.Items.Item("32").Enabled = false;
                _Form.Items.Item("36").Enabled = false;


                string _SQL1 = "select \"U_ItemCode\",\"U_ItmDscrp\",\"U_Quantity\",\"U_Price\",\"U_Whs\",\"U_Project\",\"U_TaxCode\", " +
                                " \"U_Total\",\"U_ChasisNo\",\"U_EngineNo\",\"U_ModealYr\",\"U_RegNo\",\"U_KeyNo\" from \"@BOOKINGCRM\" T0 " +
                                 " inner join \"@BOOKINGCRC1\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.\"U_DocNum\" = '" + DocNum + "' ";

                Utilities.ExecuteSQL(ref _RS1, _SQL1);
                for (int i = 0; i < _RS1.RecordCount; i++)
                {
                    _Matrix.FlushToDataSource();
                    _ChildDBDataSource.SetValue("U_ItemCode", i, _RS1.Fields.Item("U_ItemCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ItmDscrp", i, _RS1.Fields.Item("U_ItmDscrp").Value.ToString());
                    _ChildDBDataSource.SetValue("U_Quantity", i, _RS1.Fields.Item("U_Quantity").Value.ToString());
                    _ChildDBDataSource.SetValue("U_Price", i, _RS1.Fields.Item("U_Price").Value.ToString());
                    _ChildDBDataSource.SetValue("U_Whs", i, _RS1.Fields.Item("U_Whs").Value.ToString());
                    _ChildDBDataSource.SetValue("U_Project", i, _RS1.Fields.Item("U_Project").Value.ToString());
                    _ChildDBDataSource.SetValue("U_TaxCode", i, _RS1.Fields.Item("U_TaxCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_Total", i, _RS1.Fields.Item("U_Total").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ChasisNo", i, _RS1.Fields.Item("U_ChasisNo").Value.ToString());
                    _ChildDBDataSource.SetValue("U_EngineNo", i, _RS1.Fields.Item("U_EngineNo").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ModealYr", i, _RS1.Fields.Item("U_ModealYr").Value.ToString());
                    _ChildDBDataSource.SetValue("U_RegNo", i, _RS1.Fields.Item("U_RegNo").Value.ToString());
                    _ChildDBDataSource.SetValue("U_KeyNo", i, _RS1.Fields.Item("U_KeyNo").Value.ToString());
                    _Matrix.LoadFromDataSource();
                    _RS1.MoveNext();
                }

                _Matrix1 = (SAPbouiCOM.Matrix)Form.Items.Item("28").Specific;
                string _SQL2 = "select \"U_DocType\", T1.\"U_DocNum\",\"U_DocDate\",\"U_TotAmoun\",\"U_CheqNo\",\"U_CheqDt\",\"U_BankNm\", " +
                                " \"U_CheqAmt\",\"U_BnkTrnsA\",\"U_TransDt\",\"U_CrdCrAmt\",\"U_CashAmt\",\"U_EntrBy\" from \"@BOOKINGCRM\" T0 " +
                                 " inner join \"@BOOKINGCRC2\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.\"U_DocNum\" = '" + DocNum + "'";

                Utilities.ExecuteSQL(ref _RS2, _SQL2);

                for (int i = 0; i < _RS2.RecordCount; i++)
                {
                    _Matrix1.AddRow(1, -1);
                    _Matrix1.FlushToDataSource();
                    _ChildDBDataSource1.SetValue("U_DocType", i, _RS2.Fields.Item("U_DocType").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_DocNum", i, _RS2.Fields.Item("U_DocNum").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_DocDate", i, Convert.ToDateTime(_RS2.Fields.Item("U_DocDate").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource1.SetValue("U_TotAmoun", i, _RS2.Fields.Item("U_TotAmoun").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_CheqNo", i, _RS2.Fields.Item("U_CheqNo").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_CheqDt", i, Convert.ToDateTime(_RS2.Fields.Item("U_CheqDt").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource1.SetValue("U_BankNm", i, _RS2.Fields.Item("U_BankNm").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_CheqAmt", i, _RS2.Fields.Item("U_CheqAmt").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_BnkTrnsA", i, _RS2.Fields.Item("U_BnkTrnsA").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_TransDat", i, Convert.ToDateTime(_RS2.Fields.Item("U_TransDt").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource1.SetValue("U_CrdCrAmt", i, _RS2.Fields.Item("U_CrdCrAmt").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_CashAmt", i, _RS2.Fields.Item("U_CashAmt").Value.ToString());
                    _ChildDBDataSource1.SetValue("U_EntrBy", i, _RS2.Fields.Item("U_EntrBy").Value.ToString());
                    _RS2.MoveNext();
                    _Matrix1.LoadFromDataSource();
                }

                _Matrix2 = (SAPbouiCOM.Matrix)Form.Items.Item("32").Specific;
                string _SQL3 = "select \"U_StartDt\",\"U_EndDt\",\"U_FolupTyp\",\"U_FolupSta\",\"U_Remarks\",\"U_FolupBy\",\"U_CreateAc\", " +
                                    " \"U_AsignTo\",\"U_AsignBy\",\"U_ActStDat\",\"U_ActEndDt\",\"U_ActNo\",\"U_ActStatu\" from \"@BOOKINGCRM\" T0 " +
                                     " inner join \"@BOOKINGCRC3\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.\"U_DocNum\" = '" + DocNum + "'";

                Utilities.ExecuteSQL(ref _RS3, _SQL3);
                for (int i = 0; i < _RS3.RecordCount; i++)
                {
                    _Matrix2.AddRow(1, -1);
                    _Matrix2.FlushToDataSource();

                    _ChildDBDataSource2.SetValue("U_StartDt", i, Convert.ToDateTime(_RS3.Fields.Item("U_StartDt").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource2.SetValue("U_EndDt", i, Convert.ToDateTime(_RS3.Fields.Item("U_EndDt").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource2.SetValue("U_FolupTyp", i, _RS3.Fields.Item("U_FolupTyp").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_FolupSta", i, _RS3.Fields.Item("U_FolupSta").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_Remarks", i, _RS3.Fields.Item("U_Remarks").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_FolupBy", i, _RS3.Fields.Item("U_FolupBy").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_CreateAc", i, _RS3.Fields.Item("U_CreateAc").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_AsignTo", i, _RS3.Fields.Item("U_AsignTo").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_AsignBy", i, _RS3.Fields.Item("U_AsignBy").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_ActStDat", i, Convert.ToDateTime(_RS3.Fields.Item("U_ActStDat").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource2.SetValue("U_ActEndDt", i, Convert.ToDateTime(_RS3.Fields.Item("U_ActEndDt").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource2.SetValue("U_ActNo", i, _RS3.Fields.Item("U_ActNo").Value.ToString());
                    _ChildDBDataSource2.SetValue("U_ActStatu", i, _RS3.Fields.Item("U_ActStatu").Value.ToString());

                    _RS3.MoveNext();
                    _Matrix2.LoadFromDataSource();
                }
                //_ComboBox = (SAPbouiCOM.ComboBox)_Matrix3.Columns.Item("V_4").Cells.Item(_Matrix.VisualRowCount).Specific;
                //_ComboBox.Select("Pending", SAPbouiCOM.BoSearchKey.psk_ByDescription);
                _Matrix3 = (SAPbouiCOM.Matrix)Form.Items.Item("36").Specific;
                string _SQL4 = "select \"U_Departm\",\"U_UsrCode\",\"U_UsrName\",T1.\"U_Status\",\"U_Decision\",\"U_Remarks\",\"U_Date\", " +
                                " \"U_Time\" from \"@BOOKINGCRM\" T0 " +
                                 " inner join \"@BOOKINGCRC4\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.\"U_DocNum\" = '" + DocNum + "'";

                Utilities.ExecuteSQL(ref _RS4, _SQL4);
                for (int i = 0; i < _RS4.RecordCount; i++)
                {
                    if (i > 0)
                    {
                        _Matrix3.AddRow(1, -1);
                    }
                    _Matrix3.FlushToDataSource();
                    _ChildDBDataSource3.SetValue("U_Departme", i, _RS4.Fields.Item("U_Departm").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_UsrCode", i, _RS4.Fields.Item("U_UsrCode").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_UsrName", i, _RS4.Fields.Item("U_UsrName").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_Status", i, _RS4.Fields.Item("U_Status").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_Decision", i, _RS4.Fields.Item("U_Decision").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_Remarks", i, _RS4.Fields.Item("U_Remarks").Value.ToString());
                    _ChildDBDataSource3.SetValue("U_Date", i, Convert.ToDateTime(_RS4.Fields.Item("U_Date").Value.ToString()).ToString("yyyyMMdd"));
                    _ChildDBDataSource3.SetValue("U_Time", i, _RS4.Fields.Item("U_Time").Value.ToString());
                    _RS4.MoveNext();
                    _Matrix3.LoadFromDataSource();

                }

            }
            _Form.Items.Item("25").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            _Form.Freeze(false);
        }

        #endregion

    }
}
