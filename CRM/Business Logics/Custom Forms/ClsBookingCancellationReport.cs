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
using SAPbobsCOM;
using SAPbouiCOM;

namespace CRM
{
    class ClsBookingCancellationReport : Base
    {

        #region Variable Declaration

        private static HanaConnection _conn;
        private static HanaDataAdapter _da;
        private static HanaParameter _parm1;
        private static HanaParameter _parm2;
        private SAPbouiCOM.Grid _Grid, _Grid1;
        private SAPbouiCOM.DataTable _GridDataTable, _GridDataTable1;
        SAPbouiCOM.OptionBtn _OptionBtn;
        private SAPbouiCOM.RowHeaders _RowHeaders, _RowHeaders1;
        private string DeliveryNo = string.Empty;
        private string sInvoiceNo = string.Empty;
        string strDel = string.Empty;
        public static string _CheckVal = string.Empty;
        public static string StrFromDate1 = string.Empty;
        public static string StrToDate1 = string.Empty;
        public static string StrCustomer1 = string.Empty;
        public static string StrStatus1 = string.Empty;
        public static string _StrPostingDate = string.Empty;
        public static string _StrStatus = string.Empty;
        private static int _errors = 0;
        public string BPLId = string.Empty, FinalBPLId = string.Empty;
        public string _QuerySelection = string.Empty;
        int j = 0;
        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;
        StringBuilder result = new StringBuilder();
        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsBookingCancellationReport()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsBookingCancellationReport()
        {
        }

        #endregion

        #region ITEM EVENT

        public override void Item_Event(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            #region BEFORE ACTION

            if (pVal.Before_Action)
            {
                switch (pVal.EventType)
                {
                    #region ITEM PRESSED

                    case BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "6")
                        {
                            j = 0;
                            FinalBPLId = string.Empty;
                            _Grid = (SAPbouiCOM.Grid)Form.Items.Item("4").Specific;
                            for (int i = 0; i < _Grid.Rows.Count; i++)
                            {
                                string CheckStatus = _Grid.DataTable.Columns.Item("Active").Cells.Item(i).Value.ToString();
                                if (CheckStatus == "Y")
                                {
                                    BPLId = "'" + _Grid.DataTable.Columns.Item("BPLId").Cells.Item(i).Value.ToString() + "'";
                                    if (j > 0)
                                        FinalBPLId = FinalBPLId.ToString() + "," + BPLId.ToString();
                                    else
                                        FinalBPLId = BPLId;

                                    j++;
                                }
                            }
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtSalPer").Specific).Value != "")
                            {
                                result.Append(" And T7.\"SlpName\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtSalPer").Specific).Value + "'");
                            }
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtSchme").Specific).Value != "")
                            {
                                result.Append(" And T6.\"PrjCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtSchme").Specific).Value + "'");
                            }
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value != "")
                            {
                                result.Append(" And TM1.\"U_Model\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value + "'");
                            }

                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtVarnt").Specific).Value != "")
                            {
                                result.Append(" AND TM1.\"U_Variant\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtVarnt").Specific).Value + "'");
                            }

                        }
                        break;

                    #endregion

                    #region key down event
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "txtMdl" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value == "")
                            {
                                string Query = "select distinct \"U_Model\" as \"Model\" from OITM where \"U_Model\" is not null";
                                this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Model";
                                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, Utilities.Hana_ADOConnection(Query), false);

                            }
                        }

                        if (pVal.ItemUID == "txtVarnt" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtVarnt").Specific).Value == "" && ((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value !="")
                            {
                                //string Query = "select distinct  SUBSTRING(T1.\"ItemName\", (Locate(T1.\"ItemName\", '-', 1)) + 1, (Locate(T1.\"ItemName\", '-', -2)) -1) \"Varient\"" +
                                //                "from \"OITM\" T1 where \"U_Branch\" = 'Vehicle' and SUBSTRING(T1.\"ItemName\", (Locate(T1.\"ItemName\", '-', 1)) + 1, (Locate(T1.\"ItemName\", '-', -2)) - 1)<>'' ";

                                string Query = "select distinct \"U_Variant\" from OITM where \"U_Model\" = '" + (((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value) + "'"; this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Varient";
                                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, Utilities.Hana_ADOConnection(Query), false);

                            }
                        }
                       
                        if (pVal.ItemUID == "txtSchme" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtSchme").Specific).Value == "")
                            {
                                string Query = "select \"PrjCode\" as \"Scheme Code\",\"PrjName\" as \"Scheme Name\" from OPRJ";
                                this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Scheme";
                                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, Utilities.Hana_ADOConnection(Query), false);

                            }
                        }

                        if (pVal.ItemUID == "txtSalPer" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtSalPer").Specific).Value == "")
                            {
                                string Query = "select \"SlpName\" as \"Sales Person Name\" from OSLP";
                                this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Sales Person";
                                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, Utilities.Hana_ADOConnection(Query), false);

                            }
                        }

                        break;
                        #endregion


                }
            }
            else
            {
                switch (pVal.EventType)
                {
                    #region ITEM PRESSED

                    case BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "6" && pVal.Action_Success)
                        {
                            FilterGridValue();
                        }
                        if (pVal.ItemUID == "8")
                        {
                            _Grid1.Rows.ExpandAll();
                        }
                        if (pVal.ItemUID == "9")
                        {
                            _Grid1.Rows.CollapseAll();
                        }
                        break;

                        #endregion


                }
            }

            #endregion

        }

        #endregion

        #region CHOOSE FROM LIST EVENT

        public override void ChooseFromList_Event()
        {
            try
            {

                if (this._CFLEvent.ItemUID == "txtMdl")
                {
                    _Form.Freeze(true);
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                    _Form.Freeze(false);
                }
                if (this._CFLEvent.ItemUID == "txtVarnt")
                {
                    _Form.Freeze(true);
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtVarnt").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                    _Form.Freeze(false);
                }
               
                if (this._CFLEvent.ItemUID == "txtSchme")
                {
                    _Form.Freeze(true);
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtSchme").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                    _Form.Freeze(false);
                }
                if (this._CFLEvent.ItemUID == "txtSalPer")
                {
                    _Form.Freeze(true);
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtSalPer").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                    _Form.Freeze(false);
                }

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.ToString());
            }
        }

        #endregion

        #region FormDefault

        public void FormDefault()
        {
            _Form.Freeze(true);
            _OptionBtn = (SAPbouiCOM.OptionBtn)_Form.Items.Item("13").Specific;
            _OptionBtn.GroupWith("12");
            _OptionBtn.Selected = true;

            var now = DateTime.Now;
            string startOfMonth = new DateTime(now.Year, now.Month, 1).ToString("yyyyMMdd");
            int DaysInMonth = DateTime.DaysInMonth(now.Year, now.Month);
            string lastDay = DateTime.Now.ToString("yyyyMMdd");

            ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value = startOfMonth.ToString();
            ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value = lastDay.ToString();

            _GridDataTable = null;
            _Form.DataSources.DataTables.Add("dtReviewList");
            _Grid = (SAPbouiCOM.Grid)Form.Items.Item("4").Specific;
            _GridDataTable = _Form.DataSources.DataTables.Item("dtReviewList");

            _GridDataTable1 = null;
            _Form.DataSources.DataTables.Add("dtReviewList1");
            _Grid1 = (SAPbouiCOM.Grid)Form.Items.Item("5").Specific;
            _GridDataTable1 = _Form.DataSources.DataTables.Item("dtReviewList1");

            string strSQL = string.Empty, strSQL1 = string.Empty;

            strSQL = "select 'Y' as \"Active\",\"BPLId\",(select \"CompnyName\" from OADM) as \"Company Name\",\"BPLName\" as \"Branch Name\" from OBPL where \"Disabled\" = 'N' order by \"BPLId\" ";

            _GridDataTable.ExecuteQuery(strSQL);
            _Grid.DataTable = _GridDataTable;

            _RowHeaders = _Grid.RowHeaders;
            _RowHeaders.TitleObject.Caption = "#";
            //_Grid.RowHeaders.Width = 20;
            _Grid.Columns.Item("Active").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

            for (int i = 0; i < _Grid.Rows.Count; i++)
            {
                _RowHeaders.SetText(i, (i + 1).ToString());

            }
            _Grid.AutoResizeColumns();
            _Grid.Columns.Item(1).Editable = false;
            _Grid.Columns.Item(2).Editable = false;
            _Grid.Columns.Item(3).Editable = false;

            /******************/
            strSQL1 = "select distinct OP.\"OpenDate\" \"Enquiry Date\",BO.\"U_OpporNo\" \"Enquiry Number\",BO.\"U_OpporNm\" \"Enquiry Vehicle\", T0.\"DocNum\" \"Booking Number\",T0.\"DocDate\" \"Booked Date\",T0.\"CardCode\" \"Customer Code\", T0.\"CardName\" \"Customer Name\", RD.\"Address\"," +
                        "T4.\"Cellular\" \"Cell Phone\"," +
                        //"SUBSTRING(T1.\"Dscription\", 1, (Locate(T1.\"Dscription\", '-', 1)) - 1) \"Model\"," +
                        //"SUBSTRING(T1.\"Dscription\", (Locate(T1.\"Dscription\", '-', 1)) + 1, (Locate(T1.\"Dscription\", '-', -2)) - 1) \"Varient\"," +
                        //"SUBSTRING(TM1.\"ItemName\", (Locate(TM1.\"ItemName\", '-', -2)) + 1) \"Color\", " +
                         " TM1.\"U_Model\" \"Model\",TM1.\"U_Variant\" \"Variant\"," +
                        "T1.\"Dscription\" \"Vehicle\",BO1.\"U_Whs\" \"Ware House\"," +
                        "T0.\"BPLName\" \"Branch\", T6.\"PrjName\" \"Scheme\", BO1.\"U_Price\" \"Price\",ITM.\"Price\" -T0.\"DiscSum\" \"Deal Value\"," +
                        "T0.\"DocTotal\" \"Document Total\",BO.\"U_AdvAmount\" \"Advance Amount\",BG.\"U_PymntTrm\" \"Payment Method\",BO.\"U_RfndAmt\" \"Refund Amount\"," +
                        "BO.\"U_DocNum\" \"Booking Cancellation Number\",BO.\"U_DocDate\" \"Booking Cancellation  Date\"," +
                        "BO.\"U_CancResion\" \"Booking Cancellation Reason\",BO.\"U_Status\" \"Booking Cancellation Status\"," +
                        "BO.\"U_BookedBy\" \"Booked By\"  from \"@BOOKINGRM\" BO " +
                        "INNER JOIN \"@BOOKINGCC1\" BO1 ON BO.\"DocEntry\" = BO1.\"DocEntry\" " +
                        "INNER JOIN \"@BOOKINGCRM\" BG ON BO.\"DocEntry\" = BG.\"DocEntry\" " +
                        "INNER JOIN \"ORDR\" T0 ON T0.\"DocEntry\" = BG.\"U_DocEntry\" and T0.\"DocNum\"=BG.\"U_SODocNum\" " +
                        "INNER JOIN \"RDR1\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                        "INNER JOIN \"OCRD\" T4 ON T4.\"CardCode\" = T0.\"CardCode\" " +
                        "LEFT OUTER JOIN \"CRD1\" RD ON RD.\"CardCode\" = T4.\"CardCode\" " +
                        "INNER JOIN \"OCRG\" T5 ON T5.\"GroupCode\" = T4.\"GroupCode\" " +
                        "LEFT OUTER JOIN \"OPRJ\" T6 ON T0.\"Project\" = T6.\"PrjCode\" " +
                        "LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = T0.\"SlpCode\" " +
                        "LEFT OUTER JOIN \"OPR1\" T8 ON T8.\"DocId\" = T1.\"DocEntry\" and T8.\"ObjType\" = T1.\"ObjType\""+  
                        "INNER JOIN \"OOPR\" OP on T8.\"OpprId\" = OP.\"OpprId\""+
                        "LEFT OUTER JOIN OOST ST ON T8.\"Step_Id\" = ST.\"Num\"  and T8.\"Line\" = (select  max(L.\"Line\") from \"OPR1\" L where L.\"OpprId\" = T8.\"OpprId\")  " +
                        "INNER JOIN \"ITM1\" TM on TM.\"ItemCode\" = T1.\"ItemCode\"  " +
                        "INNER JOIN \"OITM\" TM1 on TM1.\"ItemCode\" = TM.\"ItemCode\"  " +
                        "LEFT JOIN \"ITM1\" ITM on ITM.\"ItemCode\" = T1.\"ItemCode\" and ITM.\"PriceList\" = '2' " +
                        "INNER JOIN \"OITB\" TB on TB.\"ItmsGrpCod\"= TM1.\"ItmsGrpCod\" " +
                       "where BO.\"U_DocDate\" between (SELECT ADD_MONTHS(NEXT_DAY(LAST_DAY(CURRENT_DATE)),-1) FROM DUMMY) AND (SELECT LAST_DAY(CURRENT_DATE) FROM DUMMY) order by BO.\"U_DocDate\" ";

          
            //"where T0.\"BPLId\" in (" + FinalBPLId + ") AND T0.\"DocDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + " order by T0.\"DocDate\" ";// --and TM."PriceList" = '2'";

            _GridDataTable1.ExecuteQuery(strSQL1);
            _Grid1.DataTable = _GridDataTable1;

            _RowHeaders1 = _Grid1.RowHeaders;
            _RowHeaders1.TitleObject.Caption = "#";
            ////_Grid.RowHeaders.Width = 20;
            //_Grid1.Columns.Item("Active").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

            for (int i = 0; i < _Grid1.Rows.Count; i++)
            {
                _RowHeaders1.SetText(i, (i + 1).ToString());

            }
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JANUARY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("FEBRUARY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("MARCH")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("APRIL")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("MAY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JUNE")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JULY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("AUGUST")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("SEPTEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("OCTOBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("NOVEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("DECEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Average Sale")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            _Grid1.Columns.Item(0).TitleObject.Sortable = true;
            _Grid1.Columns.Item(6).TitleObject.Sortable = true;
            _Grid1.AutoResizeColumns();
            _Grid1.Columns.Item(0).Editable = false;
            _Grid1.Columns.Item(1).Editable = false;
            _Grid1.Columns.Item(2).Editable = false;
            _Grid1.Columns.Item(3).Editable = false;
            _Grid1.Columns.Item(4).Editable = false;
            _Grid1.Columns.Item(5).Editable = false;
            _Grid1.Columns.Item(6).Editable = false;
            _Grid1.Columns.Item(7).Editable = false;
            _Grid1.Columns.Item(8).Editable = false;
            _Grid1.Columns.Item(9).Editable = false;
            _Grid1.Columns.Item(10).Editable = false;
            _Grid1.Columns.Item(11).Editable = false;
            _Grid1.Columns.Item(12).Editable = false;
            _Grid1.Columns.Item(13).Editable = false;
            _Grid1.Columns.Item(14).Editable = false;
            _Grid1.Columns.Item(15).Editable = false;
            _Grid1.Columns.Item(16).Editable = false;
            _Grid1.Columns.Item(17).Editable = false;
            _Grid1.Columns.Item(18).Editable = false;
            _Grid1.Columns.Item(19).Editable = false;
            _Grid1.Columns.Item(20).Editable = false;
            _Grid1.Columns.Item(21).Editable = false;
            _Grid1.Columns.Item(22).Editable = false;
            _Grid1.Columns.Item(23).Editable = false;
            _Grid1.Columns.Item(24).Editable = false;
            _Grid1.Columns.Item(25).Editable = false;
            //_Grid1.Columns.Item(26).Editable = false;
            /******************/

            _Form.Freeze(false);

        }

        #endregion

        #region FilterGridValue

        public void FilterGridValue()
        {
            //_GridDataTable1 = null;
            //_Form.DataSources.DataTables.Add("dtReviewList1");
            //_Grid1 = (SAPbouiCOM.Grid)Form.Items.Item("5").Specific;
            //_GridDataTable1 = _Form.DataSources.DataTables.Item("dtReviewList1");

            this._GridDataTable1.Rows.Clear();
            if (FinalBPLId != "")
            {
                string strSQL = string.Empty, strSQL1 = string.Empty;
                _OptionBtn = (SAPbouiCOM.OptionBtn)_Form.Items.Item("13").Specific;
                if (_OptionBtn.Selected == true)
                {
                    strSQL1 = "select distinct OP.\"OpenDate\" \"Enquiry Date\",BO.\"U_OpporNo\" \"Enquiry Number\",BO.\"U_OpporNm\" \"Enquiry Vehicle\", T0.\"DocNum\" \"Booking Number\",T0.\"DocDate\" \"Booked Date\",T0.\"CardCode\" \"Customer Code\", T0.\"CardName\" \"Customer Name\", RD.\"Address\"," +
                        "T4.\"Cellular\" \"Cell Phone\"," +
                         //"SUBSTRING(T1.\"Dscription\", 1, (Locate(T1.\"Dscription\", '-', 1)) - 1) \"Model\"," +
                         //"SUBSTRING(T1.\"Dscription\", (Locate(T1.\"Dscription\", '-', 1)) + 1, (Locate(T1.\"Dscription\", '-', -2)) - 1) \"Varient\"," +
                         //"SUBSTRING(TM1.\"ItemName\", (Locate(TM1.\"ItemName\", '-', -2)) + 1) \"Color\", " +
                         " TM1.\"U_Model\" \"Model\",TM1.\"U_Variant\" \"Variant\",TM1.\"U_Colour\" \"Color\"," +
                        "T1.\"Dscription\" \"Vehicle\",BO1.\"U_Whs\" \"Ware House\"," +
                        "T0.\"BPLName\" \"Branch\", T6.\"PrjName\" \"Scheme\", BO1.\"U_Price\" \"Price\",ITM.\"Price\" -T0.\"DiscSum\" \"Deal Value\"," +
                        "T0.\"DocTotal\" \"Document Total\",BO.\"U_AdvAmount\" \"Advance Amount\",BG.\"U_PymntTrm\" \"Payment Method\",BO.\"U_RfndAmt\" \"Refund Amount\"," +
                        "BO.\"U_DocNum\" \"Booking Cancellation Number\",BO.\"U_DocDate\" \"Booking Cancellation  Date\"," +
                        "BO.\"U_CancResion\" \"Booking Cancellation Reason\",BO.\"U_Status\" \"Booking Cancellation Status\"," +
                        "BO.\"U_BookedBy\" \"Booked By\"  from \"@BOOKINGRM\" BO " +
                        "INNER JOIN \"@BOOKINGCC1\" BO1 ON BO.\"DocEntry\" = BO1.\"DocEntry\" " +
                        "INNER JOIN \"@BOOKINGCRM\" BG ON BO.\"DocEntry\" = BG.\"DocEntry\" " +
                        "INNER JOIN \"ORDR\" T0 ON T0.\"DocEntry\" = BG.\"U_DocEntry\" and T0.\"DocNum\"=BG.\"U_SODocNum\" " +
                        "INNER JOIN \"RDR1\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                        "INNER JOIN \"OCRD\" T4 ON T4.\"CardCode\" = T0.\"CardCode\" " +
                        "LEFT OUTER JOIN \"CRD1\" RD ON RD.\"CardCode\" = T4.\"CardCode\" " +
                        "INNER JOIN \"OCRG\" T5 ON T5.\"GroupCode\" = T4.\"GroupCode\" " +
                        "LEFT OUTER JOIN \"OPRJ\" T6 ON T0.\"Project\" = T6.\"PrjCode\" " +
                        "LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = T0.\"SlpCode\" " +
                        "LEFT OUTER JOIN \"OPR1\" T8 ON T8.\"DocId\" = T1.\"DocEntry\" and T8.\"ObjType\" = T1.\"ObjType\"" +
                        "INNER JOIN \"OOPR\" OP on T8.\"OpprId\" = OP.\"OpprId\"" +
                        "LEFT OUTER JOIN OOST ST ON T8.\"Step_Id\" = ST.\"Num\"  and T8.\"Line\" = (select  max(L.\"Line\") from \"OPR1\" L where L.\"OpprId\" = T8.\"OpprId\")  " +
                        "INNER JOIN \"ITM1\" TM on TM.\"ItemCode\" = T1.\"ItemCode\"  " +
                        "INNER JOIN \"OITM\" TM1 on TM1.\"ItemCode\" = TM.\"ItemCode\"  " +
                        "LEFT JOIN \"ITM1\" ITM on ITM.\"ItemCode\" = T1.\"ItemCode\" and ITM.\"PriceList\" = '2' " +
                        "INNER JOIN \"OITB\" TB on TB.\"ItmsGrpCod\"= TM1.\"ItmsGrpCod\" "  +
                        "where T0.\"BPLId\" in (" + FinalBPLId + ") AND BO.\"U_DocDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + " order by BO.\"U_DocDate\" ";// --and TM."PriceList" = '2'";

                }
                else
                {

                    strSQL1 = "select A.\"ItmsGrpNam\" \"Model\" , C.\"Total Booking Cancellation\" " +
                   "from(select count(\"Total Enquiry\") \"Total Enquiry\", \"ItmsGrpNam\" " +
                   "from(select distinct T.\"OpprId\" \"Total Enquiry\", TB.\"ItmsGrpNam\" \"ItmsGrpNam\" from \"OOPR\" T  " +
                   "LEFT OUTER JOIN \"OITM\" TM1 on SUBSTRING(TM1.\"ItemName\", 1, (Locate(TM1.\"ItemName\", '-', -2)) - 1) = T.\"Name\" " +
                   "INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\"  " +
                   "LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = T.\"SlpCode\"  " +
                   "LEFT JOIN  \"OPRJ\" T6 ON T.\"PrjCode\" = T6.\"PrjCode\" " +
                   "INNER JOIN \"OUSR\" T0 ON T.\"UserSign\" = T0.\"USERID\" " +
                   "INNER JOIN \"USR6\" SR on  SR.\"UserCode\" = T0.\"USER_CODE\" where  SR.\"BPLId\" in (" + FinalBPLId + ") and T.\"OpenDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' "+ result +"   )group by \"ItmsGrpNam\") A " +
                   "Left JOIN(select sum(T.\"Quantity\") \"Total Booking\", TB.\"ItmsGrpNam\" from \"QUT1\" T " +
                   "INNER JOIN  \"OQUT\" DR ON DR.\"DocEntry\" = T.\"DocEntry\" " +
                   "LEFT OUTER JOIN \"OITM\" TM1 on T.\"ItemCode\" = TM1.\"ItemCode\" " +
                   "INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\" " +
                   "LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = T.\"SlpCode\" " +
                   "Left JOIN \"OPRJ\" T6 ON DR.\"Project\" = T6.\"PrjCode\" where DR.\"BPLId\" in (" + FinalBPLId + ") and DR.\"DocDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' "+ result +"   group by TB.\"ItmsGrpNam\") B ON A.\"ItmsGrpNam\" = B.\"ItmsGrpNam\" " +
                   "Left JOIN(select sum(T.\"Quantity\") \"Total Booking\",TB.\"ItmsGrpNam\" from \"RDR1\" T " +
                   "INNER JOIN ORDR DR ON DR.\"DocEntry\" = T.\"DocEntry\"  " +
                   "LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = DR.\"SlpCode\" " +
                   "Left JOIN \"OPRJ\" T6 ON T.\"Project\" = T6.\"PrjCode\" " +
                   "LEFT OUTER JOIN \"OITM\" TM1 on T.\"ItemCode\" = TM1.\"ItemCode\" " +
                   "INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\" where DR.\"BPLId\" in (" + FinalBPLId + ") and DR.\"DocDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + "    group by TB.\"ItmsGrpNam\") Q ON A.\"ItmsGrpNam\" = Q.\"ItmsGrpNam\" " +
                   "LEFT JOIN(select sum(\"Total Booking Cancellation\") \"Total Booking Cancellation\",\"ItmsGrpNam\" " +
                   "from(select distinct BO1.\"U_Quantity\" \"Total Booking Cancellation\",TM1.\"ItemName\" ,BO.\"U_CancReqDt\",TB.\"ItmsGrpNam\" from \"@BOOKINGCC1\" BO1 " +
                   "INNER JOIN \"@BOOKINGRM\" BO ON BO.\"DocEntry\" = BO1.\"DocEntry\" " +
                   "INNER JOIN \"@BOOKINGCRM\" BG ON BO.\"DocEntry\" = BG.\"DocEntry\" " +
                   "INNER JOIN \"ORDR\" DR ON DR.\"DocEntry\" = BG.\"U_DocEntry\" and DR.\"DocNum\"=BG.\"U_SODocNum\" " +
                   "LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = DR.\"SlpCode\" " +
                   "LEFT OUTER JOIN \"OITM\" TM1 on BO1.\"U_ItemCode\" = TM1.\"ItemCode\" " +
                   "INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\" " +
                   "INNER JOIN \"OUSR\" T0 ON BO.\"UserSign\" = T0.\"USERID\" " +
                   "INNER JOIN \"USR6\" SR on  SR.\"UserCode\" = T0.\"USER_CODE\" where SR.\"BPLId\" in (" + FinalBPLId + ") and BO.\"U_DocDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + "  ) group by \"ItmsGrpNam\")  C ON A.\"ItmsGrpNam\" = C.\"ItmsGrpNam\" ";
                }
                _GridDataTable1.ExecuteQuery(strSQL1);
                _Grid1.DataTable = _GridDataTable1;
                result.Clear();
                _RowHeaders1 = _Grid1.RowHeaders;
                _RowHeaders1.TitleObject.Caption = "#";

                for (int i = 0; i < _Grid1.Rows.Count; i++)
                {
                    _RowHeaders1.SetText(i, (i + 1).ToString());

                }
                //_Grid1.CollapseLevel = 1;

                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JANUARY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("FEBRUARY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("MARCH")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("APRIL")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("MAY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JUNE")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JULY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("AUGUST")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("SEPTEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("OCTOBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("NOVEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("DECEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Average Sale")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                _Grid1.AutoResizeColumns();
                if (_OptionBtn.Selected == true)
                {

                    _Grid1.Columns.Item(0).Editable = false;
                    _Grid1.Columns.Item(1).Editable = false;
                    _Grid1.Columns.Item(2).Editable = false;
                    _Grid1.Columns.Item(3).Editable = false;
                    _Grid1.Columns.Item(4).Editable = false;
                    _Grid1.Columns.Item(5).Editable = false;
                    _Grid1.Columns.Item(6).Editable = false;
                    _Grid1.Columns.Item(7).Editable = false;
                    _Grid1.Columns.Item(8).Editable = false;
                    _Grid1.Columns.Item(9).Editable = false;
                    _Grid1.Columns.Item(10).Editable = false;
                    _Grid1.Columns.Item(11).Editable = false;
                    _Grid1.Columns.Item(12).Editable = false;
                    _Grid1.Columns.Item(13).Editable = false;
                    _Grid1.Columns.Item(14).Editable = false;
                    _Grid1.Columns.Item(15).Editable = false;
                    _Grid1.Columns.Item(16).Editable = false;
                    _Grid1.Columns.Item(17).Editable = false;
                    _Grid1.Columns.Item(18).Editable = false;
                    _Grid1.Columns.Item(19).Editable = false;
                    _Grid1.Columns.Item(20).Editable = false;

                }
                else
                {
                    _Grid1.Columns.Item(0).Editable = false;
                    _Grid1.Columns.Item(1).Editable = false;
                    //_Grid1.Columns.Item(2).Editable = false;
                   // _Grid1.Columns.Item(3).Editable = false;
                   // _Grid1.Columns.Item(4).Editable = false;
                    //_Grid1.Columns.Item(5).Editable = false;
                    //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Enquiry")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    ///((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Quotation")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Booking")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Booking Cancellation")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    //_Grid1.Columns.Item(6).Editable = false;
                    _Form.Update();
                }

            }
        }

        #endregion

    }
}
