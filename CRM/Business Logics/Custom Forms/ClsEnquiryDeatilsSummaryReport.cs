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
    class ClsEnquiryDeatilsSummaryReport : Base
    {

        #region Variable Declaration

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
        public string BPLId = string.Empty,FinalBPLId = string.Empty;
        int j = 0;
        StringBuilder result = new StringBuilder();
        StringBuilder result1 = new StringBuilder();
        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsEnquiryDeatilsSummaryReport()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsEnquiryDeatilsSummaryReport()
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
                            for(int i = 0;i<_Grid.Rows.Count;i++)
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
                        }
                        //result----
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("txtSalPer").Specific).Value != "")
                        {
                            result.Append(" And T6.\"SlpName\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtSalPer").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("txtSchme").Specific).Value != "")
                        {
                            result.Append(" And T2.\"PrjCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtSchme").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value != "")
                        {
                            result.Append(" And TM1.\"U_Model\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value + "'");
                        }

                        if (((SAPbouiCOM.EditText)_Form.Items.Item("txtVarnt").Specific).Value != "")
                        {
                            result.Append(" AND TM1.\"U_Variant\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtVarnt").Specific).Value + "'");
                        }

                        if (((SAPbouiCOM.EditText)_Form.Items.Item("txtColor").Specific).Value != "")
                        {
                            result.Append(" And TM1.\"U_Colour\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtColor").Specific).Value + "'");
                        }
                        //-----result 1-----
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("txtSalPer").Specific).Value != "")
                        {
                            result1.Append(" And T6.\"SlpName\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtSalPer").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("txtSchme").Specific).Value != "")
                        {
                            result1.Append(" And T2.\"PrjCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtSchme").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value != "")
                        {
                            result1.Append(" And T0.\"U_Model\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value + "'");
                        }

                        if (((SAPbouiCOM.EditText)_Form.Items.Item("txtVarnt").Specific).Value != "")
                        {
                            result1.Append(" AND T0.\"U_Variant\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtVarnt").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("txtColor").Specific).Value != "")
                        {
                            result1.Append(" And TM1.\"U_Colour\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtColor").Specific).Value + "'");
                        }

                        break;

                    #endregion

                    #region Key Down

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "txtMdl" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value == "")
                            {
                                //string Query = "select distinct RTRIM(SUBSTRING(T0.\"Name\", 1 , (Locate(T0.\"Name\", '-', 1)) - 1)) \"Model\"  from OOPR T0 where SUBSTRING(T0.\"Name\", 1 , (Locate(T0.\"Name\", '-', 1)) - 1)<>'' ";
                                string Query = "select distinct \"U_Model\" as \"Model\" from OITM where \"U_Model\" is not null";
                                this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Model";
                                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, Utilities.Hana_ADOConnection(Query), false);

                            }
                        }

                        if (pVal.ItemUID == "txtVarnt" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtVarnt").Specific).Value == "")
                            {
                                //string Query = "select distinct  SUBSTRING(T0.\"Name\", (Locate(T0.\"Name\", '-', -2)) + 1) \"Varient\"  from \"OOPR\" T0 ";
                                string Query = "select distinct \"U_Variant\" from OITM where \"U_Model\" = '" + (((SAPbouiCOM.EditText)_Form.Items.Item("txtMdl").Specific).Value) + "'"; 
                                this._Object = new ClsChooseFromList();
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


                        if (pVal.ItemUID == "txtColor" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtColor").Specific).Value == "")
                            {
                                string Query = "select distinct \"U_Colour\" from OITM where \"U_Variant\" = '" + (((SAPbouiCOM.EditText)_Form.Items.Item("txtVarnt").Specific).Value) + "'";
                                this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Color";
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
                        if(pVal.ItemUID == "8")
                        {
                            _Grid1.Rows.ExpandAll();
                        }
                        if (pVal.ItemUID == "9")
                        {
                            _Grid1.Rows.CollapseAll();
                        }
                        break;

                    #endregion
                    #region Double Click
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ColUID == "RowsHeader")
                        {
                            Utilities.Message("Please wait operation is in process.........", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            string _FromDate = ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString();
                            string _ToDate = ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString();
                            string _Model = _Grid1.DataTable.Columns.Item("Model").Cells.Item(pVal.Row).Value.ToString();

                            //_Grid.CommonSetting.SetRowBackColor(count, -1);//-41
                            //count = pVal.Row + 1;
                            //_Grid1.CommonSetting.SetRowBackColor(pVal.Row + 1, 3198718);

                            this._Object = new ClsEnquirySubReport();
                            this.OpenChildForm(Constants.Forms.EnquirySubReport, ref pVal, true);
                            ((ClsEnquirySubReport)_Object).FormDefault(_FromDate, _ToDate, _Model, FinalBPLId, result);

                            Utilities.Message("Operation completed sucessfully.", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
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
                if (this._CFLEvent.ItemUID == "txtColor")
                {
                    _Form.Freeze(true);
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtColor").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
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
            //string lastDay1 = new DateTime(now.Year, now.Month, DaysInMonth).ToString("yyyyMMdd");
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
            strSQL1 = "select  distinct  T0.\"OpenDate\" \"Enquiry Date\", "+
                      " PL.\"BPLName\" \"Branch\", " +
                      " T0.\"CardName\" \"Customer Name\",T7.\"Cellular\" \"Cell Phone\", T7.\"Phone1\" \"Res Phone\",T8.\"GroupName\" \"Occupation\"," +
                      //"T0.\"Name\" \"Enquiry vehicle\", " +
                      " TM1.\"U_Model\" \"Model\",TM1.\"U_Variant\" \"Variant\"," +
                      " T0.\"UserSign\", " +
                      " T6.\"SlpName\" \"Entered By\", "+
                      " case when T0.\"Status\"='O' then 'Open' when T0.\"Status\"='C' then 'Close' when T0.\"Status\"='W' then 'Won' when T0.\"Status\"='L' then 'Lose' End \"Enquiry Status\", " +
                      " T3.\"Descript\" \"Level of Interest\" ,T4.\"Descript\" \"Enquiry Type\", T2.\"PrjName\" \"Scheme\", "+
                      " T5.\"Descript\" \"Last Stage\" from \"OOPR\" T0 "+
                      " INNER JOIN OPR1 T1 ON T0.\"OpprId\" = T1.\"OpprId\" "+
                      " LEFT OUTER JOIN OPRJ T2 ON T0.\"PrjCode\" = T2.\"PrjCode\" " +
                      " LEFT OUTER JOIN OOIR T3 ON T0.\"IntRate\" = T3.\"Num\" " +
                      " LEFT OUTER JOIN OOSR T4 ON T0.\"Source\" = T4.\"Num\" " +
                      " INNER JOIN OOST T5 ON T1.\"Step_Id\" = T5.\"Num\"  and T1.\"Line\"= (select  max(L.\"Line\") from \"OPR1\" L where L.\"OpprId\"=T0.\"OpprId\") " +
                      " LEFT JOIN  OSLP T6 ON T6.\"SlpCode\" =T0.\"SlpCode\" "+
                      " INNER JOIN OCRD T7 ON T7.\"CardCode\" =T0.\"CardCode\" "+
                      " INNER JOIN OCRG T8 ON T8.\"GroupCode\" =T0.\"CardGroup\" " +
                      "LEFT OUTER JOIN \"OUSR\" T9 ON T0.\"UserSign\" = T9.\"USERID\" " +
                      "LEFT OUTER JOIN\"USR6\" SR on  SR.\"UserCode\" = T9.\"USER_CODE\" " +
                      "LEFT OUTER JOIN \"OBPL\" PL ON PL.\"BPLId\"= SR.\"BPLId\" " +
                      "INNER JOIN \"OITM\" TM1 ON TM1.\"U_Model\"= T0.\"U_Model\" and TM1.\"U_Variant\"= T0.\"U_Variant\" " +
                      "where T0.\"OpenDate\" between (SELECT ADD_MONTHS(NEXT_DAY(LAST_DAY(CURRENT_DATE)),-1) FROM DUMMY) AND (SELECT LAST_DAY(CURRENT_DATE) FROM DUMMY) order by T0.\"OpenDate\" ";

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
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
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
            //_Grid1.Columns.Item(15).Editable = false;
            
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
                    //strSQL1 = "select  T0.\"OpenDate\" \"Enquiry Date\", " +
                    // " (select distinct PL.\"BPLName\" from \"OUSR\" OS INNER JOIN \"USR6\" R6 on R6.\"UserCode\" = OS.\"USER_CODE\" " +
                    // " INNER JOIN \"OBPL\" PL ON PL.\"BPLId\"= R6.\"BPLId\" " +
                    // " where R6.\"BPLId\" in (SELECT T1.\"BPLId\" FROM \"OUSR\" T0 INNER JOIN \"USR6\" T1 ON T1.\"UserCode\" = T0.\"USER_CODE\" " +
                    // " WHERE T0.\"USERID\" in (" + FinalBPLId + ")) ) \"Branch\", " +
                    // " T0.\"CardName\" \"Customer Name\",T7.\"Cellular\" \"Cell Phone\", T7.\"Phone1\" \"Res Phone\",T8.\"GroupName\" \"Occupation\",T0.\"Name\" \"Enquiry vehicle\",  T0.\"UserSign\", " +
                    // " T6.\"SlpName\" \"Entered By\", " +
                    // " case when T0.\"Status\"='O' then 'Open' when T0.\"Status\"='C' then 'Close' when T0.\"Status\"='W' then 'Won' when T0.\"Status\"='L' then 'Lose' End \"Enquiry Status\", " +
                    // " T3.\"Descript\" \"Level of Interest\" ,T4.\"Descript\" \"Enquiry Type\", T2.\"PrjName\" \"Scheme\", " +
                    // " T5.\"Descript\" \"Last Stage\" from \"OOPR\" T0 " +
                    // " INNER JOIN OPR1 T1 ON T0.\"OpprId\" = T1.\"OpprId\" " +
                    // " INNER JOIN OPRJ T2 ON T0.\"PrjCode\" = T2.\"PrjCode\" " +
                    // " INNER JOIN OOIR T3 ON T0.\"IntRate\" = T3.\"Num\" " +
                    // " INNER JOIN OOSR T4 ON T0.\"Source\" = T4.\"Num\" " +
                    // " INNER JOIN OOST T5 ON T1.\"Step_Id\" = T5.\"Num\"  and T1.\"Line\"= (select  max(L.\"Line\") from \"OPR1\" L where L.\"OpprId\"=T0.\"OpprId\") " +
                    // " LEFT JOIN  OSLP T6 ON T6.\"SlpCode\" =T0.\"SlpCode\" " +
                    // " INNER JOIN OCRD T7 ON T7.\"CardCode\" =T0.\"CardCode\" " +
                    // " INNER JOIN OCRG T8 ON T8.\"GroupCode\" =T0.\"CardGroup\" " +
                    // " INNER JOIN \"OUSR\" T9 ON T0.\"UserSign\" = T9.\"USERID\" "+
                    // " INNER JOIN \"USR6\" SR on  SR.\"UserCode\" = T0.\"USER_CODE\" " +
                    // " where SR.\"BPLId\" in (" + FinalBPLId + ") AND T0.\"OpenDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "'  AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' order by T0.\"OpenDate\" ";

                    strSQL1 = "select distinct  T0.\"OpenDate\" \"Enquiry Date\", " +
                     " PL.\"BPLName\" \"Branch\", " +
                     " T0.\"CardName\" \"Customer Name\",T7.\"Cellular\" \"Cell Phone\", T7.\"Phone1\" \"Res Phone\",T8.\"GroupName\" \"Occupation\", " +
                     " TM1.\"U_Model\" \"Model\",TM1.\"U_Variant\" \"Variant\"," +
                     " T0.\"UserSign\", " +
                     " T6.\"SlpName\" \"Entered By\", " +
                     " case when T0.\"Status\"='O' then 'Open' when T0.\"Status\"='C' then 'Close' when T0.\"Status\"='W' then 'Won' when T0.\"Status\"='L' then 'Lose' End \"Enquiry Status\", " +
                     " T3.\"Descript\" \"Level of Interest\" ,T4.\"Descript\" \"Enquiry Type\", T2.\"PrjName\" \"Scheme\", " +
                     " T5.\"Descript\" \"Last Stage\" from \"OOPR\" T0 " +
                     " INNER JOIN OPR1 T1 ON T0.\"OpprId\" = T1.\"OpprId\" " +
                     " LEFT JOIN OPRJ T2 ON T0.\"PrjCode\" = T2.\"PrjCode\" " +
                     " LEFT OUTER JOIN OOIR T3 ON T0.\"IntRate\" = T3.\"Num\" " +
                     " LEFT OUTER JOIN OOSR T4 ON T0.\"Source\" = T4.\"Num\" " +
                     " INNER JOIN OOST T5 ON T1.\"Step_Id\" = T5.\"Num\"  and T1.\"Line\"= (select  max(L.\"Line\") from \"OPR1\" L where L.\"OpprId\"=T0.\"OpprId\") " +
                     " LEFT JOIN  OSLP T6 ON T6.\"SlpCode\" =T0.\"SlpCode\" " +
                     " INNER JOIN OCRD T7 ON T7.\"CardCode\" =T0.\"CardCode\" " +
                     " INNER JOIN OCRG T8 ON T8.\"GroupCode\" =T0.\"CardGroup\" " +
                     " LEFT OUTER JOIN \"OUSR\" T9 ON T0.\"UserSign\" = T9.\"USERID\" " +
                     " LEFT OUTER JOIN \"USR6\" SR on  SR.\"UserCode\" = T9.\"USER_CODE\" " +
                     " LEFT OUTER JOIN \"OBPL\" PL ON PL.\"BPLId\"= SR.\"BPLId\" " +
                     " INNER JOIN \"OITM\" TM1 ON TM1.\"U_Model\"= T0.\"U_Model\" and TM1.\"U_Variant\"= T0.\"U_Variant\" " +
                     " where SR.\"BPLId\" in (" +FinalBPLId+ ") AND T0.\"OpenDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "'  AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "'" + result + " order by T0.\"OpenDate\" ";

                }
                else
                {
                    strSQL1 = "select \"ItmsGrpNam\" \"Model\" ,\"Total Enquiry\",\"Total Booking\" \"Total Quotation\" " +
                              " from(select count(\"Total Enquiry\") \"Total Enquiry\", 0 as \"Total Booking\", \"ItmsGrpNam\" from  (select  distinct T0.\"OpprId\" \"Total Enquiry\", T0.\"U_Model\" \"ItmsGrpNam\" from \"OOPR\" T0  " +
                              " LEFT OUTER JOIN \"OITM\" TM1 on TM1.\"U_Model\" = T0.\"U_Model\" and TM1.\"U_Variant\" = T0.\"U_Variant\"   " +
                              " INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\"  " +
                              " LEFT JOIN OSLP T6 ON T6.\"SlpCode\" = T0.\"SlpCode\"" +
                              "LEFT OUTER JOIN OPRJ T2 ON T0.\"PrjCode\" = T2.\"PrjCode\""+
                              " INNER JOIN \"OUSR\" OU ON T0.\"UserSign\" = OU.\"USERID\" " +
                              " INNER JOIN \"USR6\" SR on  SR.\"UserCode\" = OU.\"USER_CODE\" " +
                              " where  SR.\"BPLId\" in (" +FinalBPLId+ ")  and T0.\"OpenDate\" >= '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' and T0.\"OpenDate\" <= '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result1 + " )group by \"ItmsGrpNam\"" +
                               "UNION ALL"+
                              " select 0 as \"Total Enquiry\",sum(T.\"Quantity\") \"Total Booking\", TB.\"ItmsGrpNam\" from \"QUT1\" T " +
                              " INNER JOIN  \"OQUT\" DR ON DR.\"DocEntry\" = T.\"DocEntry\" " +
                              " LEFT JOIN OPRJ T2 ON T.\"Project\" = T2.\"PrjCode\" " +
                              " LEFT OUTER JOIN \"OITM\" TM1 on T.\"ItemCode\" = TM1.\"ItemCode\" " +
                              " INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\" " +
                              " LEFT JOIN OSLP T6 ON T6.\"SlpCode\" = DR.\"SlpCode\"" +
                              " where DR.\"BPLId\" in (" +FinalBPLId+ ") and DR.\"DocDate\" >='" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' " +
                              " and DR.\"DocDate\" <='" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + " group by TB.\"ItmsGrpNam\")";
                }
                _GridDataTable1.ExecuteQuery(strSQL1);
                _Grid1.DataTable = _GridDataTable1;
                result.Clear();
                result1.Clear();
                _RowHeaders1 = _Grid1.RowHeaders;
                _RowHeaders1.TitleObject.Caption = "#";

                for (int i = 0; i < _Grid1.Rows.Count; i++)
                {
                    _RowHeaders1.SetText(i, (i + 1).ToString());

                }
                //  _Grid1.CollapseLevel = 1;

               
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
                    //_Grid1.Columns.Item(15).Editable = false;
                }
                else
                {
                    _Grid1.Columns.Item(0).Editable = false;
                    _Grid1.Columns.Item(1).Editable = false;
                    _Grid1.Columns.Item(2).Editable = false;
                    ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Enquiry")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Quotation")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                    _Form.Update();
                }
            }
        }

        #endregion

    }
}
