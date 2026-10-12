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
    class ClsDeliverySummaryReport : Base
    {

        #region Variable Declaration

        private SAPbouiCOM.Grid _Grid, _Grid1;
        private SAPbouiCOM.DataTable _GridDataTable, _GridDataTable1;

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
        int j = 0;
        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsDeliverySummaryReport()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsDeliverySummaryReport()
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

        #region FormDefault

        public void FormDefault()
        {
            _Form.Freeze(true);

            ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value = DateTime.Now.Year.ToString();
            _GridDataTable = null;
            _Form.DataSources.DataTables.Add("dtReviewList");
            _Grid = (SAPbouiCOM.Grid)Form.Items.Item("4").Specific;
            _GridDataTable = _Form.DataSources.DataTables.Item("dtReviewList");

            _GridDataTable1 = null;
            _Form.DataSources.DataTables.Add("dtReviewList1");
            _Grid1 = (SAPbouiCOM.Grid)Form.Items.Item("5").Specific;
            _GridDataTable1 = _Form.DataSources.DataTables.Item("dtReviewList1");


            string strSQL = string.Empty, strSQL1 = string.Empty;


            strSQL = "select 'Y' as \"Active\",\"BPLId\",(select \"CompnyName\" from OADM) as \"Company Name\",\"BPLName\" as \"Branch Name\" from \"HKTMLIVE\".OBPL where \"Disabled\" = 'N' ";

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
            strSQL1 = "select   \"Item Group\",\"Vehicle\",  \"JANUARY\" , \"FEBRUARY\" , \"MARCH\" , \"APRIL\" , \"MAY\" , \"JUNE\" ,\"JULY\" ,\"AUGUST\" , \"SEPTEMBER\" , \"OCTOBER\" , \"NOVEMBER\" , \"DECEMBER\", " +
                               " ifnull(\"JANUARY\",0) + ifnull(\"FEBRUARY\",0) + ifnull(\"MARCH\",0) + ifnull(\"APRIL\",0) + ifnull(\"MAY\",0) + ifnull(\"JUNE\",0) + ifnull(\"JULY\",0) + ifnull(\"AUGUST\",0) + ifnull(\"SEPTEMBER\",0) + ifnull(\"OCTOBER\",0) + ifnull(\"NOVEMBER\",0) + ifnull(\"DECEMBER\",0) \"Total\" " +
                               //" Round((ifnull(\"JANUARY\",0) + ifnull(\"FEBRUARY\",0) + ifnull(\"MARCH\",0) + ifnull(\"APRIL\",0) + ifnull(\"MAY\",0) + ifnull(\"JUNE\",0) + ifnull(\"JULY\",0) + ifnull(\"AUGUST\",0) + ifnull(\"SEPTEMBER\",0) + ifnull(\"OCTOBER\",0) + ifnull(\"NOVEMBER\",0) + ifnull(\"DECEMBER\",0)) /12,3 ) \"Average Sale\" " +
                               " from (Select " +
                               " T0.\"Dscription\" \"Vehicle\",T5.\"ItmsGrpNam\" \"Item Group\"," +
                              " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='1' then T0.\"Quantity\" else 0 end) As \"JANUARY\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\"))='2'  then T0.\"Quantity\" else 0 end) As \"FEBRUARY\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='3' then T0.\"Quantity\" else 0 end) As \"MARCH\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='4' then T0.\"Quantity\" else 0 end) As \"APRIL\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='5' then T0.\"Quantity\" else 0 end) As \"MAY\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='6' then T0.\"Quantity\" else 0 end) As \"JUNE\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='7' then T0.\"Quantity\" else 0 end) As \"JULY\", " +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='8' then T0.\"Quantity\" else 0 end) As \"AUGUST\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='9' then T0.\"Quantity\" else 0 end) As \"SEPTEMBER\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='10' then T0.\"Quantity\" else 0 end) As \"OCTOBER\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) = '11' then T0.\"Quantity\" else 0 end) As \"NOVEMBER\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='12' then T0.\"Quantity\" else 0 end) As \"DECEMBER\" " +
                               " from \"DLN1\" T0 " +
                               " INNER JOIN \"ODLN\" T3 on T3.\"DocEntry\"= T0.\"DocEntry\" " +
                               " INNER JOIN \"OITM\" T4 on T0.\"ItemCode\" =T4.\"ItemCode\" " +
                               " INNER JOIN \"OITB\" T5 on T4.\"ItmsGrpCod\" = T5.\"ItmsGrpCod\" " +
                               " where Year(To_Dats(T0.\"DocDate\"))= Year(To_Dats('" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value + "')) and T3.\"BPLId\" in ('1','2','3','4') " +
                               " group by T5.\"ItmsGrpNam\",T0.\"Dscription\"  " +
                               " order by T5.\"ItmsGrpNam\")";

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
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JANUARY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("FEBRUARY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("MARCH")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("APRIL")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("MAY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JUNE")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JULY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("AUGUST")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("SEPTEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("OCTOBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("NOVEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("DECEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Average Sale")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

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
              strSQL1=  "select   \"Item Group\",\"Vehicle\",  \"JANUARY\" , \"FEBRUARY\" , \"MARCH\" , \"APRIL\" , \"MAY\" , \"JUNE\" ,\"JULY\" ,\"AUGUST\" , \"SEPTEMBER\" , \"OCTOBER\" , \"NOVEMBER\" , \"DECEMBER\", " +
                               " ifnull(\"JANUARY\",0) + ifnull(\"FEBRUARY\",0) + ifnull(\"MARCH\",0) + ifnull(\"APRIL\",0) + ifnull(\"MAY\",0) + ifnull(\"JUNE\",0) + ifnull(\"JULY\",0) + ifnull(\"AUGUST\",0) + ifnull(\"SEPTEMBER\",0) + ifnull(\"OCTOBER\",0) + ifnull(\"NOVEMBER\",0) + ifnull(\"DECEMBER\",0) \"Total\" " +
                               //" Round((ifnull(\"JANUARY\",0) + ifnull(\"FEBRUARY\",0) + ifnull(\"MARCH\",0) + ifnull(\"APRIL\",0) + ifnull(\"MAY\",0) + ifnull(\"JUNE\",0) + ifnull(\"JULY\",0) + ifnull(\"AUGUST\",0) + ifnull(\"SEPTEMBER\",0) + ifnull(\"OCTOBER\",0) + ifnull(\"NOVEMBER\",0) + ifnull(\"DECEMBER\",0)) /12,3 ) \"Average Sale\" " +
                               " from (Select " +
                               " T0.\"Dscription\" \"Vehicle\",T5.\"ItmsGrpNam\" \"Item Group\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='1' then T0.\"Quantity\" else 0 end) As \"JANUARY\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\"))='2'  then T0.\"Quantity\" else 0 end) As \"FEBRUARY\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='3' then T0.\"Quantity\" else 0 end) As \"MARCH\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='4' then T0.\"Quantity\" else 0 end) As \"APRIL\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='5' then T0.\"Quantity\" else 0 end) As \"MAY\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='6' then T0.\"Quantity\" else 0 end) As \"JUNE\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='7' then T0.\"Quantity\" else 0 end) As \"JULY\", " +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='8' then T0.\"Quantity\" else 0 end) As \"AUGUST\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='9' then T0.\"Quantity\" else 0 end) As \"SEPTEMBER\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='10' then T0.\"Quantity\" else 0 end) As \"OCTOBER\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) = '11' then T0.\"Quantity\" else 0 end) As \"NOVEMBER\"," +
                               " SUM(case when  Month(To_Dats(T3.\"DocDate\")) ='12' then T0.\"Quantity\" else 0 end) As \"DECEMBER\" " +
                               " from \"DLN1\" T0 " +
                               " INNER JOIN \"ODLN\" T3 on T3.\"DocEntry\"= T0.\"DocEntry\" " +
                               " INNER JOIN \"OITM\" T4 on T0.\"ItemCode\" =T4.\"ItemCode\" " +
                               " INNER JOIN \"OITB\" T5 on T4.\"ItmsGrpCod\" = T5.\"ItmsGrpCod\" " +
                               " where Year(To_Dats(T0.\"DocDate\"))= Year(To_Dats('" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value + "')) and T3.\"BPLId\" in (" + FinalBPLId + ") " +
                               " group by T5.\"ItmsGrpNam\",T0.\"Dscription\"  " +
                               " order by T5.\"ItmsGrpNam\")";

                _GridDataTable1.ExecuteQuery(strSQL1);
                _Grid1.DataTable = _GridDataTable1;

                _RowHeaders1 = _Grid1.RowHeaders;
                _RowHeaders1.TitleObject.Caption = "#";

                for (int i = 0; i < _Grid1.Rows.Count; i++)
                {
                    _RowHeaders1.SetText(i, (i + 1).ToString());

                }
                //_Grid1.CollapseLevel = 1;

                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JANUARY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("FEBRUARY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("MARCH")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("APRIL")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("MAY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JUNE")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("JULY")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("AUGUST")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("SEPTEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("OCTOBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("NOVEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("DECEMBER")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                //((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Average Sale")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
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

            }
        }

    }
}
