using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRM
{
    class ClsBookingSubReport:Base
    {

        #region Variable Declaration

        private SAPbouiCOM.Grid _Grid;
        private SAPbouiCOM.DataTable _GridDataTable;
        private SAPbouiCOM.EditTextColumn _Column;
        private SAPbouiCOM.RowHeaders _RowHeaders;
        public static string _FromDt = string.Empty, _ToDt = string.Empty;
        public static int count = 1;
        #endregion

        #region Item Events
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.Before_Action)
            {
                switch (pVal.EventType)
                {

                    //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:

                    //    break;
                }
            }
            #endregion

            #region After Action
            else
            {
                switch (pVal.EventType)
                {

                    #region et_ITEM_PRESSED

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "100")
                        {
                            _Form.Close();
                        }
                        break;

                        #endregion
                }
            }

            #endregion
        }
        #endregion

        #region default method
        public void FormDefault(string _FromDate, string _ToDate, string _Model, string _FinalBPLID, StringBuilder result)
        {
            try
            {
                _Form.Freeze(true);

                _FromDt = _FromDate;
                _ToDt = _ToDate;

                Utilities.Application.SBO_Application.StatusBar.SetText("Please wait calculating summary.........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                _Form.DataSources.DataTables.Add("dtReviewList");
                _Grid = (SAPbouiCOM.Grid)Form.Items.Item("21").Specific;
                _GridDataTable = _Form.DataSources.DataTables.Item("dtReviewList");
                string strSQL = string.Empty;

                        strSQL = "select   OP.\"OpenDate\" \"Enquiry Date\",T0.\"DocDate\" \"Booked Date\", T0.\"BPLName\" \"Branch\", T0.\"CardName\" \"Customer\", RD.\"Address\", " +
                       " T4.\"Cellular\" \"Cell Phone\",T5.\"GroupName\" \"Occupation\"," +
                        //" SUBSTRING(T1.\"Dscription\", 1, (Locate(T1.\"Dscription\", '-', 1)) - 1) \"Model\"," +
                        //" SUBSTRING(T1.\"Dscription\", (Locate(T1.\"Dscription\", '-', 1)) + 1, (Locate(T1.\"Dscription\", '-', -2)) - 1) \"Varient\"," +
                        //" SUBSTRING(TM1.\"ItemName\", (Locate(TM1.\"ItemName\", '-', -2)) + 1) \"Color\"," +
                        " TM1.\"U_Model\" \"Model\",TM1.\"U_Variant\" \"Variant\",TM1.\"U_Colour\" \"Color\"," +
                       " T1.\"Dscription\" \"Vehicle\",Case When  T0.\"DocStatus\" = 'O' then 'OPEN' " +
                       " When T0.\"DocStatus\" = 'C' then  'Close' end \"Status\", T6.\"PrjName\" \"Scheme\",ST.\"Descript\" \"Last Stage\", " +
                       "ITM.\"Price\" \"MRP\", T0.\"DiscSum\" \"Discount Amount\",ITM.\"Price\" -T0.\"DiscSum\" \"Deal Value\", " +
                       "T0.\"VatSum\" \"Vat Amount\", T0.\"DocTotal\" \"Document Total\"," +
                       "T7.\"SlpName\" \"Entered By\" from \"ORDR\" T0 " +
                       "INNER JOIN RDR1 T1 ON T1.\"DocEntry\" = T0.\"DocEntry\"" +
                       "LEFT OUTER JOIN \"OITM\" TM1 on T1.\"ItemCode\" = TM1.\"ItemCode\"" +
                       "INNER JOIN \"OCRD\" T4 ON T4.\"CardCode\" = T0.\"CardCode\"" +
                       "LEFT OUTER JOIN \"CRD1\" RD ON RD.\"CardCode\" = T4.\"CardCode\"" +
                       "INNER JOIN \"OCRG\" T5 ON T5.\"GroupCode\" = T4.\"GroupCode\"" +
                       "LEFT OUTER JOIN \"OPRJ\" T6 ON T0.\"Project\" = T6.\"PrjCode\"" +
                       "LEFT OUTER JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = T0.\"SlpCode\"" +
                       "LEFT OUTER JOIN \"ITM1\" ITM on ITM.\"ItemCode\" = TM1.\"ItemCode\" and ITM.\"PriceList\" = '2' " +
                       "INNER JOIN \"OITB\" TB on TB.\"ItmsGrpCod\"= TM1.\"ItmsGrpCod\"" +
                       "LEFT OUTER JOIN \"OPR1\" T8 ON T8.\"DocId\" = T1.\"DocEntry\" and T8.\"ObjType\" = T1.\"ObjType\" " +
                       "LEFT OUTER JOIN \"OOPR\" OP on T8.\"OpprId\" = OP.\"OpprId\"" +
                       "Left JOIN OOST ST ON T8.\"Step_Id\" = ST.\"Num\"  and T8.\"Line\" = (select  max(L.\"Line\") from \"OPR1\" L where L.\"OpprId\" = T8.\"OpprId\")" +
                       "where T0.\"BPLId\" in (" + _FinalBPLID + ") AND T0.\"DocDate\" between '" + _FromDate + "' AND '" + _ToDate + "' AND TB.\"ItmsGrpNam\" = '" + _Model + "'  " + result + " order by T0.\"DocDate\" ";//  

                _GridDataTable.ExecuteQuery(strSQL);
                _Grid.DataTable = _GridDataTable;
                _RowHeaders = _Grid.RowHeaders;
                _RowHeaders.TitleObject.Caption = "#";
                _Grid.RowHeaders.Width = 50;

                for (int i = 0; i < _Grid.Rows.Count; i++)
                {
                    _RowHeaders.SetText(i, (i + 1).ToString());
                }
                _Grid.Columns.Item(0).Editable = false;
                _Grid.Columns.Item(1).Editable = false;
                _Grid.Columns.Item(2).Editable = false;
                _Grid.Columns.Item(3).Editable = false;
                _Grid.Columns.Item(4).Editable = false;
                _Grid.Columns.Item(5).Editable = false;
                _Grid.Columns.Item(6).Editable = false;
                _Grid.Columns.Item(7).Editable = false;
                _Grid.Columns.Item(8).Editable = false;
                _Grid.Columns.Item(9).Editable = false;
                _Grid.Columns.Item(10).Editable = false;
                _Grid.Columns.Item(11).Editable = false;
                _Grid.Columns.Item(12).Editable = false;
                _Grid.Columns.Item(13).Editable = false;
                _Grid.Columns.Item(14).Editable = false;
                _Grid.Columns.Item(15).Editable = false;
                _Grid.Columns.Item(16).Editable = false;
                _Grid.Columns.Item(17).Editable = false;
                _Grid.Columns.Item(18).Editable = false;
                _Grid.Columns.Item(19).Editable = false;
                Utilities.Message("Operation completed sucessfully.", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                _Form.Freeze(false);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        #endregion
    }
}
