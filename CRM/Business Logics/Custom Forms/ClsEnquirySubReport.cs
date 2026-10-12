using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRM
{
    class ClsEnquirySubReport :Base
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

                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:

                        break;
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

                strSQL = "select distinct T0.\"OpenDate\" \"Enquiry Date\", " +
                      " PL.\"BPLName\" \"Branch\", " +
                      " T0.\"CardName\" \"Customer Name\",T7.\"Cellular\" \"Cell Phone\", T7.\"Phone1\" \"Res Phone\",T8.\"GroupName\" \"Occupation\"," +
                      "T0.\"Name\" \"Enquiry vehicle\", " +
                      " TM1.\"U_Model\" \"Model\",TM1.\"U_Variant\" \"Variant\"," +
                      "T0.\"UserSign\", " +
                      " T6.\"SlpName\" \"Entered By\", " +
                      " case when T0.\"Status\"='O' then 'Open' when T0.\"Status\"='C' then 'Close' when T0.\"Status\"='W' then 'Won' when T0.\"Status\"='L' then 'Lose' End \"Enquiry Status\", " +
                      " T3.\"Descript\" \"Level of Interest\" ,T4.\"Descript\" \"Enquiry Type\", T2.\"PrjName\" \"Scheme\", " +
                      " T5.\"Descript\" \"Last Stage\" from \"OOPR\" T0 " +
                      " INNER JOIN OPR1 T1 ON T0.\"OpprId\" = T1.\"OpprId\" " +
                      " LEFT OUTER JOIN OPRJ T2 ON T0.\"PrjCode\" = T2.\"PrjCode\" " +
                      " LEFT OUTER JOIN OOIR T3 ON T0.\"IntRate\" = T3.\"Num\" " +
                      " LEFT OUTER JOIN OOSR T4 ON T0.\"Source\" = T4.\"Num\" " +
                      " INNER JOIN  OOST T5 ON T1.\"Step_Id\" = T5.\"Num\"  and T1.\"Line\"= (select  max(L.\"Line\") from \"OPR1\" L where L.\"OpprId\"=T0.\"OpprId\") " +
                      " LEFT JOIN  OSLP T6 ON T6.\"SlpCode\" =T0.\"SlpCode\" " +
                      " INNER JOIN OCRD T7 ON T7.\"CardCode\" =T0.\"CardCode\" " +
                      " INNER JOIN OCRG T8 ON T8.\"GroupCode\" =T0.\"CardGroup\" " +
                      " INNER JOIN \"OUSR\" T9 ON T0.\"UserSign\" = T9.\"USERID\" " +
                     " INNER JOIN \"USR6\" SR on  SR.\"UserCode\" = T9.\"USER_CODE\" " +
                     " INNER JOIN \"OBPL\" PL ON PL.\"BPLId\"= SR.\"BPLId\" " +
                     " INNER JOIN \"OITM\" TM1 ON TM1.\"U_Model\"= T0.\"U_Model\" and TM1.\"U_Variant\"= T0.\"U_Variant\" " +
                     " where SR.\"BPLId\" in (" + _FinalBPLID + ") AND T0.\"OpenDate\" between '" + _FromDate + "' AND '" + _ToDate + "' AND TM1.\"U_Model\" =  '" + _Model + "'  " + result + " order by T0.\"OpenDate\" ";//  

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
