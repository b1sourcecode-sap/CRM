using SAPbouiCOM;
using System;
using System.Text;
using System.Windows.Forms;

namespace CRM
{
    class ClsDeliveryReport : Base
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
        public string BPLId = string.Empty, FinalBPLId = string.Empty;
        int j = 0;
        StringBuilder result = new StringBuilder();
        StringBuilder result1 = new StringBuilder();
        public static int count = 1;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsDeliveryReport()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsDeliveryReport()
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
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("25").Specific).Value != "")
                        {
                            result.Append(" And T7.\"SlpName\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("25").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("27").Specific).Value != "")
                        {
                            result.Append(" And T6.\"PrjCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("27").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("23").Specific).Value != "")
                        {
                            result.Append(" And TM1.\"U_Model\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("23").Specific).Value + "'");
                        }

                        if (((SAPbouiCOM.EditText)_Form.Items.Item("26").Specific).Value != "")
                        {
                            result.Append(" AND TM1.\"U_Variant\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("26").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value != "")
                        {
                            result.Append(" AND TM1.\"U_Colour\"  = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value + "'");
                        }

                        if (((SAPbouiCOM.EditText)_Form.Items.Item("25").Specific).Value != "")
                        {
                            result1.Append(" And T7.\"SlpName\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("25").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("27").Specific).Value != "")
                        {
                            result1.Append(" And T6.\"PrjCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("27").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("23").Specific).Value != "")
                        {
                            result1.Append(" And T.\"U_Model\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("23").Specific).Value + "'");
                        }

                        if (((SAPbouiCOM.EditText)_Form.Items.Item("26").Specific).Value != "")
                        {
                            result1.Append(" AND T.\"U_Variant\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("26").Specific).Value + "'");
                        }
                        if (((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value != "")
                        {
                            result1.Append(" AND TM1.\"U_Colour\"  = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value + "'");
                        }

                        break;

                    #endregion

                    #region key down event
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "23" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("23").Specific).Value == "")
                            {
                                string Query = "select distinct \"U_Model\" as \"Model\" from OITM where \"U_Model\" is not null";
                                this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Model";
                                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, Utilities.Hana_ADOConnection(Query), false);

                            }
                        }

                        if (pVal.ItemUID == "26" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("26").Specific).Value == "" && ((SAPbouiCOM.EditText)_Form.Items.Item("23").Specific).Value != "")
                            {
                                //string Query = "select distinct  SUBSTRING(T1.\"ItemName\", (Locate(T1.\"ItemName\", '-', 1)) + 1, (Locate(T1.\"ItemName\", '-', -2)) - 1) \"Varient\"" +
                                //                "from \"OITM\" T1 where \"U_Branch\" = 'Vehicle' and SUBSTRING(T1.\"ItemName\", (Locate(T1.\"ItemName\", '-', 1)) + 1, (Locate(T1.\"ItemName\", '-', -2)) - 1)<> '' ";
                                string Query = "select distinct \"U_Variant\" from OITM where \"U_Model\" = '" + (((SAPbouiCOM.EditText)_Form.Items.Item("23").Specific).Value) + "'";
                                this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Varient";
                                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, Utilities.Hana_ADOConnection(Query), false);

                            }
                        }
                        if (pVal.ItemUID == "24" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value == "" && ((SAPbouiCOM.EditText)_Form.Items.Item("26").Specific).Value != "" && ((SAPbouiCOM.EditText)_Form.Items.Item("23").Specific).Value != "")
                            {
                                //string Query = "select distinct SUBSTRING(T1.\"ItemName\", (Locate(T1.\"ItemName\", '-', -2)) + 1)  \"Color\"" +
                                //                "from \"OITM\" T1 where \"U_Branch\" = 'Vehicle'; ";
                                string Query = "select distinct \"U_Colour\" from OITM where \"U_Model\" = '" + (((SAPbouiCOM.EditText)_Form.Items.Item("23").Specific).Value) + "' AND \"U_Variant\" = '" + (((SAPbouiCOM.EditText)_Form.Items.Item("26").Specific).Value) + "' ";

                                this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Color";
                                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, Utilities.Hana_ADOConnection(Query), false);

                            }
                        }
                        if (pVal.ItemUID == "27" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("27").Specific).Value == "")
                            {
                                string Query = "select \"PrjCode\" as \"Scheme Code\",\"PrjName\" as \"Scheme Name\" from OPRJ";
                                this._Object = new ClsChooseFromList();
                                ((ClsChooseFromList)this._Object).FormTitle = "Scheme";
                                ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref BubbleEvent, Utilities.Hana_ADOConnection(Query), false);

                            }
                        }

                        if (pVal.ItemUID == "25" && pVal.CharPressed == 9)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("25").Specific).Value == "")
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

                            this._Object = new ClsDeliverySubReport();
                            this.OpenChildForm(Constants.Forms.DeliverySubReport, ref pVal, true);
                            ((ClsDeliverySubReport)_Object).FormDefault(_FromDate, _ToDate, _Model, FinalBPLId, result);
                          
                            Utilities.Message("Operation completed sucessfully.", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }

                        break;
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

                if (this._CFLEvent.ItemUID == "23")
                {
                    _Form.Freeze(true);
                    ((SAPbouiCOM.EditText)_Form.Items.Item("23").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                    _Form.Freeze(false);
                }
                if (this._CFLEvent.ItemUID == "26")
                {
                    _Form.Freeze(true);
                    ((SAPbouiCOM.EditText)_Form.Items.Item("26").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                    _Form.Freeze(false);
                }
                if (this._CFLEvent.ItemUID == "24")
                {
                    _Form.Freeze(true);
                    ((SAPbouiCOM.EditText)_Form.Items.Item("24").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                    _Form.Freeze(false);
                }
                if (this._CFLEvent.ItemUID == "27")
                {
                    _Form.Freeze(true);
                    ((SAPbouiCOM.EditText)_Form.Items.Item("27").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
                    _Form.Freeze(false);
                }
                if (this._CFLEvent.ItemUID == "25")
                {
                    _Form.Freeze(true);
                    ((SAPbouiCOM.EditText)_Form.Items.Item("25").Specific).Value = _CFLEvent.oCFL_SelectedResult.Rows[0][0].ToString();
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
            strSQL1 = "select distinct OP.\"OpenDate\" \"Enquiry Date\", T3.\"DocDate\" \"Booked Date\", T0.\"DocDate\" \"Sold Date\", T0.\"CardName\" \"Name\", RD.\"Address\", " +
                        " T4.\"Cellular\" \"Cell Phone\",T5.\"GroupName\" \"Occupation\"," +
                       //" SUBSTRING(T1.\"Dscription\", 1, (Locate(T1.\"Dscription\", '-', 1)) - 1) \"Model\"," +
                       //" SUBSTRING(T1.\"Dscription\", (Locate(T1.\"Dscription\", '-', 1)) + 1, (Locate(T1.\"Dscription\", '-', -2)) - 1) \"Varient\"," +
                       //" SUBSTRING(TM1.\"ItemName\", (Locate(TM1.\"ItemName\", '-', -2)) + 1) \"Color\"," +
                       " TM1.\"U_Model\" \"Model\",TM1.\"U_Variant\" \"Variant\",TM1.\"U_Colour\" \"Color\"," +
                          "T1.\"Dscription\" \"Vehicle\",  T6.\"PrjName\" \"Scheme\", " +
                        " T14.\"DistNumber\" \"Engine\",T14.\"MnfSerial\" \"VIN\",T14.\"U_RegNo\" \"Reg No\", " +
                        "ITM.\"Price\" \"MRP\", T0.\"DiscSum\" \"Discount Amount\",ITM.\"Price\" -T0.\"DiscSum\" \"Deal Value\"," +
                        " T7.\"SlpName\" \"Sales Employee\" from \"ODLN\" T0 " +
                        " INNER JOIN \"DLN1\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                        " INNER JOIN \"RDR1\" T2 ON T2.\"DocEntry\" = T1.\"BaseEntry\" and T2.\"LineNum\" = T1.\"BaseLine\" " +
                        " INNER JOIN \"ORDR\" T3 ON T3.\"DocEntry\" = T2.\"DocEntry\" " +
                        " INNER JOIN \"OITM\" TM1 ON TM1.\"ItemCode\" = T1.\"ItemCode\"" +
                        " LEFT JOIN \"ITM1\" ITM on ITM.\"ItemCode\" = T1.\"ItemCode\" and ITM.\"PriceList\" = '2'" +
                        " INNER JOIN \"OITB\" TB ON TM1.\"ItmsGrpCod\" = TB.\"ItmsGrpCod\"" +
                        " INNER JOIN \"OCRD\" T4 ON T4.\"CardCode\" = T0.\"CardCode\" " +
                        " LEFT OUTER JOIN \"CRD1\" RD ON RD.\"CardCode\" = T4.\"CardCode\" " +
                        " INNER JOIN \"OCRG\" T5 ON T5.\"GroupCode\" = T4.\"GroupCode\" " +
                        " LEFT OUTER JOIN \"OPRJ\" T6 ON T0.\"Project\" = T6.\"PrjCode\" " +
                        " LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = T0.\"SlpCode\" " +
                        " LEFT OUTER JOIN \"OPR1\" PR ON PR.\"DocId\" = T1.\"DocEntry\" and PR.\"ObjType\" = T1.\"ObjType\" " +
                        " LEFT JOIN \"OOPR\" OP on PR.\"OpprId\" = OP.\"OpprId\"" +
                        " INNER JOIN SRI1 I1 on T1.\"ItemCode\" = I1.\"ItemCode\" and(T1.\"DocEntry\" = I1.\"BaseEntry\" and T1.\"ObjType\" = I1.\"BaseType\")  " +
                        " INNER JOIN OSRN T14 on T14.\"ItemCode\" = I1.\"ItemCode\" and I1.\"SysSerial\" = T14.\"SysNumber\" "+
                        " where T0.\"DocDate\" between (SELECT ADD_MONTHS(NEXT_DAY(LAST_DAY(CURRENT_DATE)),-1) FROM DUMMY) AND (SELECT LAST_DAY(CURRENT_DATE) FROM DUMMY) order by T0.\"DocDate\" ";

            _GridDataTable1.ExecuteQuery(strSQL1);
            _Grid1.DataTable = _GridDataTable1;

            _RowHeaders1 = _Grid1.RowHeaders;
            _RowHeaders1.TitleObject.Caption = "#";
           

            for (int i = 0; i < _Grid1.Rows.Count; i++)
            {
                _RowHeaders1.SetText(i, (i + 1).ToString());

            }

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
            //_Grid1.Columns.Item(19).Editable = false;
            //_Grid1.Columns.Item(20).Editable = false;
            
            /******************/

            _Form.Freeze(false);

        }
        #endregion

        #region FilterGridValue
        public void FilterGridValue()
        {
            this._GridDataTable1.Rows.Clear();
            if (FinalBPLId != "")
            {
                string strSQL = string.Empty, strSQL1 = string.Empty;
                _OptionBtn = (SAPbouiCOM.OptionBtn)_Form.Items.Item("13").Specific;
                if (_OptionBtn.Selected == true)
                {
                    strSQL1 = "select distinct OP.\"OpenDate\" \"Enquiry Date\", T3.\"DocDate\" \"Booked Date\", T0.\"DocDate\" \"Sold Date\", T0.\"CardName\" \"Name\", RD.\"Address\", " +
                       " T4.\"Cellular\" \"Cell Phone\",T5.\"GroupName\" \"Occupation\"," +
                       //" SUBSTRING(T1.\"Dscription\", 1, (Locate(T1.\"Dscription\", '-', 1)) - 1) \"Model\"," +
                       //" SUBSTRING(T1.\"Dscription\", (Locate(T1.\"Dscription\", '-', 1)) + 1, (Locate(T1.\"Dscription\", '-', -2)) - 1) \"Varient\"," +
                       //" SUBSTRING(TM1.\"ItemName\", (Locate(TM1.\"ItemName\", '-', -2)) + 1) \"Color\"," +
                       " TM1.\"U_Model\" \"Model\",TM1.\"U_Variant\" \"Variant\",TM1.\"U_Colour\" \"Color\","+
                       " T1.\"Dscription\" \"Vehicle\",  T6.\"PrjName\" \"Scheme\", " +
                       " T14.\"DistNumber\" \"Engine\",T14.\"MnfSerial\" \"VIN\",T14.\"U_RegNo\" \"Reg No\", " +
                       " ITM.\"Price\" \"MRP\", T0.\"DiscSum\" \"Discount Amount\",ITM.\"Price\" -T0.\"DiscSum\" \"Deal Value\", " +
                       " T7.\"SlpName\" \"Sales Employee\" from \"ODLN\" T0 " +
                       " INNER JOIN \"DLN1\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                       " INNER JOIN \"RDR1\" T2 ON T2.\"DocEntry\" = T1.\"BaseEntry\" and T2.\"LineNum\" = T1.\"BaseLine\" " +
                       " INNER JOIN \"ORDR\" T3 ON T3.\"DocEntry\" = T2.\"DocEntry\" " +
                       " INNER JOIN \"OITM\" TM1 ON TM1.\"ItemCode\" = T1.\"ItemCode\""+
                       " LEFT JOIN \"ITM1\" ITM on ITM.\"ItemCode\" = T1.\"ItemCode\" and ITM.\"PriceList\" = '2'"+
                       " INNER JOIN \"OITB\" TB ON TM1.\"ItmsGrpCod\" = TB.\"ItmsGrpCod\"" +
                       " INNER JOIN \"OCRD\" T4 ON T4.\"CardCode\" = T0.\"CardCode\" " +
                       " LEFT OUTER JOIN \"CRD1\" RD ON RD.\"CardCode\" = T4.\"CardCode\" " +
                       " LEFT OUTER JOIN \"OCRG\" T5 ON T5.\"GroupCode\" = T4.\"GroupCode\" " +
                       " LEFT OUTER JOIN \"OPRJ\" T6 ON T0.\"Project\" = T6.\"PrjCode\" " +
                       " LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = T0.\"SlpCode\" " +
                       " LEFT OUTER JOIN \"OPR1\" PR ON PR.\"DocId\" = T1.\"DocEntry\" and PR.\"ObjType\" = T1.\"ObjType\" " +
                       " LEFT JOIN \"OOPR\" OP on PR.\"OpprId\" = OP.\"OpprId\"" +
                       " INNER JOIN SRI1 I1 on T1.\"ItemCode\" = I1.\"ItemCode\" and(T1.\"DocEntry\" = I1.\"BaseEntry\" and T1.\"ObjType\" = I1.\"BaseType\")  " +
                       " INNER JOIN OSRN T14 on T14.\"ItemCode\" = I1.\"ItemCode\" and I1.\"SysSerial\" = T14.\"SysNumber\" "+
                       " where T0.\"BPLId\" in (" + FinalBPLId + ") AND T0.\"DocDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + " order by T0.\"DocDate\" ";// --and TM."PriceList" = '2'";

                }
                else
                {
                    strSQL1 = "select  \"ItmsGrpNam\" \"Model\" ,\"Total Enquiry\",  \"Total Quotation\",  \"Total Booking\", \"Total Delivery\" ," +
                                "\"Total Invoice\" from ( " +
                                "(select count(\"Total Enquiry\") \"Total Enquiry\", 0 as\"Total Quotation\",  0 as \"Total Booking\", 0 as \"Total Allocation\" ,0 as \"Total Delivery\", 0 as \"Total Invoice\",\"ItmsGrpNam\" from  (select  distinct T0.\"OpprId\" \"Total Enquiry\", T0.\"U_Model\" \"ItmsGrpNam\" from \"OOPR\" T0  " +
                              " LEFT OUTER JOIN \"OITM\" TM1 on TM1.\"U_Model\" = T0.\"U_Model\" and TM1.\"U_Variant\" = T0.\"U_Variant\"   " +
                              " INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\"  " +
                              " LEFT JOIN OSLP T6 ON T6.\"SlpCode\" = T0.\"SlpCode\"" +
                              "LEFT OUTER JOIN OPRJ T2 ON T0.\"PrjCode\" = T2.\"PrjCode\"" +
                              " INNER JOIN \"OUSR\" OU ON T0.\"UserSign\" = OU.\"USERID\" " +
                              " INNER JOIN \"USR6\" SR on  SR.\"UserCode\" = OU.\"USER_CODE\" " +
                              " where  SR.\"BPLId\" in (" + FinalBPLId + ")  and T0.\"OpenDate\" >= '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' and T0.\"OpenDate\" <= '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result1 + " )group by \"ItmsGrpNam\")   " +
                                "UNION ALL (select 0 as\"Total Enquiry\", sum(T.\"Quantity\") \"Total Quotation\",  0 as \"Total Booking\", 0 as \"Total Allocation\" ,0 as \"Total Delivery\",0 as \"Total Invoice\", " +
                                "TB.\"ItmsGrpNam\" \"ItmsGrpNam\"  from \"QUT1\" T " +
                                "INNER JOIN  \"OQUT\" DR ON DR.\"DocEntry\" = T.\"DocEntry\" " +
                                "LEFT OUTER JOIN \"OITM\" TM1 on T.\"ItemCode\" = TM1.\"ItemCode\" " +
                                "INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\" " +
                                "LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = T.\"SlpCode\" " +
                                "Left JOIN \"OPRJ\" T6 ON DR.\"Project\" = T6.\"PrjCode\" " +
                                "where DR.\"BPLId\" in (" + FinalBPLId + ") and DR.\"DocDate\" between  '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + " group by TB.\"ItmsGrpNam\") " +

                                "UNION ALL  (select 0 as\"Total Enquiry\", 0 as\"Total Quotation\", sum(T.\"Quantity\") \"Total Booking\", 0 as \"Total Allocation\" ,0 as \"Total Delivery\",0 as \"Total Invoice\",TB.\"ItmsGrpNam\" \"ItmsGrpNam\" from \"RDR1\" T " +
                                "INNER JOIN ORDR DR ON DR.\"DocEntry\" = T.\"DocEntry\" " +
                                "LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = DR.\"SlpCode\" " +
                                "Left JOIN \"OPRJ\" T6 ON T.\"Project\" = T6.\"PrjCode\" " +
                                "LEFT OUTER JOIN \"OITM\" TM1 on T.\"ItemCode\" = TM1.\"ItemCode\" " +
                                "INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\" " +
                                "where DR.\"BPLId\" in (" + FinalBPLId + ") and DR.\"DocDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + " group by TB.\"ItmsGrpNam\") " +

                                "UNION ALL  (select 0 as\"Total Enquiry\",0 as\"Total Quotation\",  0 as \"Total Booking\", sum(L1.\"AllocQty\") \"Total Allocation\" ,0 as \"Total Delivery\",0 as \"Total Invoice\",TB.\"ItmsGrpNam\" \"ItmsGrpNam\"  from \"RDR1\" R1 " +
                                "INNER JOIN \"ORDR\" DR ON DR.\"DocEntry\" = R1.\"DocEntry\"" +
                                "LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = DR.\"SlpCode\" " +
                                "Left JOIN \"OPRJ\" T6 ON R1.\"Project\" = T6.\"PrjCode\" " +
                                "LEFT OUTER JOIN \"OITM\" TM1 on R1.\"ItemCode\" = TM1.\"ItemCode\" " +
                                "INNER JOIN \"OITB\" TB ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\" " +
                                "LEFT JOIN \"OITL\" TL on TL.\"DocEntry\" = R1.\"DocEntry\" and TL.\"DocType\" = R1.\"ObjType\" " +
                                "INNER JOIN \"ITL1\"  L1 On L1.\"LogEntry\" = TL.\"LogEntry\" " +
                                "where DR.\"BPLId\" in (" + FinalBPLId + ") and  R1.\"TargetType\" <> '15' and  DR.\"DocDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + " group by TB.\"ItmsGrpNam\") " +

                                "UNION ALL  (select 0 as\"Total Enquiry\",0 as\"Total Quotation\",  0 as \"Total Booking\",0 as \"Total Allocation\", sum(T.\"Quantity\") \"Total Delivery\",0 as \"Total Invoice\"," +
                                "TB.\"ItmsGrpNam\" \"ItmsGrpNam\" from \"DLN1\" T " +
                                "INNER JOIN ODLN DL ON DL.\"DocEntry\" = T.\"DocEntry\" " +
                                "LEFT JOIN  \"OSLP\" T7 ON T7.\"SlpCode\" = DL.\"SlpCode\" " +
                                "Left JOIN \"OPRJ\" T6 ON T.\"Project\" = T6.\"PrjCode\" " +
                                "LEFT OUTER JOIN \"OITM\" TM1 on T.\"ItemCode\" = TM1.\"ItemCode\" " +
                                "INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\" " +
                                "where DL.\"BPLId\" in (" + FinalBPLId + ") and DL.\"DocDate\"  between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + " group by TB.\"ItmsGrpNam\") " +

                                "UNION ALL  (select  0 as\"Total Enquiry\",0 as\"Total Quotation\",  0 as \"Total Booking\", 0 as \"Total Allocation\",0 as \"Total Delivery\",sum(T.\"Quantity\") \"Total Invoice\", " +
                                "TB.\"ItmsGrpNam\" \"ItmsGrpNam\" from \"INV1\" T " +
                                "INNER JOIN \"OINV\" DL ON DL.\"DocEntry\" = T.\"DocEntry\" " +
                                "LEFT JOIN \"OSLP\" T7 ON T7.\"SlpCode\" = DL.\"SlpCode\" " +
                                "Left JOIN \"OPRJ\" T6 ON T.\"Project\" = T6.\"PrjCode\" " +
                                "LEFT OUTER JOIN \"OITM\" TM1 on T.\"ItemCode\" = TM1.\"ItemCode\" " +
                                "INNER JOIN \"OITB\" TB  ON TB.\"ItmsGrpCod\" = TM1.\"ItmsGrpCod\" " +
                                "where DL.\"BPLId\" in (" + FinalBPLId + ") and DL.\"DocDate\" between '" + ((SAPbouiCOM.EditText)_Form.Items.Item("11").Specific).Value.ToString() + "' AND '" + ((SAPbouiCOM.EditText)_Form.Items.Item("10").Specific).Value.ToString() + "' " + result + " group by TB.\"ItmsGrpNam\") )";


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
                //  _Grid1.CollapseLevel = 1;
              
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
                    //_Grid1.Columns.Item(19).Editable = false;
                    //_Grid1.Columns.Item(20).Editable = false;
                }
                else
                {
                    _Grid1.Columns.Item(0).Editable = false;
                    _Grid1.Columns.Item(1).Editable = false;
                    _Grid1.Columns.Item(2).Editable = false;
                    _Grid1.Columns.Item(3).Editable = false;
                    _Grid1.Columns.Item(4).Editable = false;
                    _Grid1.Columns.Item(5).Editable = false;
                    ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Enquiry")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Quotation")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Booking")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Delivery")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    ((SAPbouiCOM.EditTextColumn)_Grid1.Columns.Item("Total Invoice")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                    _Form.Update();
                }

            }
        }
        #endregion
    }
}
