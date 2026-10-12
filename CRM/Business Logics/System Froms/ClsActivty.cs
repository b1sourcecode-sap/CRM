using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRM
{
    class ClsActivty : Base
    {

        SAPbouiCOM.EditText _EditText;
        public static string StartDate = string.Empty;
        public static string EndDate = string.Empty;
        public static int ActivtyNo = 0;
        public static int _RowNo = 0;
        SAPbouiCOM.Matrix _Matrix2;
        SAPbouiCOM.ComboBox _ComboBox;
        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RS2 = null;
        public static bool _Flag = false;


        #region Form Default
        internal void FormDefault( string cardCode,int RowNo,bool Flag)
        {
            try
            {
                SAPbouiCOM.Form _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                ((SAPbouiCOM.EditText)_Form.Items.Item("9").Specific).Value = cardCode.ToString();
                _RowNo = RowNo;
                _Flag = Flag;
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }

        internal void FormDefaultExit(string cardCode, int RowNo, bool Flag, string _ActivtyNo)
        {
            try
            {
                SAPbouiCOM.Form _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                //((SAPbouiCOM.EditText)_Form.Items.Item("9").Specific).Value = cardCode.ToString();
                _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                ((SAPbouiCOM.EditText)_Form.Items.Item("5").Specific).Value = _ActivtyNo;
                _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _RowNo = RowNo;
                _Flag = Flag;
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
                        if (pVal.ItemUID == "1")//&& _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE
                        {

                        //    StartDate = _Form.DataSources.DBDataSources.Item("OCLG").GetValue("Recontact", 0).Trim();
                        //    EndDate = _Form.DataSources.DBDataSources.Item("OCLG").GetValue("endDate", 0).Trim();
                        //    _Form.Close();
                        //    _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;

                        //    _Matrix2 = (SAPbouiCOM.Matrix)Form.Items.Item("32").Specific;
                        //    _Matrix2.AddRow(1, -1);

                        //    ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_7").Cells.Item(_RowNo).Specific).Value = StartDate.ToString();


                            //_Matrix2.FlushToDataSource();
                            //_ChildDBDataSource2.SetValue("U_CreateAc", 0, ".");
                            //_Matrix2.LoadFromDataSource();


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
                        if (pVal.ItemUID == "1" && pVal.Action_Success == true && _Flag == true) // && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        {
                            
                            StartDate = _Form.DataSources.DBDataSources.Item("OCLG").GetValue("Recontact", 0).Trim();
                            EndDate = _Form.DataSources.DBDataSources.Item("OCLG").GetValue("endDate", 0).Trim();
                            ActivtyNo = Convert.ToInt16(_Form.DataSources.DBDataSources.Item("OCLG").GetValue("ClgCode", 0).Trim()) - 1;//  
                            _Form.Close();

                            string _SQL = "select case when T0.\"AttendUser\" is null then T1.\"firstName\" || ' ' || T1.\"lastName\" else T2.\"U_NAME\" END \"AssignedTo\", " +
                                       " T3.\"U_NAME\" as \"AssignedBy\" ,Case when T0.\"inactive\" = 'N' and \"Closed\" = 'N' then 'Pending' When T0.\"inactive\" = 'N' then 'Closed' else 'InActive' END as \"Status\" from OCLG T0 left outer join OHEM T1 ON T1.\"empID\" = T0.\"AttendEmpl\" " +
                                       " left outer join OUSR T2 ON T2.\"USERID\" = T0.\"AttendUser\" " +
                                       " left outer join OUSR T3 ON T3.\"USERID\" = T0.\"AssignedBy\" where \"ClgCode\" = '" + ActivtyNo + "'";
                            Utilities.ExecuteSQL(ref _RS, _SQL);

                            string AssignedTo = _RS.Fields.Item(0).Value.ToString();
                            string AssignedBy = _RS.Fields.Item(1).Value.ToString();
                            string ActStatus = _RS.Fields.Item(2).Value.ToString();

                            _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;

                            DateTime StrtDate = Utilities.ToDate(StartDate.ToString());
                            DateTime EdDate = Utilities.ToDate(StartDate.ToString());

                            _Matrix2 = (SAPbouiCOM.Matrix)Form.Items.Item("32").Specific;
                            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_14").Cells.Item(_RowNo).Specific).Value = StrtDate.ToString("yyyyMMdd");
                            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_13").Cells.Item(_RowNo).Specific).Value = EdDate.ToString("yyyyMMdd");
                            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_12").Cells.Item(_RowNo).Specific).Value = AssignedTo;
                            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_15").Cells.Item(_RowNo).Specific).Value = AssignedBy;
                            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_11").Cells.Item(_RowNo).Specific).Value = ActivtyNo.ToString();
                            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_5").Cells.Item(_RowNo).Specific).Value = Utilities.Application.Company.UserName;
                            ((SAPbouiCOM.EditText)_Matrix2.Columns.Item("V_16").Cells.Item(_RowNo).Specific).Value = ActStatus;
                            
                            //_Matrix2.FlushToDataSource();
                            //_ChildDBDataSource2.SetValue("U_CreateAc", 0, ".");
                            //_Matrix2.LoadFromDataSource();

                          
                        }
                       
                        break;
                }
            }

            #endregion

        }
        #endregion

    }
}
