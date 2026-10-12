using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRM
{
    class ClsSalesOpportunity : Base
    {

        #region Variable declaration

        SAPbobsCOM.Recordset _RS = null, _RS1 = null, _RS2 = null, _RS3 = null, _RS4 = null;
        private SAPbouiCOM.Matrix _Matrix, _Matrix1, _Matrix2, _Matrix3;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource, _ChildDBDataSource1, _ChildDBDataSource2, _ChildDBDataSource3;
        SAPbouiCOM.ComboBox _ComboBox;
        public string _CardCode = string.Empty;
        public string _SODocEntry = string.Empty; 
        private int RowToDelete;
        private static string _BPCode = string.Empty,_ModelName = string.Empty;
        private static int _UserCd = 0;
        private static int _OpprNo = 0;
        private bool OpportunityStatus = false;
        public static SAPbouiCOM.Form _InquiryForm;
        private SAPbouiCOM.EditTextColumn _Column;        
        private SAPbouiCOM.Item _ExistItem;
        private SAPbouiCOM.Item _NewItem;
        private SAPbouiCOM.Button _Button;
        private SAPbouiCOM.EditText _EditText;
        private SAPbouiCOM.Grid _Grid;
        private SAPbouiCOM.DataTable _GridDataTable;
        SAPbouiCOM.Folder _Folder;
        private bool flag = false;
        private SAPbouiCOM.RowHeaders _RowHeaders;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsSalesOpportunity()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsSalesOpportunity()
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
                        if (pVal.ItemUID == "1" && (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                        {
                           _BPCode = _Form.DataSources.DBDataSources.Item("OOPR").GetValue("CardCode", 0).ToString();
                            _UserCd =Convert.ToInt16(_Form.DataSources.DBDataSources.Item("OOPR").GetValue("SlpCode", 0).ToString());
                            _OpprNo = Convert.ToInt16(_Form.DataSources.DBDataSources.Item("OOPR").GetValue("OpprId", 0).ToString());
                            _ModelName = _Form.DataSources.DBDataSources.Item("OOPR").GetValue("Name", 0).ToString();
                        }
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            OpportunityStatus = true;
                         
                        }
                        if (pVal.ItemUID == "Folups")
                        {
                            _Form.Freeze(true);
                            _Form.PaneLevel = 5001;
                            BrowseDataFromTable(_Form.DataSources.DBDataSources.Item("OOPR").GetValue("OpprId", 0).ToString());
                            _Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            _Form.Freeze(false);
                        }
                        break;

                    #endregion

                    #region et_KEY_DOWN
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "56" && pVal.CharPressed == 9 && pVal.ColUID == "U_SchDays")
                        {
                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
                            int Days = Convert.ToInt16(((SAPbouiCOM.EditText)_Matrix.Columns.Item("U_SchDays").Cells.Item(pVal.Row).Specific).Value);
                            string oSQL = "select add_days(CURRENT_DATE, " + Days + ") as \"DT1\" from dummy";
                            Utilities.ExecuteSQL(ref _RS1, oSQL);
                            DateTime Dt = DateTime.Parse(_RS1.Fields.Item("DT1").Value.ToString());
                            ((SAPbouiCOM.EditText)_Matrix.Columns.Item("U_SchDt").Cells.Item(pVal.Row).Specific).Value = Dt.ToString("yyyyMMdd");
                            
                            //SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("U_ActNo");
                            //_Column.Editable = true;
                            //SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
                            //_LinkedButton.LinkedObjectType = "33";
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
                    #region et_ITEM_PRESSED
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && pVal.Action_Success == true && pVal.InnerEvent == false && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            AddFollowUpActivty(_BPCode, _UserCd, _OpprNo.ToString(), _ModelName);
                        }
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && OpportunityStatus == true)
                        {
                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("56").Specific;
                            string oSQL = "select \"Descript\",\"U_ActNo\",IFNULL(\"U_SchDays\",0) as \"U_SchDays\",\"U_SchDt\",\"Line\",\"U_AssTo\",\"U_Remarks\" from OPR1 INNER JOIN OOST ON OPR1.\"Step_Id\" = OOST.\"Num\" " +
                                          "where \"OpprId\" = '" + _OpprNo + "' AND \"Line\" = (select Max(\"Line\") from OPR1 where \"OpprId\" = '" + _OpprNo + "')";
                            Utilities.ExecuteSQL(ref _RS2, oSQL);
                            if (Convert.ToInt16(_RS2.Fields.Item("U_SchDays").Value.ToString()) > 0)
                                ADDUpdateFollowUpActivty(_BPCode, _UserCd, _OpprNo.ToString(), _RS2.Fields.Item("Descript").Value.ToString(), _RS2.Fields.Item("U_ActNo").Value.ToString(), _RS2.Fields.Item("U_SchDt").Value.ToString(), _RS2.Fields.Item("Line").Value.ToString(), _RS2.Fields.Item("U_AssTo").Value.ToString(), _RS2.Fields.Item("U_Remarks").Value.ToString());
                            OpportunityStatus = false;


                        }
                        break;
                    #endregion

                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "GrdFol" && pVal.ColUID == "In Active" && pVal.Action_Success == true && pVal.InnerEvent == false)
                        {

                            _Grid = (SAPbouiCOM.Grid)_Form.Items.Item("GrdFol").Specific;
                            if (_Grid.DataTable.Columns.Item("In Active").Cells.Item(pVal.Row).Value == "Y")
                            {
                                if (Utilities.Application.SBO_Application.MessageBox("Are you sure you want close Followup", 1, "OK", "Cancel") == 1)
                                {
                                    ClosedActivty(_Grid.DataTable.Columns.Item("Followup No").Cells.Item(pVal.Row).Value.ToString());
                                }
                                else
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            _Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        }
                        break;
                }
            }

            #endregion

        }

        #endregion

        #region MENU EVENT
        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        {
            #region BEFORE ACTION
            if (pVal.BeforeAction)
            {

            }
            #endregion

            #region AFTER ACTION
            else
            {
                try
                {
                    Form.Freeze(true);
                    switch (pVal.MenuUID)
                    {
                        case Constants.System_Menus.mnu_FIRST:
                            BrowseDataFromTable(_Form.DataSources.DBDataSources.Item("OOPR").GetValue("OpprId", 0).ToString());
                            break;

                        case Constants.System_Menus.mnu_LAST:
                            BrowseDataFromTable(_Form.DataSources.DBDataSources.Item("OOPR").GetValue("OpprId", 0).ToString());
                            break;

                        case Constants.System_Menus.mnu_NEXT:
                            BrowseDataFromTable(_Form.DataSources.DBDataSources.Item("OOPR").GetValue("OpprId", 0).ToString());
                            break;

                        case Constants.System_Menus.mnu_PREVIOUS:
                            BrowseDataFromTable(_Form.DataSources.DBDataSources.Item("OOPR").GetValue("OpprId", 0).ToString());
                            break;
                        case Constants.System_Menus.mnu_FIND:

                            break;
                        case Constants.System_Menus.mnu_ADD:

                            break;
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
                finally { Form.Freeze(false); }
            }
            #endregion
        }
        #endregion    

        #region FORM DATA EVENT
        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            #region AFTER ACTION
            if (!BusinessObjectInfo.BeforeAction)
            {
                switch (BusinessObjectInfo.EventType)
                {
                    #region DATA LOAD
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        try
                        {
                            _Form.Freeze(true);
                            BrowseDataFromTable(_Form.DataSources.DBDataSources.Item("OOPR").GetValue("OpprId", 0).ToString());

                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally
                        {
                            _Form.Freeze(false);
                        }
                        break;
                    #endregion
                }
            }
            #endregion

            #region BEFORE ACTION
            else
            {
                switch (BusinessObjectInfo.EventType)
                {
                    #region DATA LOAD
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:

                        break;
                    #endregion


                }
            }
            #endregion
        }
        #endregion

        #region ADD FOLLOWUP ACTIVITY

        private void AddFollowUpActivty(string oCardCode, int UserId,string OppID,string _Model)
        {
            SAPbobsCOM.Contacts oContacts = (SAPbobsCOM.Contacts)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oContacts);
            try
            {
                Utilities.ExecuteSQL(ref _RS, "select \"Code\" from OCLS where \"U_Stage\" = 'Enquiry followup'");
                oContacts.Activity = SAPbobsCOM.BoActivities.cn_Conversation;
                oContacts.ActivityType = -1;//Convert.ToInt32(_RS.Fields.Item("Code").Value.ToString());
                oContacts.Subject = _RS.Fields.Item("Code").Value.ToString();
                oContacts.CardCode = oCardCode;
                oContacts.HandledBy = Utilities.Application.Company.UserSignature;
                oContacts.SalesEmployee = UserId;
                oContacts.Details = "Enquiry Creation : '" + _Model + "' ";
                oContacts.Recontact = Utilities.TodayDateFollowTheeDays();
                oContacts.StartTime = Utilities.TodayCurrentTime();
                oContacts.Notes = "Enquiry Creation : '" + _Model + "' ";
                oContacts.Reminder = SAPbobsCOM.BoYesNoEnum.tYES;
                oContacts.ReminderType = SAPbobsCOM.BoDurations.du_Days;
                
                int RetVal = oContacts.Add();
                if (RetVal != 0)
                {
                    string _errMsg = string.Empty;
                    int _errCode = 0;
                    Utilities.Application.Company.GetLastError(out _errCode, out _errMsg);
                }
                else
                {
                    string _ActivityID = "";
                    Utilities.ExecuteSQL(ref _RS3, "select Max(\"OpprId\") as \"OpprId\" from OOPR where \"CardCode\" = '" + oCardCode + "'");
                    string Date = Utilities.TodayDateFollowTheeDays().ToString("yyyyMMdd");
                    Utilities.Application.Company.GetNewObjectCode(out _ActivityID);
                    Utilities.ExecuteSQL(ref _RS, "update OPR1 set \"U_SchDays\" = 3 , \"U_SchDt\" = '" + Date + "',\"U_ActNo\" = '" + _ActivityID + "' where \"OpprId\" = '" + _RS3.Fields.Item("OpprId").Value.ToString() + "' ");
                    
                }
            }
            catch (Exception ex)
            { 
                Utilities.ShowErrorMessage(ex.Message); 
                //return 0; 
            }
        }

        #endregion

        #region UPDATE FOLLOWUP ACTIVITY

        private void ADDUpdateFollowUpActivty(string oCardCode, int UserId, string OppID,string CommentsDescription,string _ActNo,string ScheduleDt,string LineNo,string _AssignTo,string _Remarks)
        {
            SAPbobsCOM.Contacts oContacts = (SAPbobsCOM.Contacts)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oContacts);
            try
            {
                if (oContacts.GetByKey(Convert.ToInt16(_ActNo)))
                {
                    Utilities.ExecuteSQL(ref _RS, "select \"Code\" from OCLS where \"U_Stage\" = '" + CommentsDescription + "'");
                    oContacts.Activity = SAPbobsCOM.BoActivities.cn_Conversation;
                    oContacts.ActivityType = -1;//Convert.ToInt32(_RS.Fields.Item("Code").Value.ToString());
                    oContacts.Subject = _RS.Fields.Item("Code").Value.ToString();
                    oContacts.CardCode = oCardCode;
                    int test = Utilities.Application.Company.UserSignature;
                    if (_AssignTo != "0")
                        oContacts.HandledBy = Convert.ToInt16(_AssignTo);
                    else
                    oContacts.HandledBy = Utilities.Application.Company.UserSignature;

                    oContacts.SalesEmployee = UserId;
                    //oContacts.Details = CommentsDescription.ToString();
                    oContacts.Details = _Remarks.ToString();
                    oContacts.Recontact =Convert.ToDateTime(ScheduleDt);
                    oContacts.StartTime = Utilities.TodayCurrentTime();
                    oContacts.Notes = CommentsDescription.ToString();
                    //oContacts.Reminder = SAPbobsCOM.BoYesNoEnum.tYES;
                    //oContacts.ReminderType = SAPbobsCOM.BoDurations.du_Days;
                                       
                    int RetVal = oContacts.Update();

                    if (RetVal != 0)
                    {
                        string _errMsg = string.Empty;
                        int _errCode = 0;
                        Utilities.Application.Company.GetLastError(out _errCode, out _errMsg);
                    }
                    else
                    {
                        //string _ActivityID = "";
                        //Utilities.Application.Company.GetNewObjectCode(out _ActivityID);
                        //Utilities.ExecuteSQL(ref _RS, "update OPR1 set \"U_ActNo\" = '" + _ActivityID + "' where \"OpprId\" = '" + _OpprNo + "' AND \"Line\" = '" + LineNo + "' ");

                    }
                }
                else
                {
                    Utilities.ExecuteSQL(ref _RS, "select \"Code\" from OCLS where \"U_Stage\" = '" + CommentsDescription + "'");
                    oContacts.Activity = SAPbobsCOM.BoActivities.cn_Conversation;
                    oContacts.ActivityType = -1;//Convert.ToInt32(_RS.Fields.Item("Code").Value.ToString());
                    oContacts.Subject = _RS.Fields.Item("Code").Value.ToString();
                    oContacts.CardCode = oCardCode;

                    if (_AssignTo != "0")
                        oContacts.HandledBy = Convert.ToInt16(_AssignTo);
                    else
                        oContacts.HandledBy = Utilities.Application.Company.UserSignature;

                    oContacts.SalesEmployee = UserId;
                    //oContacts.Details = CommentsDescription.ToString();
                    oContacts.Details = _Remarks.ToString();
                    oContacts.Recontact = Convert.ToDateTime(ScheduleDt);
                    oContacts.StartTime = Utilities.TodayCurrentTime();
                    oContacts.Notes = CommentsDescription.ToString();
                    //oContacts.Reminder = SAPbobsCOM.BoYesNoEnum.tYES;
                    //oContacts.ReminderType = SAPbobsCOM.BoDurations.du_Days;

                    int RetVal = oContacts.Add();

                    if (RetVal != 0)
                    {
                        string _errMsg = string.Empty;
                        int _errCode = 0;
                        Utilities.Application.Company.GetLastError(out _errCode, out _errMsg);
                    }
                    else
                    {
                        string _ActivityID = "";
                        Utilities.Application.Company.GetNewObjectCode(out _ActivityID);
                        Utilities.ExecuteSQL(ref _RS, "update OPR1 set \"U_ActNo\" = '" + _ActivityID + "' where \"OpprId\" = '" + _OpprNo + "' AND \"Line\" = '" + LineNo + "' ");
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                //return 0; 
            }
        }

        #endregion

        #region UPDATE FOLLOWUP ACTIVITY

        private void ClosedActivty(string _ActNo)
        {
            SAPbobsCOM.Contacts oContacts = (SAPbobsCOM.Contacts)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oContacts);
            try
            {
                if (oContacts.GetByKey(Convert.ToInt16(_ActNo)))
                {
                    oContacts.Closed = SAPbobsCOM.BoYesNoEnum.tYES;
                   
                    int RetVal = oContacts.Update();

                    if (RetVal != 0)
                    {
                        string _errMsg = string.Empty;
                        int _errCode = 0;
                        Utilities.Application.Company.GetLastError(out _errCode, out _errMsg);
                        Utilities.Message(_errMsg.ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    else
                    {
                        Utilities.Message("Operation completed sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }
               
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                //return 0; 
            }
        }

        #endregion

        #region Form Default
        internal void FormDefault()
        {
            try
            {
                _Form.Title = "Inquiry";

                _ExistItem = Form.Items.Item("135");
                _NewItem = Form.Items.Add("Folups", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                _NewItem.Height = _ExistItem.Height;
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width;
                _NewItem.Width = _ExistItem.Width + 100;
                _Folder = (SAPbouiCOM.Folder)_NewItem.Specific;
                _Folder.GroupWith("135");
                _Folder.Caption = "Followup Details";

                _ExistItem = _Form.Items.Item("56");
                _NewItem = _Form.Items.Add("GrdFol", SAPbouiCOM.BoFormItemTypes.it_GRID);
                _NewItem.FromPane = 5001;
                _NewItem.ToPane = 5001;
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.Top = _ExistItem.Top;
                _NewItem.Height = _ExistItem.Height;
                _Grid = (SAPbouiCOM.Grid)_NewItem.Specific;
                _Form.DataSources.DataTables.Add("dtFol");
                _Grid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                _GridDataTable = _Form.DataSources.DataTables.Item("dtFol");
                _ExistItem = _Form.Items.Item("GrdFol");

                string strSQL = "SELECT "
                                + " CAST('' as varchar(50)) as \"Stage\" , "
                                //+ " CAST('' as varchar(50)) as \"Previous Followup No\" , "
                                + " CAST('' as varchar(50)) as \"Followup No\" , "
                                 + " CAST('' as varchar(50)) as \"Followup Date\" , "
                                 + " CAST('' as varchar(50)) as \"Followup Time\" , "
                                 + " CAST('' as varchar(50)) as \"Remarks\" , "
                                 + " CAST('' as varchar(50)) as \"Content\" , "
                                 + " CAST('' as varchar(50)) as \"Assign To\" , "
                                 + " CAST('' as varchar(50)) as \"Status\" ,"
                                 + " CAST('' as varchar(50)) as \"Assign To\", "
                                 + " CAST('' as varchar(1)) as \"In Active\" "
                                 + "FROM DUMMY ";

                _GridDataTable.ExecuteQuery(strSQL);
                _Grid.DataTable = _GridDataTable;
              
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion

        #region BROWSE DATA FROM TABLE
        private bool BrowseDataFromTable(string _DocEntry)
        {
            try
            {
                _Form.Freeze(true);
                string strSQL = string.Empty;
                 strSQL = "select * from (select T2.\"Descript\" as \"Stage\", " +
                          //" T4.\"prevActvty\" as \"Previous Followup No\", " +
                          "T4.\"ClgCode\" as \"Followup No\",T4.\"Recontact\" as \"Followup Date\", " +
                          "T4.\"BeginTime\" as \"Followup Time\",T4.\"Details\" as \"Remarks\",T4.\"Notes\" as \"Content\", " +
                          " case when T4.\"Closed\" = 'Y' then 'Closed' else 'Open' END as \"Status\",T5.\"U_NAME\" as \"Assign To\",T4.\"Closed\" as \"In Active\" FROM OOPR T0 INNER JOIN OPR1 T1 ON T0.\"OpprId\" = T1.\"OpprId\" " +
                          "INNER JOIN OOST T2 ON T1.\"Step_Id\" = T2.\"Num\" INNER JOIN OCLG T3 ON T3.\"ClgCode\" = T1.\"U_ActNo\" INNER JOIN OCLG T4 ON T3.\"ClgCode\" = T4.\"prevActvty\" "+
                          "LEFT OUTER JOIN OUSR T5 ON T5.\"USERID\" = T4.\"AttendUser\"" +
                          "where T1.\"OpprId\" = '" + _DocEntry + "' " +
                          "UNION ALL "+
                          "select T2.\"Descript\" as \"Stage\", " +
                          //"T3.\"prevActvty\" as \"Previous Followup No\", "+
                          "T1.\"U_ActNo\" as \"Followup No\",T3.\"Recontact\" as \"Followup Date\", "+
                          "T3.\"BeginTime\" as \"Followup Time\",T3.\"Details\" as \"Remarks\",T3.\"Notes\" as \"Content\", " +
                          " case when T3.\"Closed\" = 'Y' then 'Closed' else 'Open' END as \"Status\",T4.\"U_NAME\" as \"Assign To\",T3.\"Closed\" as \"In Active\" FROM OOPR T0 INNER JOIN OPR1 T1 ON T0.\"OpprId\" = T1.\"OpprId\" " +
                          "INNER JOIN OOST T2 ON T1.\"Step_Id\" = T2.\"Num\" INNER JOIN OCLG T3 ON T3.\"ClgCode\" = T1.\"U_ActNo\" "+
                          "LEFT OUTER JOIN OUSR T4 ON T4.\"USERID\" = T3.\"AttendUser\"" +
                          "where T1.\"OpprId\" = '" + _DocEntry + "')A order by \"Followup No\" ";
                _GridDataTable.ExecuteQuery(strSQL);
                _Grid.DataTable = _GridDataTable;

                if (_GridDataTable.Rows.Count > 0)
                {
                    // Form.Items.Item("BtnFill").Enabled = false;
                    _NewItem.Enabled = false;
                    _Grid.DataTable = _GridDataTable;
                    _RowHeaders = _Grid.RowHeaders;
                    _RowHeaders.TitleObject.Caption = "#";
                    for (int i = 0; i < _Grid.Rows.Count; i++)
                    {
                        _RowHeaders.SetText(i, (i + 1).ToString());
                    }
                    _Column = (SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(1);
                    _Column.LinkedObjectType = "33";

                    _Column = (SAPbouiCOM.EditTextColumn)_Grid.Columns.Item(8);
                    _Column.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                  
                    _Grid.Columns.Item(0).Editable = false;
                    _Grid.Columns.Item(1).Editable = false;
                    _Grid.Columns.Item(2).Editable = false;
                    _Grid.Columns.Item(3).Editable = false;
                    _Grid.Columns.Item(4).Editable = false;
                    _Grid.Columns.Item(5).Editable = false;
                    _Grid.Columns.Item(6).Editable = false;
                    _Grid.Columns.Item(7).Editable = false;
                   // _Grid.Columns.Item(8).Editable = false;
                    _Grid.AutoResizeColumns();
                }

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                return false;
            }
            finally
            {
                _Form.Freeze(false);
                _Form.Update();
                GC.Collect();
            }
            return true;
        }
        #endregion

    }
}
