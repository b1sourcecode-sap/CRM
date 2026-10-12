using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using SAPbobsCOM;

namespace CRM
{
    class ClsEquipmentCard : Base
    {

        #region VARIABLE DECLARATION

        private SAPbobsCOM.Recordset _Rs, _RS, _RSRptQuery;
        SAPbouiCOM.Item _ExistItem, _NewItem, _NewItem1;
        SAPbouiCOM.Grid _Grid = null;
        SAPbouiCOM.Folder _Folder;
        private SAPbouiCOM.DataTable _GridDataTable;
        SAPbouiCOM.Button _Button, _Button1;
        private SAPbouiCOM.RowHeaders _RowHeaders;
        private SAPbouiCOM.GridColumn  oComboBoxColumn;
        private SAPbouiCOM.ComboBox _ComboBox;
        SAPbobsCOM.UserTable  UDTData;
        public static string _ServiceCode;
        public static int _GridRowCount,_SelectedRow=-1;
        public int _rtnVal;
        public static Boolean blnClose = false;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsEquipmentCard()
            : base()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsEquipmentCard()
        {
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
                    _Form.Freeze(true);
                    switch (pVal.MenuUID)
                    {
                        case Constants.System_Menus.mnu_FIRST:

                            BrowseDataFromTable();
                            _Form.Items.Item("BtnFill").Enabled = false;
                            _Form.Items.Item("BtnOk").Enabled = false;
                            _Form.Items.Item("CreatBP").Enabled = false;
                            break;

                        case Constants.System_Menus.mnu_LAST:
                            _Form.Items.Item("BtnFill").Enabled = false;
                            _Form.Items.Item("BtnOk").Enabled = false;
                            _Form.Items.Item("CreatBP").Enabled = false;
                            break;

                        case Constants.System_Menus.mnu_NEXT:
                            _Form.Items.Item("BtnFill").Enabled = false;
                            _Form.Items.Item("BtnOk").Enabled = false;
                            _Form.Items.Item("CreatBP").Enabled = false;
                            break;

                        case Constants.System_Menus.mnu_PREVIOUS:
                            // _Form.Items.Item("BtnFill").Enabled = false;
                            // _Form.Items.Item("BtnOk").Enabled = false;
                            //  _Form.Items.Item("CreatBP").Enabled = false;
                            break;
                        case Constants.System_Menus.mnu_FIND:
                            _Form.Items.Item("BtnFill").Enabled = false;
                            _Form.Items.Item("BtnOk").Enabled = false;
                            _Form.Items.Item("CreatBP").Enabled = true;
                            break;
                        case Constants.System_Menus.mnu_ADD:
                            //_Form.Items.Item("BtnFill").Enabled = false;
                            _Form.Items.Item("BtnOk").Enabled = false;
                            _Form.Items.Item("CreatBP").Enabled = true;
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
                finally { _Form.Freeze(false ); }
            }
            #endregion
        }
        #endregion    

        #region ITEM EVENTS
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region BEFORE ACTION
            if (pVal.Before_Action)
            {
                switch (pVal.EventType)
                {
                    #region et_VALIDATE
                    //case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED://Commented On 13.06.2022
                       
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemUID == "GrdServSch")
                        {                           
                        }                                               
                        break;

                    #endregion

                    #region et_PICKER_CLICKED

                    case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
                        try
                        {
                            if (pVal.ItemUID == "GrdServSch")
                            {
                                int RowID = pVal.Row;
                                if (CheckStatus(pVal.Row) == true)
                                {
                                    if ((_Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Completed" || _Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Cancelled") && _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(pVal.Row).Value.ToString() != "" && pVal.CharPressed != Constants.Keys.TAB)//&& pVal .ColUID =="1")
                                    {
                                        Utilities.ShowWarningMessage("Scheme has been completed.");
                                        oBubbleEvent = false;
                                    }
                                }
                            }
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

                    #region et_COMBO_SELECT

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        try
                        {
                            if (pVal.ItemUID == "GrdServSch")
                            {
                                if (CheckStatus(pVal.Row) == true)
                                {
                                    if ((_Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Completed" || _Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Cancelled") && _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(pVal.Row).Value.ToString() != "" && pVal.CharPressed != Constants.Keys.TAB)//&& pVal .ColUID =="1")
                                    {
                                        DateTime DateValue = Convert.ToDateTime(_Grid.DataTable.Columns.Item("ServiceDueDate").Cells.Item(pVal.Row).Value.ToString());
                                        DateTime _CurrentDate = DateTime.Now;
                                        if (DateValue <= _CurrentDate)
                                        {
                                            //_Grid.CommonSetting.SetRowBackColor(pVal.Row + 1, 255 - 255 - 255);
                                            Utilities.ShowWarningMessage("Scheme has been completed.");
                                            oBubbleEvent = false;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally
                        {
                            _Form.Freeze(false);
                            _Form.Visible = true;
                        }
                    break;

                    #endregion

                    #region  KEY DOWN
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                            if (pVal.ItemUID == "GrdServSch" && (pVal.ColUID == "ServiceSerialNo" || pVal.ColUID == "ServiceDueDate"
                            || pVal.ColUID == "ServiceDoneDate" || pVal.ColUID == "ServiceType" || pVal.ColUID == "Mileage" || pVal.ColUID == "DeliveryDate"
                            || pVal.ColUID == "SalesDeliveryDate" || pVal.ColUID == "Status") && pVal.CharPressed != Constants.Keys.TAB)
                        {
                            _Form.Freeze(true);
                            try
                            {
                                if (CheckStatus(pVal.Row) == true)
                                {
                                    //=======================
                                    if ((_Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Completed" ||
                                        _Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Cancelled") &&
                                        _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(pVal.Row).Value.ToString() != "" &&
                                        pVal.CharPressed != Constants.Keys.TAB)
                                    {
                                        //_Grid.CommonSetting.SetRowBackColor(pVal.Row + 1, 255 - 255 - 255);
                                        Utilities.ShowWarningMessage("Scheme has been completed.");
                                        oBubbleEvent = false;
                                    }
                                }
                                //=======================
                            }
                            finally { _Form.Freeze(false); }
                        }
                        break;
                      
                    #endregion
                  
                }
            }
            #endregion

            #region AFTER ACTION
            else
            {
                switch (pVal.EventType)
                {
                    #region et_CHOOSE_FROM_LIST

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST :
                        try
                        {
                            _Form.Items.Item("CreatBP").Enabled = false;
                            _Form.Items.Item("BtnFill").Enabled = false;
                            _Form.Items.Item("BtnOk").Enabled = false;
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

                    #region et_CLICK

                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        try
                        {
                            _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                            if (pVal.ItemUID == "SSchedule")
                            {
                                _Form.PaneLevel = 7;
                                if (_Form.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && ((SAPbouiCOM.EditText)this._Form.Items.Item("45").Specific).Value != "")
                                {
                                    BrowseDataFromTable();
                                    _Form.Items.Item("CreatBP").Enabled = false;
                                }

                            }
                            if (pVal.ItemUID == "GrdServSch")
                            {
                                if (pVal.ColUID == "#")
                                    _SelectedRow = pVal.Row;
                                else
                                    _SelectedRow = -1;
                            }
                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally
                        {
                            _Form.Freeze(false);
                            _Form.Visible = true;
                        }
                        break;

                    #endregion

                    #region et_ITEM_PRESSED

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        try
                        {
                            if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            {
                                if (pVal.ItemUID == "CreatBP")
                                {
                                    string oCardCode = ((SAPbouiCOM.EditText)this._Form.Items.Item("48").Specific).Value.Trim();
                                    string oSQL = "Select Count(*) from OINS where \"customer\"='" + oCardCode.Trim() + "'";
                                    Utilities.ExecuteSQL(ref _Rs, oSQL);
                                    if (_Rs.Fields.Item(0).Value.ToString().Trim() == "0")
                                        _Form.Items.Item("CreatBP").Enabled = true;
                                    else
                                        _Form.Items.Item("CreatBP").Enabled = false;
                                }
                            }
                        if (pVal.ItemUID == "CreatBP")
                        {
                            string oCardCode = ((SAPbouiCOM.EditText)this._Form.Items.Item("48").Specific).Value.Trim();
                            string oEngineSerialNo = ((SAPbouiCOM.EditText)this._Form.Items.Item("44").Specific).Value.Trim();
                            if (oEngineSerialNo == "")
                            {
                                Utilities.ShowErrorMessage("Enter Serial Number");
                                _Form.Items.Item("CreatBP").Enabled = true ;
                                return;
                            }
                            else if (oCardCode == "")
                            {
                                Utilities.ShowErrorMessage("Enter CardCode :");
                                _Form.Items.Item("CreatBP").Enabled = true;
                                return;
                            }
                            else
                            {
                                CreateBP(oCardCode, oEngineSerialNo);
                               // _Form.Items.Item("CreatBP").Enabled = false;
                            }
                        }

                        if (pVal.ItemUID == "BtnFill")
                        {
                            if (_Form.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            {
                                FillGrid();
                                SaveData();
                                _Form.Items.Item("BtnFill").Enabled = false;
                                _Form.Close();
                            }
                        }
                        if (pVal.ItemUID == "BtnOk")
                        {
                           if (_Button.Caption == "Update")
                            {
                                UpdateData();
                                if (pVal.Action_Success)
                                {
                                    //_Button.Caption = "OK";
                                    _Form.Items.Item("BtnFill").Enabled = false;
                                    _NewItem.Enabled = false;
                                }
                            }
                        }

                        if (pVal.ItemUID == "3")
                        {
                            _Form.Items.Item("BtnOk").Enabled = false;
                            _Form.Items.Item("BtnFill").Enabled = false;
                            //_Form.Items.Item("CreatBP").Enabled = false;
                        }
                        if (pVal.ItemUID == "PrntSmtCrd")
                        {

                            string strMenuID = null;
                            _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                           
                            string _EngineNo = ((SAPbouiCOM.EditText)_Form.Items.Item("44").Specific).Value.ToString().Trim();
                            string RptQuery = "select \"MenuUID\" from OCMN where \"ObjectKey\"='RCRI0009'";
                            Utilities.ExecuteSQL(ref _RSRptQuery, RptQuery);

                            strMenuID = _RSRptQuery.Fields.Item("MenuUID").Value.ToString();
                            Utilities.Application.SBO_Application.ActivateMenuItem(strMenuID);
                            _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                            SAPbouiCOM.EditText oEditDocEntry = (SAPbouiCOM.EditText)_Form.Items.Item("1000003").Specific;
                            oEditDocEntry.Value = _EngineNo.ToString();
                            _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            blnClose = true;
                        }
                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally
                        {
                            //_Form.Freeze(false);
                            //_Form.Visible = true;
                        }
                        break;

                    #endregion

                    #region  KEY DOWN

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        try
                        {
                            if (pVal.ItemUID == "GrdServSch" && (pVal.ColUID == "ServiceSerialNo" || pVal.ColUID == "ServiceDueDate"
                            || pVal.ColUID == "ServiceDoneDate" || pVal.ColUID == "ServiceType" || pVal.ColUID == "Mileage" || pVal.ColUID == "DeliveryDate"
                            || pVal.ColUID == "SalesDeliveryDate" || pVal.ColUID == "Status") && pVal.CharPressed != Constants.Keys.TAB)
                        {
                              if (CheckStatus(pVal.Row) == true)
                                {
                                    //=======================
                                    if ((_Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Completed" ||
                                    _Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Cancelled") &&
                                    _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(pVal.Row).Value.ToString() != "" &&
                                    pVal.CharPressed != Constants.Keys.TAB)
                                    {
                                       // _Grid.CommonSetting.SetRowBackColor(pVal.Row + 1, 255 - 250 - 240);
                                        Utilities.ShowWarningMessage("Scheme has been completed.");
                                        oBubbleEvent = false;
                                    }
                                }
                        }
                        if (pVal.ItemUID == "GrdServSch" && (pVal.ColUID == "ServiceSerialNo" || pVal.ColUID == "ServiceDueDate"
                           || pVal.ColUID == "ServiceDoneDate" || pVal.ColUID == "ServiceType" || pVal.ColUID == "Mileage" || pVal.ColUID == "DeliveryDate"
                           || pVal.ColUID == "SalesDeliveryDate" || pVal.ColUID == "Status") && pVal.CharPressed == Constants.Keys.TAB)
                        {
                                _Form.Freeze(true);
                            try
                            {
                                
                                //=======================
                                _Button.Caption = "Update";
                                _NewItem.Enabled = true;
                                _Form.Items.Item("BtnFill").Enabled = false;
                                if (_Grid.DataTable.Columns.Item("ServiceSerialNo").Cells.Item(pVal.Row).Value.ToString() != ""   &&
                                // _Grid.DataTable.Columns.Item("ServiceDueDate").Cells.Item(pVal.Row).Value.ToString() != ""   &&
                                //_Grid.DataTable.Columns.Item("ServiceDoneDate").Cells.Item(pVal.Row).Value.ToString() != ""   &&
                                //_Grid.DataTable.Columns.Item("ServiceType").Cells.Item(pVal.Row).Value.ToString() != ""   &&
                                //_Grid.DataTable.Columns.Item("Mileage").Cells.Item(pVal.Row).Value.ToString() != ""   &&
                                //_Grid.DataTable.Columns.Item("DeliveryDate").Cells.Item(pVal.Row).Value.ToString() != "" &&
                                //_Grid.DataTable.Columns.Item("ServiceDeliveryDate").Cells.Item(pVal.Row).Value.ToString() != ""  &&
                                _Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() != "" && pVal.ColUID == "Status" && pVal.Row==_Grid.DataTable.Rows.Count-1
                                )
                                {
                                    _Grid.DataTable.Rows.Add();
                                }
                                _RowHeaders.TitleObject.Caption = "#";
                                for (int i = 0; i < _Grid.Rows.Count; i++)
                                {
                                    _RowHeaders.SetText(i, (i + 1).ToString());                                    
                                }

                              
                                if (_Grid.DataTable.Rows.Count == 1 && _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(0).Value.ToString() == "")
                                {
                                    _Form.Items.Item("BtnFill").Enabled = false;
                                }
                            }
                            finally { _Form.Freeze(false); _Form.Items.Item("CreatBP").Enabled = false; }
                        }
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

                    #region et_PICKER_CLICKED

                    case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
                        //if (pVal.ItemUID == "GrdServSch")// && pVal.ColUID == "Status")
                        //{
                        //    int RowID = pVal.Row;
                        //    if ((_Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Completed" || _Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Cancelled") && _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(pVal.Row).Value.ToString() != "" && pVal.CharPressed != Constants.Keys.TAB)//&& pVal .ColUID =="1")
                        //    {
                        //        _Grid.CommonSetting.SetRowBackColor(pVal.Row + 1, 255 - 250 - 240);
                        //        Utilities.ShowWarningMessage("Scheme has been completed.");
                        //        // _Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //        //_Grid.SetCellFocus(pVal.Row + 1, 2);
                        //        oBubbleEvent = false;
                        //       // return;
                        //    }
                        //}

                        break;

                    #endregion

                    #region et_COMBO_SELECT

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        try
                        {
                            _Form.Freeze(true);
                            if (pVal.ItemUID == "GrdServSch")
                            {
                                int RowID = pVal.Row;

                                if ((_Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Completed" || _Grid.DataTable.Columns.Item("Status").Cells.Item(pVal.Row).Value.ToString() == "Cancelled") && _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(pVal.Row).Value.ToString() != "" && pVal.CharPressed != Constants.Keys.TAB)
                                {
                                    if (CheckStatus(pVal.Row) == true)
                                    {
                                        // _Grid.CommonSetting.SetRowBackColor(pVal.Row + 1, 255 - 250 - 240);
                                        Utilities.ShowWarningMessage("Scheme has been completed.");
                                        oBubbleEvent = false;
                                    }
                                }
                                _Form.Items.Item("CreatBP").Enabled = false;
                            }
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

                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        if(pVal.ItemUID == "U_SMCardNo" && pVal.Action_Success == true)
                        {
                            if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            {
                                _Form.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }

                        break;
                 
                }
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
                            BrowseDataFromTable();
                            _Form.Items.Item("CreatBP").Enabled = false;
                            if (_Grid.DataTable.Rows.Count > 0 && _Grid.DataTable.Columns.Item("Code").Cells.Item(_Grid.DataTable.Rows.Count - 1).Value.ToString() != "")
                            {
                                _Form.Items.Item("BtnFill").Enabled = false;
                            }
                            else
                                _Form.Items.Item("BtnFill").Enabled = true;

                            if (_Grid.DataTable.Rows.Count == 1 && _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(0).Value.ToString() != "")
                                _Grid.DataTable.Rows.Add();
                            else if (_Grid.DataTable.Rows.Count > 1)
                                _Grid.DataTable.Rows.Add();

                            _RowHeaders.TitleObject.Caption = "#";
                            for (int i = 0; i < _Grid.Rows.Count; i++)
                            {
                                _RowHeaders.SetText(i, (i + 1).ToString());                                
                                if (this.CheckStatus(i) == true)
                                {
                                    if ((_Grid.DataTable.Columns.Item("Status").Cells.Item(i).Value.ToString() == "Completed" ||
                                        _Grid.DataTable.Columns.Item("Status").Cells.Item(i).Value.ToString() == "Cancelled") &&
                                        _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString() != "")
                                    {
                                        oBubbleEvent = false;
                                    }
                                }
                            }
                            _Form.Items.Item("CreatBP").Enabled = false;
                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally { _Form.Freeze(false); }
                        break;
                    #endregion

                    #region DATA UPDATE
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                        try
                        {
                            _Form.Freeze(true);
                            _RowHeaders.TitleObject.Caption = "#";
                            for (int i = 0; i < _Grid.Rows.Count; i++)
                            {
                                _RowHeaders.SetText(i, (i + 1).ToString());
                                //=======================
                                if (this.CheckStatus(i) == true)
                                {
                                    if ((_Grid.DataTable.Columns.Item("Status").Cells.Item(i).Value.ToString() == "Completed" ||
                                        _Grid.DataTable.Columns.Item("Status").Cells.Item(i).Value.ToString() == "Cancelled") &&
                                        _Grid.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value.ToString() != "")
                                    {
                                        oBubbleEvent = false;
                                    }
                                }
                            }
                            _Form.Items.Item("CreatBP").Enabled = false;
                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally { _Form.Freeze(false); }
                    break;
                    #endregion

                    #region FORM RESIZE
                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE :
                    try
                    {
                        _Form.Freeze(true);
                        _Form.Items.Item("CreatBP").Enabled = false;
                        _Grid.Columns.Item("Code").Width = 0;
                        _Grid.Columns.Item("ItemCode").Width = 0;
                        _Grid.Columns.Item("SerialNo").Width = 0;
                        _Grid.Columns.Item("CustomerCode").Width = 0;
                    }
                    catch (Exception ex)
                    {
                        Utilities.ShowErrorMessage(ex.Message);
                    }
                    finally { _Form.Freeze(false ); }
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
                        try
                        {
                            _Form.Freeze(true);
                            _Form.Items.Item("BtnFill").Enabled = false;
                            _Form.Items.Item("CreatBP").Enabled = false;
                            _NewItem.Enabled = false;                           
                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally { _Form.Freeze(false); }
                        break;
                    #endregion

                    #region DATA ADD
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        try
                        {
                            _Form.Freeze(true);
                            _Form.Items.Item("CreatBP").Enabled = false;
                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally { _Form.Freeze(false); }
                        break;
                    #endregion

                    #region DATA UPDATE
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                        try
                        {
                            _Form.Freeze(true);
                        _Form.Items.Item("CreatBP").Enabled = false;
                        _Form.Items.Item("BtnFill").Enabled = false;
                        }
                         catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally { _Form.Freeze(false); }
                        break;
                    #endregion

                    #region FORM RESIZE
                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                        try
                        {
                            _Form.Freeze(true);
                            _Form.Items.Item("CreatBP").Enabled = false;
                            _Grid.Columns.Item("Code").Width = 0;
                            _Grid.Columns.Item("ItemCode").Width = 0;
                            _Grid.Columns.Item("SerialNo").Width = 0;
                            _Grid.Columns.Item("CustomerCode").Width = 0;
                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally { _Form.Freeze(false); }
                        break;
                    #endregion
                }
            }
            #endregion
        }
        #endregion

        #region FUNCTION

        #region FORM DEFAULT
        public void FormDefault()
        {
            try
            {
                _ExistItem = _Form.Items.Item("39");
                _NewItem = _Form.Items.Add("SSchedule", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                _NewItem.Left = _ExistItem.Left;
                _NewItem.Width = _ExistItem.Width + 13;
                _NewItem.Top = _ExistItem.Top;
                _NewItem.Height = 21;
                _Folder = (SAPbouiCOM.Folder)_NewItem.Specific;
                _Folder.GroupWith("39");
                _Folder.Caption = "Service Schedule";
                this.AddControlToCustomFolder();

                if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    _Form.Items.Item("CreatBP").Enabled = true;
                    _Form.Items.Item("BtnFill").Enabled = false ;
                }
                else
                    _Form.Items.Item("CreatBP").Enabled = false;
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
                _Form.Visible = true;
            }
        }
        #endregion

        #region FILL GRID

        private bool FillGrid()
        {
            try
            {
                _Form.Freeze(true);
                string oDeliveryDate=string.Empty;
                //==Required information=================================================================================================
                string oItemCode = ((SAPbouiCOM.EditText)this._Form.Items.Item("45").Specific).Value;
                string oCustumerCode = ((SAPbouiCOM.EditText)this._Form.Items.Item("48").Specific).Value;
                string  oSerialNo = ((SAPbouiCOM.EditText)this._Form.Items.Item("44").Specific).Value;
                //int oInvoiceDocNum = Convert.ToInt32(!string.IsNullOrEmpty(((SAPbouiCOM.EditText)this._Form.Items.Item("57").Specific).Value));

                string oInvoiceDocNum = ((SAPbouiCOM.EditText)this._Form.Items.Item("57").Specific).Value;
                if (oInvoiceDocNum.Trim() != "")
                {
                    string oSQL = "Select To_Dats(\"DocDate\") from OINV where \"DocEntry\"=(Select \"DocEntry\" from OINV where \"DocNum\"=" + oInvoiceDocNum + ")";
                    Utilities.ExecuteSQL(ref _Rs, oSQL);
                    if (_Rs.RecordCount > 0)
                        oDeliveryDate = _Rs.Fields.Item(0).Value.ToString().Trim();
                }
                    //=========================================================================================================================

                    string SQL = "Select \"SuppCatNum\" from OITM where \"ItemCode\"='" + oItemCode.Trim() + "'";
                    Utilities.ExecuteSQL(ref _Rs, SQL);
                    if (_Rs.RecordCount > 0 && _Rs.Fields.Item(0).Value.ToString() != "")
                    {
                        _ServiceCode = _Rs.Fields.Item(0).Value.ToString();
                        string strSQLQuery = string.Empty;
                        //ADD_MONTHS (TO_VARCHAR (TO_DATE('20091231'), 'YYYY-MM-DD'),"U_Months" )
                        if (oInvoiceDocNum.Trim() != "")
                        {
                            strSQLQuery = "Select CAST('' as varchar(50)) as \"Code\",'" + oItemCode + "' as \"ItemCode\",'" + oCustumerCode + "' as \"CustomerCode\",'" + oSerialNo + "' as \"SerialNo\",\"U_SNo\" as \"ServiceSerialNo\",ADD_MONTHS (TO_VARCHAR(TO_DATE ('" + oDeliveryDate + "'), 'YYYY-MM-DD'),\"U_Months\" ) as \"ServiceDueDate\",cast('' as varchar(50)) as \"ServiceDoneDate\", " +
                                            " \"U_Type\" as \"ServiceType\",\"U_Milage\" as \"Mileage\",\"U_Months\" as \"Month\",TO_DATE('" + oDeliveryDate + "') as \"DeliveryDate\",TO_DATE('" + oDeliveryDate + "') as \"SalesDeliveryDate\" ,'Pending' as \"Status\" " +
                                            "from \"@SRVC_MSTR\" T0 inner join \"@SRVC_DETAIL\" T1 on T0.\"Code\"=T1.\"Code\" where T0.\"U_SrvcCod\"='" + _Rs.Fields.Item(0).Value.ToString() + "'";
                        }
                        else
                        {

                            strSQLQuery = "Select CAST('' as varchar(50)) as \"Code\",'" + oItemCode + "' as \"ItemCode\",'" + oCustumerCode + "' as \"CustomerCode\",'" + oSerialNo + "' as \"SerialNo\",\"U_SNo\" as \"ServiceSerialNo\",cast('' as varchar(50))  as \"ServiceDueDate\",cast('' as varchar(50)) as \"ServiceDoneDate\", " +
                                           " \"U_Type\" as \"ServiceType\",\"U_Milage\" as \"Mileage\",\"U_Months\" as \"Month\",cast('' as varchar(50)) as \"DeliveryDate\",cast('' as varchar(50)) as \"SalesDeliveryDate\" ,'Pending' as \"Status\" " +
                                           "from \"@SRVC_MSTR\" T0 inner join \"@SRVC_DETAIL\" T1 on T0.\"Code\"=T1.\"Code\" where T0.\"U_SrvcCod\"='" + _Rs.Fields.Item(0).Value.ToString() + "'";
                        }
                        _GridDataTable.ExecuteQuery(strSQLQuery);
                        _Grid.DataTable = _GridDataTable;
                        CheckScheme();
                    }
                    if (_GridDataTable.Rows.Count > 0)
                    {
                        _Grid.DataTable = _GridDataTable;
                        _RowHeaders = _Grid.RowHeaders;
                        _RowHeaders.TitleObject.Caption = "#";
                        for (int i = 0; i < _Grid.Rows.Count; i++)
                        {
                            _RowHeaders.SetText(i, (i + 1).ToString());
                            _Grid.DataTable.SetValue("Status", i, "Pending");
                        }
                        _Grid.AutoResizeColumns();

                        _Grid.Columns.Item("Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                        ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Status")).ValidValues.Add("Pending", "Pending");
                        ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Status")).ValidValues.Add("Cancelled", "Cancelled");
                        ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Status")).ValidValues.Add("Completed", "Completed");
                                          

                        _Grid.Columns.Item("ServiceType").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                        ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("ServiceType")).ValidValues.Add("Free", "Free");
                        ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("ServiceType")).ValidValues.Add("Paid", "Paid");
                        ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("ServiceType")).ValidValues.Add("Scheme", "Scheme");
                        this._Grid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                        _Grid.Columns.Item("Code").Width = 0;
                        _Grid.Columns.Item("ItemCode").Width = 0;
                        _Grid.Columns.Item("SerialNo").Width = 0;
                        _Grid.Columns.Item("CustomerCode").Width = 0;
                        _Grid.Columns.Item("ServiceSerialNo").TitleObject.Caption = "Service Serial No";
                        _Grid.Columns.Item("ServiceDueDate").TitleObject.Caption = "Service Due Date";
                        _Grid.Columns.Item("ServiceDoneDate").TitleObject.Caption = "Service Done Date";
                        _Grid.Columns.Item("ServiceType").TitleObject.Caption = "Service Type";
                        _Grid.Columns.Item("Mileage").TitleObject.Caption = "Mileage";
                        //_Grid.Columns.Item("Month").TitleObject.Caption = "Month";
                        _Grid.Columns.Item("DeliveryDate").TitleObject.Caption = "Delivery Date";
                        _Grid.Columns.Item("SalesDeliveryDate").TitleObject.Caption = "Sales Delivery Date";
                        _Grid.Columns.Item("Status").TitleObject.Caption = "Status";
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
            }
            return true;
        }

        private void CheckScheme()
        {
            try
            {
                _Form.Freeze(true);
                string oDeliveryDate = string.Empty;
                string oItemCode = ((SAPbouiCOM.EditText)this._Form.Items.Item("45").Specific).Value;
                string oCustumerCode = ((SAPbouiCOM.EditText)this._Form.Items.Item("48").Specific).Value;
                string oSerialNo = ((SAPbouiCOM.EditText)this._Form.Items.Item("44").Specific).Value;
                string oInvoiceDocNum = ((SAPbouiCOM.EditText)this._Form.Items.Item("57").Specific).Value;
                if (oInvoiceDocNum.Trim() != "")
                {
                    string oSQL = "Select To_Dats(\"DocDate\") from OINV where \"DocEntry\"=(Select \"DocEntry\" from OINV where \"DocNum\"=" + oInvoiceDocNum + ")";
                    Utilities.ExecuteSQL(ref _Rs, oSQL);
                    if (_Rs.RecordCount > 0)
                        oDeliveryDate = _Rs.Fields.Item(0).Value.ToString().Trim();
                }

                string strSQLQuery = " Select CAST('' as varchar(50)) as \"Code\",'" + oItemCode + "' as \"ItemCode\",'" + oCustumerCode + "' as \"CustomerCode\"," +
                                    "'" + oSerialNo + "' as \"SerialNo\",CAST('' as varchar(50)) as \"ServiceSerialNo\",TO_DATS (\"U_ToDt\") " +
                                    "  as \"ServiceDueDate\",cast('' as varchar(50)) as \"ServiceDoneDate\",\"U_TotSrvc\" as \"TotalService\"," +
                                    " 'Scheme' as \"ServiceType\",cast('' as varchar(50)) as \"Mileage\",cast('' as varchar(50)) as \"Month\"," +
                                    " TO_DATE('" + oDeliveryDate + "') as \"DeliveryDate\",TO_DATE('" + oDeliveryDate + "') as \"SalesDeliveryDate\" ,'Pending' as \"Status\" " +
                                    " from \"@SCHM_MSTR\" T0 inner join \"@SCHM_DETAIL\" T1 on T0.\"Code\"=T1.\"Code\" "+
                                    " where TO_DATE('" + oDeliveryDate + "') between TO_DATE(T1.\"U_FromDt\") and TO_DATE(T1.\"U_ToDt\")";
                //_GridDataTable.ExecuteQuery(strSQLQuery);
                //_Grid.DataTable = _GridDataTable;
                Utilities.ExecuteSQL(ref _Rs, strSQLQuery);
                if (_Rs.RecordCount > 0)
                {

                    for (int i = 0; i < _Rs.RecordCount; i++)
                    {
                        //if(Convert.ToInt32(_Rs.Fields.Item("TotalService").Value.ToString().Trim())>0)
                        //{
                        for (int j = 0; j < Convert.ToInt32(_Rs.Fields.Item("TotalService").Value.ToString().Trim()); j++)
                        {
                            _Grid.DataTable.Rows.Add();
                            int lastRowIndex = _Grid.DataTable.Rows.Count - 1;
                            _Grid.DataTable.SetValue("Code", lastRowIndex, _Rs.Fields.Item("Code").Value.ToString().Trim());
                            _Grid.DataTable.SetValue("ItemCode", lastRowIndex, _Rs.Fields.Item("ItemCode").Value.ToString().Trim());
                            _Grid.DataTable.SetValue("CustomerCode", lastRowIndex, _Rs.Fields.Item("CustomerCode").Value.ToString().Trim());
                            _Grid.DataTable.SetValue("SerialNo", lastRowIndex, oSerialNo);
                            _Grid.DataTable.SetValue("ServiceSerialNo", lastRowIndex, lastRowIndex);
                            if (_Rs.Fields.Item("ServiceDueDate").Value.ToString().Trim() != "")
                                _Grid.DataTable.SetValue("ServiceDueDate", lastRowIndex, Utilities.ToDate(_Rs.Fields.Item("ServiceDueDate").Value.ToString().Trim()));
                            else
                                _Grid.DataTable.SetValue("ServiceDueDate", lastRowIndex, _Rs.Fields.Item("ServiceDueDate").Value.ToString().Trim());
                            if (_Rs.Fields.Item("ServiceDoneDate").Value.ToString().Trim() != "")
                                _Grid.DataTable.SetValue("ServiceDoneDate", lastRowIndex, Utilities.ToDate(_Rs.Fields.Item("ServiceDoneDate").Value.ToString().Trim()));
                            else
                                _Grid.DataTable.SetValue("ServiceDoneDate", lastRowIndex, _Rs.Fields.Item("ServiceDoneDate").Value.ToString().Trim());
                            //_Grid.DataTable.SetValue("TotalService", lastRowIndex, _Rs.Fields.Item("TotalService").Value.ToString().Trim()));
                            _Grid.DataTable.SetValue("ServiceType", lastRowIndex, _Rs.Fields.Item("ServiceType").Value.ToString().Trim());
                            _Grid.DataTable.SetValue("Mileage", lastRowIndex, 0);
                            // _Grid.DataTable.SetValue("Month",   lastRowIndex, _Rs.Fields.Item("ItemCode").Value.ToString().Trim());
                            if (oDeliveryDate != "")
                            {
                                _Grid.DataTable.SetValue("DeliveryDate", lastRowIndex, Utilities.ToDate(oDeliveryDate.Trim()));
                                _Grid.DataTable.SetValue("SalesDeliveryDate", lastRowIndex, Utilities.ToDate(oDeliveryDate.Trim()));
                            }
                            else
                            {
                                _Grid.DataTable.SetValue("DeliveryDate", lastRowIndex, oDeliveryDate.Trim());
                                _Grid.DataTable.SetValue("SalesDeliveryDate", lastRowIndex, oDeliveryDate.Trim());
                            }
                            _Grid.DataTable.SetValue("Status", lastRowIndex, _Rs.Fields.Item("Status").Value.ToString().Trim());
                        }
                        _Rs.MoveNext();
                    }

                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);               
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region SAVE DATA
        private bool SaveData()
        {
            try
            {
                _Form.Freeze(true);
                UDTData = Utilities.Application.Company.UserTables.Item("SRVC_SCHDUL");              
                int j = 0;
                string oItemCode=((SAPbouiCOM.EditText)this._Form.Items.Item("45").Specific).Value;
                string oCustumerCode=((SAPbouiCOM.EditText)this._Form.Items.Item("48").Specific).Value;
                string  oSerialNo = ((SAPbouiCOM.EditText)this._Form.Items.Item("44").Specific).Value;

                for (int i = 0; i <= _Grid.Rows.Count - 1; i++)
                {
                    if (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceSerialNo")).GetText(i).ToString() != "" && ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDueDate")).GetText(i).ToString() != "")
                    {
                        string oStatus = _Grid.DataTable.Columns.Item("Status").Cells.Item(i).Value.ToString().Trim();
                        string _Code = Utilities.getMaxColumnValue("@SRVC_SCHDUL", "Code");
                        UDTData.Code = _Code.ToString();
                        UDTData.Name = _Code.ToString();
                        UDTData.UserFields.Fields.Item("U_ItemCode").Value = oItemCode;
                        UDTData.UserFields.Fields.Item("U_CustCode").Value = oCustumerCode;
                        UDTData.UserFields.Fields.Item("U_SerialNo").Value = oSerialNo;
                        string oDueDt = ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDueDate")).GetText(i).ToString();
                        //UDTData.UserFields.Fields.Item("U_SrvcCode").Value = _ServiceCode;
                        UDTData.UserFields.Fields.Item("U_SNo").Value = _Grid.DataTable.Columns.Item("ServiceSerialNo").Cells.Item(i).Value.ToString();
                        if (string.IsNullOrEmpty(oDueDt.Trim()))
                            UDTData.UserFields.Fields.Item("U_SrvDuDt").Value = oDueDt.Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_SrvDuDt").Value = Utilities.ToDate((oDueDt.Trim()));

                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDoneDate")).GetText(i).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_SrvDnDt").Value = (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDoneDate")).GetText(i)).ToString().Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_SrvDnDt").Value = Utilities.ToDate(((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDoneDate")).GetText(i)).ToString().Trim();
                        UDTData.UserFields.Fields.Item("U_SrvcType").Value = _Grid.DataTable.Columns.Item("ServiceType").Cells.Item(i).Value.ToString();
                        UDTData.UserFields.Fields.Item("U_Milage").Value = _Grid.DataTable.Columns.Item("Mileage").Cells.Item(i).Value.ToString();
                        if (string.IsNullOrEmpty((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("DeliveryDate")).GetText(i)).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_DlvryDt").Value = (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("DeliveryDate")).GetText(i)).ToString().Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_DlvryDt").Value = Utilities.ToDate((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("DeliveryDate")).GetText(i)).ToString().Trim());
                        if (string.IsNullOrEmpty((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("SalesDeliveryDate")).GetText(i)).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_SrvcDlvDt").Value = (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("SalesDeliveryDate")).GetText(i)).ToString().Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_SrvcDlvDt").Value = Utilities.ToDate((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("SalesDeliveryDate")).GetText(i)).ToString().Trim());
                        UDTData.UserFields.Fields.Item("U_Status").Value = _Grid.DataTable.Columns.Item("Status").Cells.Item(i).Value.ToString().Trim();

                        j = UDTData.Add();
                        if (j != 0)
                        {
                            Utilities.ShowErrorMessage(Utilities.Application.Company.GetLastErrorDescription());
                            return false;
                        }
                    }
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

        #region UPDATE DATA
        private bool UpdateData()
        {
            try
            {
                _Form.Freeze(true);
                string oItemCode = ((SAPbouiCOM.EditText)this._Form.Items.Item("45").Specific).Value;
                string oCustumerCode = ((SAPbouiCOM.EditText)this._Form.Items.Item("48").Specific).Value;
                string oSerialNo = ((SAPbouiCOM.EditText)this._Form.Items.Item("44").Specific).Value;
               
                UDTData = Utilities.Application.Company.UserTables.Item("SRVC_SCHDUL");
                int j = 0;

                for (int i = 0; i < _Grid.Rows.Count-1; i++)
                {
                    string oStatus = _Grid.DataTable.Columns.Item("Status").Cells.Item(i).Value.ToString().Trim();
                    if (UDTData.GetByKey(_Grid.DataTable.Columns.Item("Code").Cells.Item(i).Value.ToString()))
                    {
                        //string _Code = Utilities.getMaxColumnValue("@SRVC_SCHDUL", "Code");
                        string test = _Grid.DataTable.Columns.Item(0).Cells.Item(i).Value.ToString();
                        string a = _Grid.DataTable.Columns.Item("ServiceSerialNo").Cells.Item(i).Value.ToString();
                        string value1 = (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDoneDate")).GetText(i)).ToString();
                      
                        //UDTData.UserFields.Fields.Item("U_SrvcCode").Value = _ServiceCode;
                        UDTData.UserFields.Fields.Item("U_SNo").Value = _Grid.DataTable.Columns.Item("ServiceSerialNo").Cells.Item(i).Value.ToString();
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDueDate")).GetText(i).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_SrvDuDt").Value = ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDueDate")).GetText(i).ToString().Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_SrvDuDt").Value = Utilities.ToDate((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDueDate")).GetText(i).ToString().Trim()));
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDoneDate")).GetText(i).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_SrvDnDt").Value = (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDoneDate")).GetText(i)).ToString().Trim();
                        else
                        UDTData.UserFields.Fields.Item("U_SrvDnDt").Value = Utilities.ToDate(((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDoneDate")).GetText(i)).ToString().Trim();
                        
                        UDTData.UserFields.Fields.Item("U_SrvcType").Value = _Grid.DataTable.Columns.Item("ServiceType").Cells.Item(i).Value.ToString();
                        UDTData.UserFields.Fields.Item("U_Milage").Value = _Grid.DataTable.Columns.Item("Mileage").Cells.Item(i).Value.ToString();
                        if (string.IsNullOrEmpty((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("DeliveryDate")).GetText(i)).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_DlvryDt").Value = (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("DeliveryDate")).GetText(i)).ToString().Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_DlvryDt").Value = Utilities.ToDate((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("DeliveryDate")).GetText(i)).ToString().Trim());
                        if (string.IsNullOrEmpty((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("SalesDeliveryDate")).GetText(i)).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_SrvcDlvDt").Value = (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("SalesDeliveryDate")).GetText(i)).ToString().Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_SrvcDlvDt").Value = Utilities.ToDate((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("SalesDeliveryDate")).GetText(i)).ToString().Trim());
                        UDTData.UserFields.Fields.Item("U_Status").Value = _Grid.DataTable.Columns.Item("Status").Cells.Item(i).Value.ToString().Trim();

                       
                        j = UDTData.Update();
                        if (j != 0)
                        {
                            Utilities.ShowErrorMessage(Utilities.Application.Company.GetLastErrorDescription());
                            
                            return false;
                        }
                    }
                    else
                    {
                        string _Code = Utilities.getMaxColumnValue("@SRVC_SCHDUL", "Code");
                        UDTData.Code = _Code.ToString();
                        UDTData.Name = _Code.ToString();
                        UDTData.UserFields.Fields.Item("U_ItemCode").Value = oItemCode;
                        UDTData.UserFields.Fields.Item("U_CustCode").Value = oCustumerCode;
                        UDTData.UserFields.Fields.Item("U_SerialNo").Value = oSerialNo;
                        //UDTData.UserFields.Fields.Item("U_SrvcCode").Value = _ServiceCode;
                        UDTData.UserFields.Fields.Item("U_SNo").Value = _Grid.DataTable.Columns.Item("ServiceSerialNo").Cells.Item(i).Value.ToString();
                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDueDate")).GetText(i).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_SrvDuDt").Value = ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDueDate")).GetText(i).ToString().Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_SrvDuDt").Value = Utilities.ToDate((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDueDate")).GetText(i).ToString().Trim()));

                        if (string.IsNullOrEmpty(((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDoneDate")).GetText(i).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_SrvDnDt").Value = ((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDoneDate")).GetText(i).ToString().Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_SrvDnDt").Value = Utilities.ToDate((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceDoneDate")).GetText(i)).ToString().Trim());
                        UDTData.UserFields.Fields.Item("U_SrvcType").Value = _Grid.DataTable.Columns.Item("ServiceType").Cells.Item(i).Value.ToString();
                        UDTData.UserFields.Fields.Item("U_Milage").Value = _Grid.DataTable.Columns.Item("Mileage").Cells.Item(i).Value.ToString();
                        if (string.IsNullOrEmpty((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("DeliveryDate")).GetText(i)).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_DlvryDt").Value = (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("DeliveryDate")).GetText(i)).ToString().Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_DlvryDt").Value = Utilities.ToDate((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("DeliveryDate")).GetText(i)).ToString().Trim());
                        if (string.IsNullOrEmpty((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("SalesDeliveryDate")).GetText(i)).ToString().Trim()))
                            UDTData.UserFields.Fields.Item("U_SrvcDlvDt").Value = (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("SalesDeliveryDate")).GetText(i)).ToString().Trim();
                        else
                            UDTData.UserFields.Fields.Item("U_SrvcDlvDt").Value = Utilities.ToDate((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("SalesDeliveryDate")).GetText(i)).ToString().Trim());
                        UDTData.UserFields.Fields.Item("U_Status").Value = _Grid.DataTable.Columns.Item("Status").Cells.Item(i).Value.ToString().Trim();

                        j = UDTData.Add();
                        if (j != 0)
                        {
                            Utilities.ShowErrorMessage(Utilities.Application.Company.GetLastErrorDescription());
                            return false;
                        }
                    }

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

        #region BROWSE DATA FROM TABLE
        private bool BrowseDataFromTable()
        {
            try
            {
                _Form.Freeze(true);
                string oItemCode = ((SAPbouiCOM.EditText)this._Form.Items.Item("45").Specific).Value;
                string oCustumerCode = ((SAPbouiCOM.EditText)this._Form.Items.Item("48").Specific).Value;
                string oSerialNo = ((SAPbouiCOM.EditText)this._Form.Items.Item("44").Specific).Value;
             
                string strSQL = "Select \"Code\" as \"Code\",\"U_ItemCode\" as \"ItemCode\",\"U_CustCode\" as \"CustomerCode\",\"U_SerialNo\" as \"SerialNo\",\"U_SNo\" as \"ServiceSerialNo\",\"U_SrvDuDt\" as \"ServiceDueDate\",\"U_SrvDnDt\" as \"ServiceDoneDate\", " +
                                " \"U_SrvcType\" as \"ServiceType\",\"U_Milage\" as \"Mileage\",\"U_DlvryDt\" as \"DeliveryDate\",\"U_SrvcDlvDt\" as \"SalesDeliveryDate\" ,\"U_Status\" as \"Status\" " +
                                "from \"@SRVC_SCHDUL\" where \"U_ItemCode\"='" + oItemCode.Trim() + "' and \"U_SerialNo\"='" + oSerialNo.Trim() + "' and \"U_CustCode\"='" + oCustumerCode.Trim() + "' Order by \"U_SNo\"";

                _GridDataTable.ExecuteQuery(strSQL);
                _Grid.DataTable = _GridDataTable;
                
                if (_GridDataTable.Rows.Count > 0)
                {
                    _Form.Items.Item("BtnFill").Enabled = false;
                    _NewItem.Enabled = false;
                    _Grid.DataTable = _GridDataTable;
                    _RowHeaders = _Grid.RowHeaders;
                    _RowHeaders.TitleObject.Caption = "#";
                    for (int i = 0; i < _Grid.Rows.Count; i++)
                    {
                        _RowHeaders.SetText(i, (i + 1).ToString());
                        //_Grid.DataTable.SetValue("Status", i,"Pending");
                    }
                    _Grid.AutoResizeColumns();

                    _Grid.Columns.Item("Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;                  
                    ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Status")).ValidValues.Add("Pending", "Pending");
                    ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Status")).ValidValues.Add("Cancelled", "Cancelled");
                    ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Status")).ValidValues.Add("Completed", "Completed");

                    _Grid.Columns.Item("ServiceType").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                    ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("ServiceType")).ValidValues.Add("Free", "Free");
                    ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("ServiceType")).ValidValues.Add("Paid", "Paid");
                    ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("ServiceType")).ValidValues.Add("Scheme", "Scheme");
                    this._Grid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                    _Grid.Columns.Item("Code").Width = 0;
                    _Grid.Columns.Item("ItemCode").Width = 0;
                    _Grid.Columns.Item("SerialNo").Width =0;
                    _Grid.Columns.Item("CustomerCode").Width = 0;
                    _Grid.Columns.Item("ServiceSerialNo").TitleObject.Caption = "Service Serial No";
                    _Grid.Columns.Item("ServiceDueDate").TitleObject.Caption = "Service Due Date";
                    _Grid.Columns.Item("ServiceDoneDate").TitleObject.Caption = "Service Done Date";
                    _Grid.Columns.Item("ServiceType").TitleObject.Caption = "Service Type";
                    _Grid.Columns.Item("Mileage").TitleObject.Caption = "Mileage";
                    //_Grid.Columns.Item("Month").TitleObject.Caption = "Month";
                    _Grid.Columns.Item("DeliveryDate").TitleObject.Caption = "Delivery Date";
                    _Grid.Columns.Item("SalesDeliveryDate").TitleObject.Caption = "Sales Delivery Date";
                    _Grid.Columns.Item("Status").TitleObject.Caption = "Status";
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

        #region CHECK STATUS

        private bool CheckStatus(int oSelectedRow)
        {
            try
            {
                Form.Freeze(true);
                string oItemCode = ((SAPbouiCOM.EditText)this.Form.Items.Item("45").Specific).Value;
                string oCustumerCode = ((SAPbouiCOM.EditText)this.Form.Items.Item("48").Specific).Value;
                string oSerialNo = ((SAPbouiCOM.EditText)this.Form.Items.Item("44").Specific).Value;
                if ((((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceSerialNo")).GetText(oSelectedRow)).ToString().Trim() != "")
                {
                    string oSQL = "Select \"U_Status\" from \"@SRVC_SCHDUL\" where \"U_Status\" in ('Cancelled','Completed') and " +
                                 "\"U_SNo\"='" + (((SAPbouiCOM.EditTextColumn)_Grid.Columns.Item("ServiceSerialNo")).GetText(oSelectedRow)).ToString().Trim() + "' " +
                                 " and \"U_ItemCode\"='" + oItemCode.Trim() + "' and \"U_SerialNo\"='" + oSerialNo.Trim() + "' " +
                                 "and \"U_CustCode\"='" + oCustumerCode.Trim() + "'";
                    Utilities.ExecuteSQL(ref _Rs, oSQL);
                    if (_Rs.RecordCount > 0)
                    {
                        return true;
                    }
                    else
                        return false;

                }
                return true;
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                return false;
            }
            finally
            {
                Form.Freeze(false);
                GC.Collect();
            }
        }
        
        #endregion

        #region CREATE BP
        //Creating BP from DI API
        //After BP creation customer equipment card also gets created
        //After that service schedule tab information  
        private void CreateBP(string oBPCode,string oEnginSerNo)
        {
            try
            {
                Form.Freeze(true);
                string oSQL, oSerialNo = string.Empty, oCustumerCode = string.Empty, oItemCode = string.Empty;
                
                SAPbobsCOM.BusinessPartners oBPMaster = null;
                SAPbobsCOM.CustomerEquipmentCards oCustEquipCard = null;       
                oBPMaster = (SAPbobsCOM.BusinessPartners)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);             
                oBPMaster.CardCode = oBPCode.Trim();
                oBPMaster.CardType = BoCardTypes.cCustomer;
                 _rtnVal = oBPMaster.Add();
                 if (_rtnVal != 0)
                 {
                     Utilities.ShowErrorMessage(Utilities.Application.Company.GetLastErrorDescription());
                     Form.Items.Item("CreatBP").Enabled = true ;
                 }
                 else
                 #region CUSTOMER EQUIPMENT CARD
                 {
                     oSQL = "Select \"insID\",\"manufSN\",\"customer\",\"itemCode\" from LXM_VIRTUAL.OINS where \"customer\"='" + oBPCode.Trim() + "' and \"internalSN\"='" + oEnginSerNo.Trim() + "' ";
                     Utilities.ExecuteSQL(ref _Rs, oSQL);
                     if (_Rs.RecordCount > 0)
                     {
                         oCustEquipCard = (SAPbobsCOM.CustomerEquipmentCards)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards);
                         oCustEquipCard.InternalSerialNum = _Rs.Fields.Item(0).Value.ToString().Trim();
                         oCustEquipCard.ManufacturerSerialNum = _Rs.Fields.Item(1).Value.ToString().Trim();
                         oCustEquipCard.CustomerCode = _Rs.Fields.Item(2).Value.ToString().Trim();
                         oCustEquipCard.ItemCode = _Rs.Fields.Item(3).Value.ToString().Trim();
                         oSerialNo = _Rs.Fields.Item(0).Value.ToString().Trim();
                         oCustumerCode = _Rs.Fields.Item(2).Value.ToString().Trim();
                         oItemCode = _Rs.Fields.Item(3).Value.ToString().Trim();

                         _rtnVal = oCustEquipCard.Add();
                         if (_rtnVal != 0)
                         {
                             Utilities.ShowErrorMessage(Utilities.Application.Company.GetLastErrorDescription());
                             Form.Items.Item("CreatBP").Enabled = true;
                         }
                 #endregion
                         else
                         {                
                          #region SERVICE SCHEDULE
                            SAPbobsCOM.UserTable oTable = Utilities.Application.Company.UserTables.Item("SRVC_SCHDUL");
                            oSQL = "Select \"Code\" as \"Code\",\"U_ItemCode\" as \"ItemCode\",\"U_CustCode\" as \"CustomerCode\",\"U_SerialNo\" as \"SerialNo\",\"U_SNo\" as \"ServiceSerialNo\",\"U_SrvDuDt\" as \"ServiceDueDate\",\"U_SrvDnDt\" as \"ServiceDoneDate\", " +
                                            " \"U_SrvcType\" as \"ServiceType\",\"U_Milage\" as \"Mileage\",\"U_DlvryDt\" as \"DeliveryDate\",\"U_SrvcDlvDt\" as \"SalesDeliveryDate\" ,\"U_Status\" as \"Status\" " +
                                            "from LXM_VIRTUAL.\"@SRVC_SCHDUL\" where \"U_ItemCode\"='" + oItemCode.Trim() + "' and \"U_SerialNo\"='" + oEnginSerNo.Trim() + "' and \"U_CustCode\"='" + oCustumerCode.Trim() + "'";

                            //oSQL = "Select \"Code\" as \"Code\",\"U_ItemCode\" as \"ItemCode\",\"U_CustCode\" as \"CustomerCode\",\"U_SerialNo\" as \"SerialNo\",\"U_SNo\" as \"ServiceSerialNo\",\"U_SrvDuDt\" as \"ServiceDueDate\",\"U_SrvDnDt\" as \"ServiceDoneDate\", " +
                            //            " \"U_SrvcType\" as \"ServiceType\",\"U_Milage\" as \"Mileage\",\"U_DlvryDt\" as \"DeliveryDate\",\"U_SrvcDlvDt\" as \"SalesDeliveryDate\" ,\"U_Status\" as \"Status\" " +
                            //            "from SBODEMOIN.\"@SRVC_SCHDUL\" where \"U_ItemCode\"='S10000' and \"U_SerialNo\"='S1-000185' and \"U_CustCode\"='C60000'";

                            Utilities.ExecuteSQL(ref _Rs, oSQL);
                            if (_Rs.RecordCount > 0)
                            {
                                while (!_Rs.EoF)
                                {
                                    oSQL = "Select Max(\"Code\")+1 as \"Code\" from \"@SRVC_SCHDUL\"";
                                    Utilities.ExecuteSQL(ref _RS, oSQL);
                                    if (_RS.RecordCount > 0)
                                    {

                                        oTable.Code = _RS.Fields.Item("Code").Value.ToString().Trim();
                                        oTable.Name = _RS.Fields.Item("Code").Value.ToString().Trim();
                                    }
                                    oTable.UserFields.Fields.Item("U_ItemCode").Value = _Rs.Fields.Item("ItemCode").Value.ToString().Trim();
                                    oTable.UserFields.Fields.Item("U_CustCode").Value = _Rs.Fields.Item("CustomerCode").Value.ToString().Trim();
                                    oTable.UserFields.Fields.Item("U_SerialNo").Value = _Rs.Fields.Item("SerialNo").Value.ToString().Trim();
                                    oTable.UserFields.Fields.Item("U_SNo").Value = _Rs.Fields.Item("ServiceSerialNo").Value.ToString().Trim();
                                    oTable.UserFields.Fields.Item("U_SrvDuDt").Value = _Rs.Fields.Item("ServiceDueDate").Value.ToString().Trim();
                                    oTable.UserFields.Fields.Item("U_SrvDnDt").Value = _Rs.Fields.Item("ServiceDoneDate").Value.ToString().Trim();
                                    oTable.UserFields.Fields.Item("U_SrvcType").Value = _Rs.Fields.Item("ServiceType").Value.ToString().Trim();
                                    oTable.UserFields.Fields.Item("U_Milage").Value = _Rs.Fields.Item("Mileage").Value.ToString().Trim();
                                    oTable.UserFields.Fields.Item("U_DlvryDt").Value = _Rs.Fields.Item("DeliveryDate").Value.ToString().Trim();
                                    oTable.UserFields.Fields.Item("U_SrvcDlvDt").Value = _Rs.Fields.Item("SalesDeliveryDate").Value.ToString().Trim();
                                    oTable.UserFields.Fields.Item("U_Status").Value = _Rs.Fields.Item("Status").Value.ToString().Trim();

                                    _rtnVal = oTable.Add();
                                    if (_rtnVal != 0)
                                    {
                                        Utilities.ShowErrorMessage(Utilities.Application.Company.GetLastErrorDescription());
                                        Form.Items.Item("CreatBP").Enabled = true;
                                    }
                                    _Rs.MoveNext();
                                }
                            }
                            Form.Items.Item("CreatBP").Enabled = false;
                            #endregion
                          }
                         }
               }
                
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                Form.Freeze(false );
                GC.Collect();
            }
        }
        #endregion

        #region ADD CONTROL TO CUSTOM FOLDER
        private void AddControlToCustomFolder()
        {
            _Form.Freeze(true);
            try
            {
                //Button
                _ExistItem = _Form.Items.Item("2");
                _NewItem1 = _Form.Items.Add("CreatBP", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem1.Left = _ExistItem.Left + _ExistItem.Width + 7;
                _NewItem1.Width = _ExistItem.Width;
                _NewItem1.Height = _ExistItem.Height;
                _NewItem1.FromPane = 0;
                _NewItem1.ToPane = 0;
                _NewItem1.Top = _ExistItem.Top;
                _Button1 = (SAPbouiCOM.Button)_NewItem1.Specific;
                _Button1.Caption = "Create BP";
                _NewItem1.Enabled = false;

                _ExistItem = _Form.Items.Item("CreatBP");
                _NewItem1 = _Form.Items.Add("PrntSmtCrd", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem1.Left = _ExistItem.Left + _ExistItem.Width + 7;
                _NewItem1.Width = _ExistItem.Width + 40;
                _NewItem1.Height = _ExistItem.Height;
                _NewItem1.FromPane = 0;
                _NewItem1.ToPane = 0;
                _NewItem1.Top = _ExistItem.Top;
                _Button1 = (SAPbouiCOM.Button)_NewItem1.Specific;
                _Button1.Caption = "Print SmartCard";
              //  _NewItem1.Enabled = false;

                _ExistItem = _Form.Items.Item("100");
                _NewItem = _Form.Items.Add("GrdServSch", SAPbouiCOM.BoFormItemTypes.it_GRID);
                _NewItem.FromPane = 7;
                _NewItem.ToPane = 7;
                _NewItem.Left = 10;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.Top = _ExistItem.Top;
                _NewItem.Height = _ExistItem.Height + 20;
                _Grid = (SAPbouiCOM.Grid)_NewItem.Specific;
                _Form.DataSources.DataTables.Add("dtServSch");
                _Grid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                _GridDataTable = _Form.DataSources.DataTables.Item("dtServSch");
                _ExistItem = _Form.Items.Item("GrdServSch");

                //Button
                _ExistItem = _Form.Items.Item("GrdServSch");
                _NewItem = _Form.Items.Add("BtnFill", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem.Left = 9;
                _NewItem.Width = 65;
                _NewItem.FromPane = 7;
                _NewItem.ToPane = 7;
                _NewItem.Top = _ExistItem.Top + _ExistItem.Height;
                _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                _Button.Caption = "Fill Schedule";
                _NewItem.Enabled = true;

                //Button
                _ExistItem = _Form.Items.Item("BtnFill");
                _NewItem = _Form.Items.Add("BtnOk", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 10;
                _NewItem.Width = _ExistItem.Width;
                _NewItem.FromPane = 7;
                _NewItem.ToPane = 7;
                _NewItem.Top = _ExistItem.Top;
                _NewItem.Height = _ExistItem.Height;
                _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                _Button.Caption = "Update";
                _NewItem.Enabled = false;
              
                string strSQL = "SELECT "
                                 + @" CAST('' as varchar(50)) as ""Code""" + " , "
                                 + @" CAST('' as varchar(50)) as ""ItemCode""" + " , "
                                 + @" CAST('' as varchar(50)) as ""CustomerCode""" + " , "
                                 + @" CAST('' as varchar(50)) as ""SerialNo""" + " , "
                                 + @" CAST('' as varchar(50)) as ""ServiceSerialNo""" + " , "
                                 + @" CAST('' as varchar(50)) as ""ServiceDueDate""" + " , "
                                 + @" CAST('' as varchar(50)) as ""ServiceDoneDate""" + " , "
                                 + @" CAST('' as varchar(50)) as ""ServiceType""" + " , "
                                 + @" CAST('' as varchar(50)) as ""Mileage""" + " , "
                                 + @" CAST('' as varchar(50)) as ""Month""" + " , "
                                 + @" CAST('' as varchar(50)) as ""DeliveryDate""" + " , "
                                 + @" CAST('' as varchar(50)) as ""SalesDeliveryDate""" + " , "
                                 + @" 'Pending' as ""Status"""
                                 + " from "
                                 + @"""DUMMY""";

                _GridDataTable.ExecuteQuery(strSQL);
                _Grid.DataTable = _GridDataTable;
                _Grid.Columns.Item("Code").Width = 0;
                _Grid.Columns.Item("ItemCode").Width = 0;
                _Grid.Columns.Item("SerialNo").Width = 0;
                _Grid.Columns.Item("CustomerCode").Width = 0;

                _Grid.Columns.Item("ServiceType").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("ServiceType")).ValidValues.Add("", "");
                ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("ServiceType")).ValidValues.Add("Free", "Free");
                ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("ServiceType")).ValidValues.Add("Paid", "Paid");
                ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("ServiceType")).ValidValues.Add("Scheme", "Scheme");
               
                _Grid.Columns.Item("Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Status")).ValidValues.Add("Pending", "Pending");
                ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Status")).ValidValues.Add("Cancelled", "Cancelled");
                ((SAPbouiCOM.ComboBoxColumn)_Grid.Columns.Item("Status")).ValidValues.Add("Completed", "Completed");
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally { _Form.Freeze(false); }
        }
        #endregion

        #region BUTTON ENABLE DISABLE
        public void ButtonEnableDisable(string oMode)
        {
            try
            {
                if (oMode == "ComBoSelect")
                {
                    _Form.Items.Item("BtnFill").Enabled = false;
                    _Form.Items.Item("BtnOk").Enabled = false;
                    _Form.Items.Item("CreateBP").Enabled = false;
                }
                if (oMode == "OKMode")
                {
                    _Form.Items.Item("BtnFill").Enabled = false;
                    _Form.Items.Item("BtnOk").Enabled = false;
                    _Form.Items.Item("CreateBP").Enabled = false;
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion

        #endregion

    }
}
