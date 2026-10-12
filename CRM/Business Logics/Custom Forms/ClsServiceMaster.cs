using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace CRM
{
    class ClsServiceMaster : Base
    {
        #region Variable Declaration

        string _Sql = "";
        SAPbobsCOM.Recordset _RS;
        SAPbouiCOM.EditText _Edit;
        SAPbouiCOM.ComboBox _Combo;
        private SAPbouiCOM.Grid _Grid = null;
        SAPbouiCOM.DataTable _DataTable;
        ChooseFromListEvent _SysCFLEvent;
        Conditions _Cons;
        Condition _Con;
        private SAPbouiCOM.CommonSetting _CommonSetting;
        private SAPbouiCOM.RowHeaders _RowHeaders;
        SAPbouiCOM.Matrix _Matrix = null;
        private SAPbouiCOM.DBDataSource _DataMaster = null; //Declaring Recordset
        private SAPbouiCOM.DBDataSource _DataChild = null; //Declaring _DataMaster
        bool _isDeleteRow = false;

        #endregion

        #region CONSTRUCTOR

        public ClsServiceMaster()
            : base()
        {
        }

        #endregion

        #region FORM EVENT

        #endregion

        #region MENU EVENT
        public override void Menu_Event(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            if (pVal.BeforeAction)
            {
                switch (pVal.MenuUID)
                {
                    case Constants.System_Menus.mnu_DELETE_ROW:
                        try
                        {
                            _Form.Freeze(true);
                            if (_Matrix.VisualRowCount > 1)
                                _isDeleteRow = true;
                            else
                                Form.EnableMenu("1293", false);
                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        finally { _Form.Freeze(false); }
                        break;

                }
            }
            else
            {
                switch (pVal.MenuUID)
                {
                    case Constants.System_Menus.mnu_ADD:
                        try
                        {
                            InitializeForm();
                            // _Matrix.AutoResizeColumns();
                            //_CommonSetting = _Matrix.CommonSetting;
                            // this._Form.Items.Item("3").Enabled = true;
                            //((EditText)this._Form.Items.Item("3").Specific).Active = true;
                        }
                        catch (Exception) { }
                        break;
                    case Constants.System_Menus.mnu_DELETE_ROW:

                        try
                        {
                            _Form.Freeze(true);
                            if (_isDeleteRow = true && (_Matrix.VisualRowCount > 1) && ((SAPbouiCOM.EditText)_Matrix.Columns.Item("ColMnth").Cells.Item(_Matrix.VisualRowCount).Specific).Value.ToString() != "")
                            {
                                try
                                {
                                    int ChildDbOrgSize = _DataChild.Size;
                                    _Matrix.FlushToDataSource();
                                    _DataChild.Offset = 0;
                                    _Form.Freeze(true);

                                    for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                                    {
                                        _DataChild.SetValue("LineId", _DataChild.Offset, i.ToString());

                                        if (_DataChild.Size > i)
                                            _DataChild.Offset = i;
                                    }

                                    while (_Matrix.VisualRowCount < ChildDbOrgSize)
                                    {
                                        _DataChild.RemoveRecord(ChildDbOrgSize - 1);
                                        ChildDbOrgSize--;
                                    }

                                    _Matrix.LoadFromDataSource();
                                    _isDeleteRow = false;
                                }
                                finally
                                {
                                    _Form.Freeze(false);
                                }
                            }
                        }
                        catch (Exception) { }
                        finally { _Form.Freeze(false); }
                        break;
                }
            }
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
                    #region Right click
                    case BoEventTypes.et_RIGHT_CLICK:

                        if (_Matrix.VisualRowCount > 1)
                            Form.EnableMenu("1293", true);
                        else
                            Form.EnableMenu("1293", false);

                        break;
                    #endregion

                    #region ITEM PRESSED
                    case BoEventTypes.et_ITEM_PRESSED:
                        try
                        {
                            _Form.Freeze(true);
                            if (pVal.ItemUID == "1" && (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                if (this.validation())
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
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
            else
            {
                switch (pVal.EventType)
                {
                    #region Right click
                    case BoEventTypes.et_RIGHT_CLICK:

                        if (_Matrix.VisualRowCount > 1)
                            Form.EnableMenu("1293", true);
                        else
                            Form.EnableMenu("1293", false);

                        break;
                    #endregion

                    #region ITEM PRESSED
                    case BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && pVal.Action_Success)
                        {
                            try
                            {
                                _Form.Freeze(true);
                                if (_Form.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    InitializeForm();
                                }
                            }
                            finally
                            {
                                _Form.Freeze(false);
                            }
                        }

                        break;
                    #endregion

                    #region  Key Down
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        if (pVal.ItemUID == "matrixSrvc" && (pVal.ColUID == "ColMilag" || pVal.ColUID == "ColMnth") && pVal.CharPressed == Constants.Keys.TAB)
                        {
                            _Form.Freeze(true);
                            try
                            {
                                if (//((SAPbouiCOM.EditText)_Matrix.Columns.Item("ColMilag").Cells.Item(pVal.Row).Specific).Value.ToString() != "" && pVal.ColUID == "ColMilag" &&
                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("ColMnth").Cells.Item(pVal.Row).Specific).Value.ToString() != "" && pVal.ColUID == "ColMnth"
                                    //&& ((SAPbouiCOM.EditText)_Matrix.Columns.Item("ColType").Cells.Item(pVal.Row).Specific).Value.ToString() != "" && pVal.ColUID == "ColType"
                                    )
                                {
                                    _Matrix.FlushToDataSource();
                                    if (pVal.Row == _DataChild.Size)
                                    {
                                        _Matrix.AddRow(1, -1);
                                        _Matrix.FlushToDataSource();
                                        SetNewLine(_Matrix.VisualRowCount);
                                        _Matrix.LoadFromDataSource();
                                        _Matrix.Columns.Item("V_0").Cells.Item(_Matrix.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    }
                                    else if (_DataChild.Size > pVal.Row)
                                    {
                                        if (_DataChild.GetValue("LineId", pVal.Row) == "")
                                        {
                                            int val = pVal.Row + 1;
                                            _DataChild.SetValue("LineId", pVal.Row, val.ToString());
                                            _Matrix.FlushToDataSource();
                                            _Matrix.LoadFromDataSource();
                                        }
                                    }
                                }
                            }
                            finally { _Form.Freeze(false); }
                        }
                        break;
                    #endregion
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
                            if (_DataChild.GetValue("U_SNo", _DataChild.Size - 1).Trim() != "" && _DataChild.GetValue("U_Months", _DataChild.Size - 1).Trim() != "")
                            {
                                _Matrix.AddRow(1, -1);
                                _Matrix.FlushToDataSource();
                                SetNewLine(_Matrix.VisualRowCount);
                            }

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

                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        break;
                    #endregion

                    #region DATA ADD
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        try
                        {
                            _Form.Freeze(true);
                            if (_Matrix.VisualRowCount == 0)
                            {
                                _Matrix.AddRow(1, -1);
                                _Matrix.FlushToDataSource();
                            }
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

            #region BEFORE ACTION
            else
            {
                switch (BusinessObjectInfo.EventType)
                {
                    #region DATA LOAD
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        try
                        {


                        }
                        catch (Exception ex)
                        {
                            Utilities.ShowErrorMessage(ex.Message);
                        }
                        break;
                    #endregion

                    #region DATA ADD
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        try
                        {
                            _Form.Freeze(true);
                            if (_DataChild.GetValue("U_SNo", _DataChild.Size - 1).Trim() == "" && _DataChild.GetValue("U_Months", _DataChild.Size - 1).Trim() == "")
                                _DataChild.RemoveRecord(_DataChild.Size - 1);
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
                            if (_DataChild.GetValue("U_SNo", _DataChild.Size - 1).Trim() == "" && _DataChild.GetValue("U_Months", _DataChild.Size - 1).Trim() == "")
                                _DataChild.RemoveRecord(_DataChild.Size - 1);
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

        public void FormDefault()
        {
            try
            {
                InitializeForm();
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                Form.Visible = true;
            }
        }

        public void InitializeForm()
        {
            try
            {
                _DataMaster = _Form.DataSources.DBDataSources.Item("@SRVC_MSTR");
                _DataChild = _Form.DataSources.DBDataSources.Item("@SRVC_DETAIL");
                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("matrixSrvc").Specific;
                _Form.DataBrowser.BrowseBy = "etCode";

                _DataMaster.SetValue("Code", _DataMaster.Offset, GetTransactionNextSeries().ToString());
                _DataChild.SetValue("Code", _DataChild.Offset, _DataMaster.GetValue("Code", _DataMaster.Offset));

                _DataChild.SetValue("LineId", _DataChild.Offset, "1");
                _Matrix.LoadFromDataSource();
                int io = _Matrix.VisualRowCount;
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }

        private void SetNewLine(int Row)
        {
            _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("matrixSrvc").Specific;
            _DataChild.Offset = Row - 1;
            _DataChild.SetValue("LineId", _DataChild.Offset, Row.ToString());
            _DataChild.SetValue("U_SNo", _DataChild.Offset, "");
            _DataChild.SetValue("U_Type", _DataChild.Offset, "");
            _DataChild.SetValue("U_Milage", _DataChild.Offset, "");
            _DataChild.SetValue("U_Months", _DataChild.Offset, "");

            _Matrix.SetLineData(Row);
            _Matrix.FlushToDataSource();
        }

        private string GetTransactionNextSeries()
        {
            int number = 0;
            string result = string.Empty;
            try
            {
                _Sql = "Select \"AutoKey\" from ONNM Where \"ObjectCode\"='SRVCMSTR' ";
                Utilities.ExecuteSQL(ref _RS, _Sql);
                int.TryParse(_RS.Fields.Item("Autokey").Value.ToString(), out number);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                int.TryParse(_RS.Fields.Item("Autokey").Value.ToString(), out number);
            }
            result = number.ToString().PadLeft(8, '0');
            return result;
        }

        private bool validation()
        {
            bool rt = false;

            try
            {
                _Form.Freeze(true);

                if (((EditText)this.Form.Items.Item("EtSrvCd").Specific).Value == "")
                {
                    Utilities.ShowErrorMessage("Service code can not be blank!");
                    ((EditText)this.Form.Items.Item("EtSrvCd").Specific).Active = true;
                    return true;
                }


                if (((EditText)this.Form.Items.Item("etSrvDtl").Specific).Value == "")
                {
                    Utilities.ShowErrorMessage("Service description can not be blank!");
                    ((EditText)this.Form.Items.Item("etSrvDtl").Specific).Active = true;
                    return true;
                }

                if (_Matrix.VisualRowCount > 0)
                {
                    _Matrix.FlushToDataSource();
                    _DataChild.Offset = 0;

                    for (int i = 0; i < _Matrix.VisualRowCount - 1; i++)
                    {
                        if (_DataChild.GetValue("U_SNo", i).Trim() == "")
                        {
                            Utilities.ShowErrorMessage("Serial number can not be blank!");
                            _Matrix.Columns.Item("ColSNo").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return true;
                        }
                        if (_DataChild.GetValue("U_Type", i).Trim() == "")
                        {
                            Utilities.ShowErrorMessage("Service type can not be blank!");
                            _Matrix.Columns.Item("ColType").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return true;
                        }
                        if (_DataChild.GetValue("U_Milage", i).Trim() == "" || Convert.ToDouble(_DataChild.GetValue("U_Milage", i).Trim()) <= 0.0)
                        {
                            Utilities.ShowErrorMessage("Milage can not be blank!");
                            _Matrix.Columns.Item("ColMilag").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return true;
                        }
                        if (_DataChild.GetValue("U_Months", i).Trim() == "")
                        {
                            Utilities.ShowErrorMessage("Months can not be blank!");
                            _Matrix.Columns.Item("ColMnth").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return true;
                        }
                    }
                }

                // reset Lineid incase any row has been deleted
                if (_isDeleteRow)
                {
                    try
                    {
                        int ChildDbOrgSize = _DataChild.Size;
                        _Matrix.FlushToDataSource();
                        _DataChild.Offset = 0;
                        _Form.Freeze(true);

                        for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                        {
                            _DataChild.SetValue("LineId", _DataChild.Offset, i.ToString());

                            if (_DataChild.Size > i)
                                _DataChild.Offset = i;
                        }

                        while (_Matrix.VisualRowCount < ChildDbOrgSize)
                        {
                            _DataChild.RemoveRecord(ChildDbOrgSize - 1);
                            ChildDbOrgSize--;
                        }

                        _Matrix.LoadFromDataSource();
                        _isDeleteRow = false;
                    }
                    finally
                    {
                        _Form.Freeze(false);
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                return true;
            }
            finally
            {
                _Form.Freeze(false);
            }
            return rt;
        }
        #endregion
    }
}
