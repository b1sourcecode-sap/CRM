using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace CRM
{
    class ClsSchemeMaster : Base
    {

        #region VARIABLES

        private SAPbouiCOM.Matrix _Matrix;
        private SAPbouiCOM.DBDataSource _DBDataSource;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource;
        private SAPbouiCOM.ComboBox _ComboBox;
        private SAPbouiCOM.EditText _EditText;
        private SAPbobsCOM.Recordset _Rs;
        private bool Flags = true;
        private int RowToDelete;
        private string _Sql;
        private string strSql;
        int _LastErrorCode = 0;
        string _LastErrorMsg = " ";
        bool _isDeleteRow = false;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsSchemeMaster()
            : base()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsSchemeMaster()
        {
        }
        #endregion

        #region Item Event
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.Before_Action)
            {
                try
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            try
                            {
                                if (pVal.ItemUID == "1")
                                {
                                    _Form.Freeze(true);
                                    _Matrix.FlushToDataSource();
                                    if (Form.Mode == BoFormMode.fm_ADD_MODE && Form.Mode == BoFormMode.fm_UPDATE_MODE)
                                    {
                                        if (Validation() == false)
                                        {
                                            oBubbleEvent = false;
                                            return;
                                        }
                                        if (_DBDataSource.GetValue("U_SchmCd", _DBDataSource.Size - 1).Trim() == "" && _DBDataSource.GetValue("U_SchmNm", _DBDataSource.Size - 1).Trim() == ""
                                            // && _DBDataSource.GetValue("U_FromDt", _DBDataSource.Size - 1).Trim() == "" && _DBDataSource.GetValue("U_ToDt", _DBDataSource.Size - 1).Trim() == ""
                                            && _DBDataSource.GetValue("U_ItemGrp", _DBDataSource.Size - 1).Trim() == "" && _DBDataSource.GetValue("U_TotSrvc", _DBDataSource.Size - 1).Trim() == "")
                                        {
                                            _DBDataSource.RemoveRecord(_DBDataSource.Size - 1);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Utilities.ShowErrorMessage(ex.Message);
                            }
                            finally { _Form.Freeze(false); }
                            break;
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

            #region After Action
            else
            {
                try
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            if (pVal.ItemUID == "1")
                            {

                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
            #endregion
        }

        #endregion

        #region EVENT :Menu Event
        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            if (pVal.BeforeAction == false)
            {
                switch (pVal.MenuUID)
                {
                    case Constants.System_Menus.mnu_DELETE_ROW:
                        try
                        {
                            _isDeleteRow = true;
                            _DBDataSource.RemoveRecord(this.RowToDelete - 1);

                            _Matrix.LoadFromDataSource();

                            // reset Lineid incase any row has been deleted
                            if (_isDeleteRow)
                            {
                                try
                                {
                                    int ChildDbOrgSize = _DBDataSource.Size;
                                    _Matrix.FlushToDataSource();
                                    _DBDataSource.Offset = 0;
                                    _Form.Freeze(true);

                                    for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                                    {
                                        _DBDataSource.SetValue("LineId", _DBDataSource.Offset, i.ToString());

                                        if (_DBDataSource.Size > i)
                                            _DBDataSource.Offset = i;
                                    }

                                    while (_Matrix.VisualRowCount < ChildDbOrgSize)
                                    {
                                        _DBDataSource.RemoveRecord(ChildDbOrgSize - 1);
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

                            _isDeleteRow = false;
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
                }
            }
        }
        #endregion

        #region FormData_Event
        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            #region Before Action
            if (BusinessObjectInfo.BeforeAction)
            {
                try
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
            #endregion

            #region After Action
            else
            {
                try
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            try
                            {
                                if (BusinessObjectInfo.ActionSuccess == true)
                                {
                                    AddNewRow();
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
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
            #endregion
        }
        #endregion

        #region RightClick_Event
        public override void RightClick_Event(ref SAPbouiCOM.ContextMenuInfo oEventInfo, ref bool oBubbleEvent)
        {
            #region Before Action
            if (oEventInfo.BeforeAction)
            {
                try
                {
                    switch (oEventInfo.EventType)
                    {

                        case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK:
                            try
                            {
                                if (oEventInfo.ItemUID == "mtMatrix")
                                {
                                    this.RowToDelete = oEventInfo.Row;
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
                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
            #endregion

            #region After Action
            else
            {
                try
                {
                    switch (oEventInfo.EventType)
                    {

                    }
                }
                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
            }
            #endregion
        }
        #endregion

        #region FormDefault Method

        public void FormDefault()
        {
            try
            {
                Form.Freeze(true);
                _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("mtMatrix").Specific;
                _DBDataSource = Form.DataSources.DBDataSources.Item("@SCHM_DETAIL");
                _MasterDBDataSource = Form.DataSources.DBDataSources.Item("@SCHM_MSTR");

                _DBDataSource.Query(null);
                _MasterDBDataSource.Query(null);

                if (_MasterDBDataSource.Size == 0)
                {
                    _MasterDBDataSource.SetValue("Code", _MasterDBDataSource.Offset, "00000001");
                    _MasterDBDataSource.SetValue("DocEntry", _MasterDBDataSource.Offset, "1");
                    _MasterDBDataSource.InsertRecord(0);
                    _DBDataSource.SetValue("Code", _MasterDBDataSource.Offset, _MasterDBDataSource.GetValue("Code", _MasterDBDataSource.Offset));

                    _DBDataSource.SetValue("LineId", _DBDataSource.Offset, "1");
                    _Matrix.LoadFromDataSource();
                    _MasterDBDataSource.Query(null);
                }
                if (dataexist() == true)
                {
                    // this.ConditionalReloadData();
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                }
                else
                {
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    _MasterDBDataSource.SetValue("Code", _MasterDBDataSource.Offset, "00000001");
                    _MasterDBDataSource.SetValue("DocEntry", _MasterDBDataSource.Offset, "1");
                    _DBDataSource.SetValue("Code", _DBDataSource.Offset, _MasterDBDataSource.GetValue("Code", _MasterDBDataSource.Offset));
                    //((SAPbouiCOM.EditText)_Matrix.Columns.Item("ColSno").Cells.Item(_Matrix.RowCount).Specific).Value = _Matrix.RowCount.ToString();
                    //_Matrix.FlushToDataSource();
                    _DBDataSource.SetValue("LineId", _DBDataSource.Offset, "1");
                    _Matrix.LoadFromDataSource();
                }

                Form.EnableMenu(Constants.System_Menus.mnu_FIND, false);
                Form.EnableMenu(Constants.System_Menus.mnu_ADD, false);
                Form.EnableMenu(Constants.System_Menus.mnu_DELETE_ROW, true);
                AddNewRow();

                _Matrix.AutoResizeColumns();
                //Utilities.SerializedMartix(ref _Form, ref _Matrix);
            }

            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                Form.Freeze(false);
                Form.Visible = true;
            }
        }
        #endregion

        #region Add New Row Method

        private void AddNewRow()
        {
            try
            {
                if (_DBDataSource != null && _Matrix != null)
                {

                    if (_DBDataSource.Size == 0 || _DBDataSource.GetValue("U_SchmNm", _DBDataSource.Size - 1).Trim() != "")
                    {
                        _DBDataSource.InsertRecord(_DBDataSource.Size);
                        _DBDataSource.SetValue("Code", _DBDataSource.Size - 1, _MasterDBDataSource.GetValue("Code", 0).Trim());
                        _DBDataSource.SetValue("LineId", _DBDataSource.Size - 1, Utilities.getMaxColumnValue("@SCHM_DETAIL", "LineId"));
                    }
                    else if (_DBDataSource.Size == 1 && _DBDataSource.GetValue("Code", _DBDataSource.Size - 1).Trim() == "")
                    {
                        _DBDataSource.SetValue("Code", _DBDataSource.Size - 1, _MasterDBDataSource.GetValue("Code", 0));
                        _DBDataSource.SetValue("LineId", _DBDataSource.Size - 1, Utilities.getMaxColumnValue("@SCHM_DETAIL", "LineId"));
                    }
                    _Matrix.LoadFromDataSource();
                    _Matrix.AutoResizeColumns();
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

        #region VALIDATION METHOD

        private bool Validation()
        {
            try
            {
                _Form.Freeze(true);
                _Matrix.FlushToDataSource();
                _Matrix.LoadFromDataSource();
                bool skipRow = false;
                for (int i = 0; i <= _DBDataSource.Size - 1; i++)
                {
                    if ((i == _DBDataSource.Size - 1 || Flags == false)
                        && _DBDataSource.GetValue("U_SchmCd", _DBDataSource.Size - 1).Trim() == ""
                        && _DBDataSource.GetValue("U_SchmNm", _DBDataSource.Size - 1).Trim() == ""
                        && _DBDataSource.GetValue("U_FromDt", _DBDataSource.Size - 1).Trim() == ""
                        && _DBDataSource.GetValue("U_ToDt", _DBDataSource.Size - 1).Trim() == ""
                        && _DBDataSource.GetValue("U_ItemGrp", _DBDataSource.Size - 1).Trim() == ""
                        && _DBDataSource.GetValue("U_TotSrvc", _DBDataSource.Size - 1).Trim() == ""
                       )
                    {
                        skipRow = true;
                    }

                    if (skipRow == false)
                    {
                        if (_DBDataSource.GetValue("U_SchmCd", i).Trim() == "")
                        {
                            Utilities.ShowErrorMessage("Scheme code can not Blank");
                            _Matrix.Columns.Item("ColSchCd").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return false;
                        }

                        if (_DBDataSource.GetValue("U_SchmNm", i).Trim() == "")
                        {
                            Utilities.ShowErrorMessage("Scheme description can not Blank");
                            _Matrix.Columns.Item("ColSchNm").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return false;
                        }
                        if (_DBDataSource.GetValue("U_FromDt", i).Trim() == "")
                        {
                            Utilities.ShowErrorMessage("From date can not Blank");
                            _Matrix.Columns.Item("ColFrmDt").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return false;
                        }

                        if (_DBDataSource.GetValue("U_ToDt", i).Trim() == "")
                        {
                            Utilities.ShowErrorMessage("To date can not Blank");
                            _Matrix.Columns.Item("ColToDt").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return false;
                        }
                        if (_DBDataSource.GetValue("U_ItemGrp", i).Trim() == "")
                        {
                            Utilities.ShowErrorMessage("Item group can not Blank");
                            _Matrix.Columns.Item("ColItmGrp").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return false;
                        }
                        if (_DBDataSource.GetValue("U_TotSrvc", i).Trim() == "")
                        {
                            Utilities.ShowErrorMessage("No. of services can not Blank");
                            _Matrix.Columns.Item("ColTotSch").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return false;
                        }

                    }
                }
                if (ItemGroupValidation() == false)
                {
                    return false;
                }

                // reset Lineid incase any row has been deleted
                if (_isDeleteRow)
                {
                    try
                    {
                        int ChildDbOrgSize = _DBDataSource.Size;
                        _Matrix.FlushToDataSource();
                        _DBDataSource.Offset = 0;
                        _Form.Freeze(true);

                        for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                        {
                            _DBDataSource.SetValue("LineId", _DBDataSource.Offset, i.ToString());

                            if (_DBDataSource.Size > i)
                                _DBDataSource.Offset = i;
                        }

                        while (_Matrix.VisualRowCount < ChildDbOrgSize)
                        {
                            _DBDataSource.RemoveRecord(ChildDbOrgSize - 1);
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

                return true;
            }
            catch
            {
                throw;
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        #endregion

        #region ItemGroupValidation

        private bool ItemGroupValidation()
        {
            try
            {
                _Matrix.FlushToDataSource();
                _Matrix.LoadFromDataSource();
                string oCurrentSchemeCode, oCurrentSchemeName, oCurrentItmGrp, oPrvSchemeCode, oPrvSchemeName, oPrvItmGrp;
                string oCurrntFrmDt, oPrvFrmDt, oCurrentToDt, oPrvToDt;
                for (int i = 0; i <= _DBDataSource.Size - 1; i++)
                {
                    oCurrentSchemeCode = _DBDataSource.GetValue("U_SchmCd", i).Trim();
                    oCurrentSchemeName = _DBDataSource.GetValue("U_SchmNm", i).Trim();
                    oCurrentItmGrp = _DBDataSource.GetValue("U_ItemGrp", i).Trim();
                    oCurrntFrmDt = _DBDataSource.GetValue("U_FromDt", i).Trim();
                    oCurrentToDt = _DBDataSource.GetValue("U_ToDt", i).Trim();
                    for (int j = 1; j <= _DBDataSource.Size - 1; j++)
                    {
                        if (i != j)
                        {
                            oPrvSchemeCode = _DBDataSource.GetValue("U_SchmCd", j).Trim();
                            oPrvSchemeName = _DBDataSource.GetValue("U_SchmNm", j).Trim();
                            oPrvItmGrp = _DBDataSource.GetValue("U_ItemGrp", j).Trim();
                            oPrvFrmDt = _DBDataSource.GetValue("U_FromDt", j).Trim();
                            oPrvToDt = _DBDataSource.GetValue("U_ToDt", j).Trim();
                            if (oCurrentSchemeCode == oPrvSchemeCode && oCurrentSchemeName == oPrvSchemeName && oCurrentItmGrp == oPrvItmGrp
                                && oCurrntFrmDt == oPrvFrmDt && oCurrentToDt == oPrvToDt)
                            {
                                Utilities.ShowErrorMessage("Scheme already exist for for same item group for same date range");
                                _Matrix.Columns.Item("ColSchCd").Cells.Item(j).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return false;
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                return false;
            }
        }

        #endregion

        #region DATAEXIST
        private bool dataexist()
        {
            try
            {
                _Form.Freeze(true);
                _Sql = "Select Count(*) from \"@SCHM_MSTR\"";
                Utilities.ExecuteSQL(ref _Rs, _Sql);
                if (_Rs.RecordCount > 0 && Convert.ToInt32(_Rs.Fields.Item(0).Value.ToString()) > 0)
                {
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally { _Form.Freeze(false); }

        }
        #endregion

    }
}
