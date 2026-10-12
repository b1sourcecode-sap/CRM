using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRM
{
    class ClsSalesOrder : Base
    {

        #region VARIABLE DECLARATION

        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        SAPbouiCOM.ButtonCombo _ButtonCombo;
        SAPbouiCOM.ComboBox _ComboBox;
        SAPbouiCOM.Button _Button;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsSalesOrder()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsSalesOrder()
        {
        }
        #endregion

        #region Form Default
        internal void FormDefault()
        {
            try
            {

                //_ComboBox = _Form.Items.Item("10000329").Specific;
                //_ComboBox.ValidValues.Add("Booking Canc. Req.", "Booking Canc. Req.");

                _ExistItem = _Form.Items.Item("1");
                _NewItem = _Form.Items.Add("BookingC", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                _NewItem.Left = _ExistItem.Left + _NewItem.Width +50;
                _NewItem.Width = _ExistItem.Width + 50;
                _NewItem.Top = _ExistItem.Top;
                _NewItem.Height = 16;
                _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                _Button.Caption = "Booking Cancellation";
             
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion

        #region Menu_Event

        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        {
            try
            {
                Form.Freeze(true);

                if (pVal.BeforeAction)
                {
                    
                    switch (pVal.MenuUID)
                    {

                        case Constants.System_Menus.mnu_NEXT:
                          //_ComboBox = _Form.Items.Item("10000329").Specific;
                          //  _ComboBox.ValidValues.Add("Booking Canc. Req.", "Booking Canc. Req.");
                            break;


                        case Constants.System_Menus.mnu_PREVIOUS:
                           //_ComboBox = _Form.Items.Item("10000329").Specific;
                            //_ComboBox.ValidValues.Add("Booking Canc. Req.", "Booking Canc. Req.");
                            break;

                        case Constants.System_Menus.mnu_FIRST:
                            //_ComboBox = _Form.Items.Item("10000329").Specific;
                            //_ComboBox.ValidValues.Add("Booking Canc. Req.", "Booking Canc. Req.");
                            break;

                        case Constants.System_Menus.mnu_LAST:
                            // _ComboBox = _Form.Items.Item("10000329").Specific;
                            //_ComboBox.ValidValues.Add("Booking Canc. Req.", "Booking Canc. Req.");
                            break;
                    }
                }
                else
                {

                    switch (pVal.MenuUID)
                    {

                        case Constants.System_Menus.mnu_NEXT:
                            //_ComboBox = _Form.Items.Item("10000329").Specific;
                            //_ComboBox.ValidValues.Add("Booking Canc. Req.", "Booking Canc. Req.");
                            break;

                        case Constants.System_Menus.mnu_PREVIOUS:
                            //_ComboBox = _Form.Items.Item("10000329").Specific;
                            //_ComboBox.ValidValues.Add("Booking Canc. Req.", "Booking Canc. Req.");
                            break;

                        case Constants.System_Menus.mnu_FIRST:
                            //_ComboBox = _Form.Items.Item("10000329").Specific;
                            //_ComboBox.ValidValues.Add("Booking Canc. Req.", "Booking Canc. Req.");
                            break;

                        case Constants.System_Menus.mnu_LAST:
                            //_ComboBox = _Form.Items.Item("10000329").Specific;
                            //_ComboBox.ValidValues.Add("Booking Canc. Req.", "Booking Canc. Req.");
                            break;

                    }
                   
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                Form.Freeze(false);
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
                    //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                    //    if (pVal.ItemUID == "10000329")
                    //    {
                    //        _ComboBox = _Form.Items.Item("10000329").Specific;
                    //        //string test = _ComboBox.ValidValues.Item(3).Value.Trim();
                    //        //string test = _ComboBox.Selected.Value.ToString().Trim();
                    //        //if (_ComboBox.Selected.Value.ToString() == "Booking Canc. Req.")
                    //        //{
                    //        //    Utilities.Application.SBO_Application.MessageBox("Menu Already Exists", 1, "Ok", "", "");
                    //        //}
                    //        _Items = _Form.Items.Item("10000329");  // This is the Item Number for Copy To
                    //        _ComboBox = (SAPbouiCOM.ComboBox)_Items.Specific;
                    //        string aa = _ComboBox.Selected.Description.ToString();
                    //    }
                    //    break;

                    #endregion

                }
            }
            #endregion

            #region After ACTION ACTION
            else if (pVal.Before_Action == false)
            {

                switch (pVal.EventType)
                {
                    //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                    //    if (pVal.ItemUID == "10000329")
                    //    {
                    //        _ComboBox = _Form.Items.Item("10000329").Specific;
                    //        if (_ComboBox.Selected.Description.ToString() == "Booking Canc. Req.")
                    //        {
                    //            this._Object = new ClsBookingCancellationRequest();
                    //            Utilities.LoadForm(ref this._Object, Constants.Forms.BookingCancellationRequest);
                    //            ((ClsBookingCancellationRequest)_Object).FormDefault(_Form.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0).ToString(), _Form.DataSources.DBDataSources.Item("ORDR").GetValue("CardCode", 0).ToString(), _Form.DataSources.DBDataSources.Item("ORDR").GetValue("DocNum", 0).ToString());
                    //        }
                    //    }
                    //    break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "BookingC")
                        {
                            //_ComboBox = _Form.Items.Item("10000329").Specific;
                            //if (_ComboBox.Selected.Description.ToString() == "Booking Canc. Req.")
                            //{
                                this._Object = new ClsBookingCancellationRequest();
                                Utilities.LoadForm(ref this._Object, Constants.Forms.BookingCancellationRequest);
                                ((ClsBookingCancellationRequest)_Object).FormDefault(_Form.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0).ToString(), _Form.DataSources.DBDataSources.Item("ORDR").GetValue("CardCode", 0).ToString(), _Form.DataSources.DBDataSources.Item("ORDR").GetValue("DocNum", 0).ToString());
                            //}
                        }
                        break;
                }
            }

            #endregion

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

                        #region et_FORM_DATA_ADD Before action
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:

                            break;
                        #endregion

                        #region et_FORM_DATA_UPDATE Before action
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                           
                            break;
                        #endregion

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
                        #region et_FORM_DATA_LOAD After Action

                        //case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        //    _ComboBox = _Form.Items.Item("10000329").Specific;
                        //    _ComboBox.ValidValues.Add("Booking Canc. Req.", "Booking Canc. Req.");
                        //    break;

                        #endregion
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

        public override void RightClick_Event(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {

            if (eventInfo.BeforeAction)
            {
                Utilities.Application.SBO_Application.Menus.Item("5896").String = "Engine/VIN Selection";
                Utilities.Application.SBO_Application.Menus.Item("5904").String = "New Follow-up";
                Utilities.Application.SBO_Application.Menus.Item("5905").String = "Related Follow-up";
                Utilities.Application.SBO_Application.Menus.Item("5909").String = "Engine/VIN Transactions Report";
                Utilities.Application.SBO_Application.Menus.Item("6028").String = "Related Enquiries";

            }
        }

    }
}
