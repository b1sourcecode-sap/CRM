using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRM
{
    class ClsDelivery : Base
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

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsDelivery()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsDelivery()
        {
        }

        #endregion

        public override void RightClick_Event(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {

            if (eventInfo.BeforeAction)
            {
                Utilities.Application.SBO_Application.Menus.Item("5904").String = "New Follow-up";
                Utilities.Application.SBO_Application.Menus.Item("5905").String = "Related Follow-up";
                Utilities.Application.SBO_Application.Menus.Item("5909").String = "Engine/VIN Transactions Report";
                Utilities.Application.SBO_Application.Menus.Item("6028").String = "Related Enquiries";
            }
        }

    }
}
