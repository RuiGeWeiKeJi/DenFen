using DevExpress . XtraBars;

namespace Base
{
    public partial class FormBase :DevExpress . XtraEditors . XtraForm
    {
        public FormBase ( )
        {
            InitializeComponent ( );
        }

        protected virtual int Export ( )
        {
            return 0;
        }

        private void toolExport_ItemClick ( object sender ,ItemClickEventArgs e )
        {
            Export ( );
        }
    }
}