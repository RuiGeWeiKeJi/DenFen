using DevExpress . Utils . Paint;
using DevExpress . XtraCharts;
using DevExpress . XtraPrinting;
using DevExpress . XtraPrintingLinks;
using System;
using System . Data;
using System . Drawing;
using System . IO;
using System . Reflection;
using System . Windows . Forms;

namespace Base
{
    public partial class FormReport :FormBase
    {
        public FormReport ( )
        {
            InitializeComponent ( );

            Utility . GridViewMoHuSelect . SetFilter ( gridView1 );

            FieldInfo fi = typeof ( DevExpress . Utils . Paint . XPaint ) . GetField ( "graphics" ,BindingFlags . Static | BindingFlags . NonPublic );
            fi . SetValue ( null ,new DrawXPaint ( ) );

            ReportBll . Bll . ReportBll _bll = new ReportBll . Bll . ReportBll ( );
            DataTable tableView = _bll . GetDataTable ( );
            gridControl1 . DataSource = tableView;

            DateTime dt = _bll . getDt ( );
            HCB00601 . Caption = dt . AddDays ( 2 ) . ToString ( "MM.dd" ) + "之前";
            HCB00602 . Caption = dt . AddDays ( 5 ) . ToString ( "MM.dd" ) + "之前";
            HCB00603 . Caption = dt . AddDays ( 8 ) . ToString ( "MM.dd" ) + "之前";
            HCB00604 . Caption = dt . AddDays ( 8 ) . ToString ( "MM.dd" ) + "之后";

            RCB00801 . Caption = dt . Year . ToString ( ) + "年" + RCB00801 . Caption;
            RCB00802 . Caption = dt . Year . ToString ( ) + "年" + RCB00802 . Caption;
            RCB00803 . Caption = dt . Year . ToString ( ) + "年" + RCB00803 . Caption;
            RCB00804 . Caption = dt . Year . ToString ( ) + "年" + RCB00804 . Caption;
            RCB00805 . Caption = dt . Year . ToString ( ) + "年" + RCB00805 . Caption;
            RCB00806 . Caption = dt . Year . ToString ( ) + "年" + RCB00806 . Caption;
            RCB00807 . Caption = dt . Year . ToString ( ) + "年" + RCB00807 . Caption;
            RCB00808 . Caption = dt . Year . ToString ( ) + "年" + RCB00808 . Caption;
            RCB00809 . Caption = dt . Year . ToString ( ) + "年" + RCB00809 . Caption;
            RCB00810 . Caption = dt . Year . ToString ( ) + "年" + RCB00810 . Caption;
            RCB00811 . Caption = dt . Year . ToString ( ) + "年" + RCB00811 . Caption;
            RCB00812 . Caption = dt . Year . ToString ( ) + "年" + RCB00812 . Caption;
        }

        protected override int Export ( )
        {
            ExportToExcel ( this . Text ,gridControl1 );

            return base . Export ( );
        }

        /// <summary>
        /// DevExpress通用导出Excel,支持多个控件同时导出在同一个Sheet表
        /// eg:ExportToXlsx("",gridControl1,gridControl2);
        /// 将gridControl1和gridControl2的数据一同导出到同一张工作表
        /// </summary>
        /// <param name="title">文件名</param>
        /// <param name="panels">控件集</param>
        public void ExportToExcel ( string title ,params IPrintable [ ] panels )
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog ( );
            saveFileDialog . FileName = title;
            saveFileDialog . Title = "导出Excel";
            saveFileDialog . Filter = "Excel文件(*.xlsx)|*.xlsx|Excel文件(*.xls)|*.xls";
            DialogResult dialogResult = saveFileDialog . ShowDialog ( );
            if ( dialogResult == DialogResult . Cancel )
                return;
            string FileName = saveFileDialog . FileName;
            PrintingSystem ps = new PrintingSystem ( );
            CompositeLink link = new CompositeLink ( ps );
            ps . Links . Add ( link );
            foreach ( IPrintable panel in panels )
            {
                link . Links . Add ( CreatePrintableLink ( panel ) );
            }
            link . Landscape = true;//横向           
            //判断是否有标题，有则设置         
            //link.CreateDocument(); //建立文档
            try
            {
                int count = 1;
                //在重复名称后加（序号）
                while ( File . Exists ( FileName ) )
                {
                    if ( FileName . Contains ( ")." ) )
                    {
                        int start = FileName . LastIndexOf ( "(" );
                        int end = FileName . LastIndexOf ( ")." ) - FileName . LastIndexOf ( "(" ) + 2;
                        FileName = FileName . Replace ( FileName . Substring ( start ,end ) ,string . Format ( "({0})." ,count ) );
                    }
                    else
                    {
                        FileName = FileName . Replace ( "." ,string . Format ( "({0})." ,count ) );
                    }
                    count++;
                }

                if ( FileName . LastIndexOf ( ".xlsx" ) >= FileName . Length - 5 )
                {
                    XlsxExportOptions options = new XlsxExportOptions ( );
                    link . ExportToXlsx ( FileName ,options );
                }
                else
                {
                    XlsExportOptions options = new XlsExportOptions ( );
                    link . ExportToXls ( FileName ,options );
                }
                if ( DevExpress . XtraEditors . XtraMessageBox . Show ( "保存成功，是否打开文件？" ,"提示" ,MessageBoxButtons . YesNo ,MessageBoxIcon . Information ) == DialogResult . Yes )
                    System . Diagnostics . Process . Start ( FileName );//打开指定路径下的文件
            }
            catch ( Exception ex )
            {
                DevExpress . XtraEditors . XtraMessageBox . Show ( ex . Message );
            }
        }

        /// <summary>
        /// 创建打印Componet
        /// </summary>
        /// <param name="printable"></param>
        /// <returns></returns>
        PrintableComponentLink CreatePrintableLink ( IPrintable printable )
        {
            ChartControl chart = printable as ChartControl;
            if ( chart != null )
                chart . OptionsPrint . SizeMode = DevExpress . XtraCharts . Printing . PrintSizeMode . Stretch;
            PrintableComponentLink printableLink = new PrintableComponentLink ( ) { Component = printable };
            return printableLink;
        }

    }

    public class DrawXPaint :XPaint
    {
        public override void DrawFocusRectangle ( Graphics g ,Rectangle r ,Color foreColor ,Color backColor )
        {
            base . DrawFocusRectangle ( g ,r ,foreColor ,backColor );
            if ( !CanDraw ( r ) )
                return;
            Brush hb = Brushes . Black;
            g . FillRectangle ( hb ,new Rectangle ( r . X ,r . Y ,2 ,r . Height - 2 ) );//Left
            g . FillRectangle ( hb ,new Rectangle ( r . X ,r . Y ,r . Width - 2 ,2 ) );//Top
            g . FillRectangle ( hb ,new Rectangle ( r . Right - 2 ,r . Y ,2 ,r . Height - 2 ) );//Right
            g . FillRectangle ( hb ,new Rectangle ( r . X ,r . Bottom - 2 ,r . Width - 2 ,2 ) );//Bottom    
        }
    }
}


