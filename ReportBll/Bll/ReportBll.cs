using System;
using System . Collections . Generic;
using System . Data;
using System . Linq;
using System . Text;

namespace ReportBll . Bll
{
    public class ReportBll
    {
        private readonly Dao.ReportDao dal=null;
        public ReportBll ( )
        {
            dal = new Dao . ReportDao ( );
        }

        /// <summary>
        /// 获取数据列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetDataTable ( )
        {
            return dal . GetDataTable ( );
        }

        /// <summary>
        /// 获取服务器日期
        /// </summary>
        /// <returns></returns>
        public DateTime getDt ( )
        {
            return dal . getDt ( );
        }
    }
}
