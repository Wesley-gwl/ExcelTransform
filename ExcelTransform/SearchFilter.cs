using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTransform
{
    /// <summary>
    /// 高级查询所使用的查询参数
    /// </summary>
    public class SearchFilter
    {
        /// <summary>
        /// 过滤条件中使用的数据列
        /// </summary>
        public string Property { get; set; }

        /// <summary>
        /// 过滤条件中的操作:==、!=等
        /// </summary>
        public string Condition { get; set; }

        /// <summary>
        /// 过滤条件之间的逻辑关系：AND和OR
        /// </summary>
        public string Relation { get; set; }

        /// <summary>
        /// 过滤条件中的查询值
        /// </summary>
        public string SearchValue { get; set; }
    }
}