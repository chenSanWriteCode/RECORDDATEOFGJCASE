using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GJOPENCASERECORD.ENTITY
{
    public class OpenRecord
    {
        /// <summary>
        /// 线路名称
        /// </summary>
        public string  lineName { get; set; }
        /// <summary>
        /// 车牌号
        /// </summary>
        public string carNum { get; set; }
        /// <summary>
        /// 持卡人
        /// </summary>
        public string owiner { get; set; }
        /// <summary>
        /// 换出内胆编号
        /// </summary>
        public string ouCardNum { get; set; }
        /// <summary>
        /// 换入内胆编号
        /// </summary>
        public string inCardNum { get; set; }
        /// <summary>
        /// 记录时间
        /// </summary>
        public string  recordTime { get; set; }

    }
}
