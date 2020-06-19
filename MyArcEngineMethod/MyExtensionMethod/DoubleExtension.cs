using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyArcEngineMethod.MyExtensionMethod
{
    public static class DoubleExtension
    {
        /// <summary>
        /// 保留几位小数
        /// </summary>
        /// <param name="x">小数</param>
        /// <param name="deci">保留小数的位数</param>
        /// <returns></returns>
        public static double Round(this double x, int deci)
        {
            return Math.Round(x, deci);
        }

        /// <summary>
        /// 比较两个浮点数是否相等，可指定精度
        /// </summary>
        /// <param name="x">被比较的数</param>
        /// <param name="y">比较的数</param>
        /// <param name="deci">精度，取几位小数进行比较</param>
        /// <returns>小于返回-1 等于返回0 大于返回1</returns>
        public static int Compare(this double x, double y, int deci = 4)
        {
            double d1 = Math.Round(x, deci);
            double d2 = Math.Round(y, deci);
            if (d1 - d2 < 0)
                return -1;
            else if (d1 == d2)
                return 0;
            else
                return 1;
        }
    }
}
