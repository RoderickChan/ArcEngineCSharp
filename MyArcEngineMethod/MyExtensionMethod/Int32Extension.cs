using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyArcEngineMethod.MyExtensionMethod
{
   public static class Int32Extension
    {
       /// <summary>
       /// 求整数的n次幂的值
       /// </summary>
       /// <param name="x">整数</param>
       /// <param name="n">幂</param>
       /// <returns></returns>
       public static int Power(this int x, int n)
       {
           if (x < 0)
               throw new InvalidOperationException("Cannot operate a negative integer!");
           else if (x == 0)
               return 1;
           else
               return  (int)Math.Pow(x, n);
       }

       /// <summary>
       /// 求整数的n次幂的值
       /// </summary>
       /// <param name="x">整数</param>
       /// <param name="n">幂</param>
       /// <returns></returns>
       public static double Power(this int x, double n)
       {
           if (x < 0)
               throw new InvalidOperationException("Cannot operate a negative integer!");
           else if (x == 0)
               return 1;
           else
               return Math.Pow(x, n);
       }

    }
}
