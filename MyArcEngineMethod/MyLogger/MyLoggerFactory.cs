using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;

namespace MyArcEngineMethod.MyLogger
{
 public   class MyLoggerFactory
    {
     /// <summary>
     /// 获取实例
     /// </summary>
     /// <returns></returns>
     public static MyLogger GetInstance(StackTrace st)
     {
         MyLogger log = new MyLogger(st);
         return log;    
     }

     /// <summary>
     /// 获取实例
     /// </summary>
     /// <returns></returns>
     public static MyLogger GetInstance(StackTrace st, Boolean append)
     {      
         MyLogger log = new MyLogger(st,append);
         return log;
     }


     /// <summary>
     ///  异步获取实例
     /// </summary>
     /// <param name="st">StackTrace实例</param>
     /// <param name="append">是否为追加模式</param>
     /// <returns></returns>
     public async static Task<MyLogger> GetInstanceOnNewThread(StackTrace st, Boolean append)
     {
         await Task.Run<MyLogger>(
             () => {
                 MyLogger log = new MyLogger(st, append);
                 return log;

             }
             );
         return null;
     }

    }
}
