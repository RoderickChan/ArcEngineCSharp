using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyArcEngineMethod.MyAttribute
{
    [AttributeUsage(AttributeTargets.All, Inherited = false, AllowMultiple = true)]
    public sealed class UnFinishedAttribute:System.Attribute
    {    
           //作者
           private string _author;

           /// <summary>
           /// 作者
           /// </summary>
           public string Author
           {
               get { return _author; }
               set { _author = value; }
           }
           //日期
           private string _dateTime;

           /// <summary>
           /// 日期
           /// </summary>
           public string DateTime
           {
               get { return _dateTime; }
               set { _dateTime = value; }
           }
           //描述
           private string _discription;

           /// <summary>
           /// 描述
           /// </summary>
           public string Discription
           {
               get { return _discription; }
               set { _discription = value; }
           }
           public UnFinishedAttribute(string author,string dateTime, string discription)
           {
               _author = author;
               _dateTime = dateTime;
               _discription = discription;
           }

         
       

    
      
 
   }
}
