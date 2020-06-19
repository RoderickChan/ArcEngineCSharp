using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyArcEngineMethod.MyAttribute
{
       [AttributeUsage(AttributeTargets.All, Inherited = false, AllowMultiple = true)]
    public sealed class WaitCorrectedAttribute : Attribute
    {
        readonly string positionalString;
        // This is a positional argument
        public WaitCorrectedAttribute(string positionalString)
        {
            this.positionalString = positionalString;
        }
        public WaitCorrectedAttribute(string positionalStr, string reason, string discription)
        {
            this.positionalString = positionalStr;
            this.Reason = reason;
            this.Discription = discription;

        }

        /// <summary>
        /// 定位信息
        /// </summary>
        public string PositionalString
        {
            get { return positionalString; }
        }
        /// <summary>
        /// 修正的原因
        /// </summary>
        public string Reason { get; set; }

        /// <summary>
        /// 待修正的代码描述
        /// </summary>
        public string Discription { get; set; }
    }
   
}
