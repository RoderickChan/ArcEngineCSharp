using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyArcEngineMethod.MyAttribute
{
    [AttributeUsage(AttributeTargets.All)]
    public class AuthorAttribute:System.Attribute
    {
      
        private string _author;

        /// <summary>
        /// 作者
        /// </summary>
        public string Author
        {
            get { return _author; }
            set { _author = value; }
        }
     
        private string _version;

        /// <summary>
        /// 版本号
        /// </summary>
        public string Version
        {
            get { return _version; }
            set { _version = value; }
        }

        private string  _date;

        /// <summary>
        /// 添加日期
        /// </summary>
        public string  Date
        {
            get { return _date; }
            set { _date = value; }
        }
        
        public AuthorAttribute(string author):this(author,null, null)
        {
        }

        public AuthorAttribute(string author, string version):this(author,version,null)
        {          
        }

        public AuthorAttribute(string author, string version, string date)
        {
            this._author = author;
            this._version = version;
            this._date = date;
        }
        

    }
}
