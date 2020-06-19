using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
/*
 * 日志记录，就记录在当前运行的文件目录下
 * */
namespace MyArcEngineMethod.MyLogger
{
    /// <summary>
    /// 写入日志文件的类型
    /// </summary>
    public enum LogType 
    {
        Test,Fine,Info,Warn,Error,Message,
    }

   public class MyLogger:IDisposable
    {
       //日志文件路径，不存在则创建
       private string logTxtPath = AppDomain.CurrentDomain.BaseDirectory+"MyLogger.txt";
       private bool disposed;//标识资源是否已被释放
       private ReaderWriterLockSlim rwls = new ReaderWriterLockSlim();
       StreamWriter _sw = null;
       public StackTrace stackTrace { get; set; }


       internal MyLogger(StackTrace st, bool append = true)
       {
           stackTrace = st;
           if (!File.Exists(logTxtPath))
           {
               _sw = File.CreateText(logTxtPath);
               append = false;
           }
           else
               _sw = new StreamWriter(logTxtPath, append, Encoding.UTF8, 1024);
           if (!append)
           {
               string str = Process.GetCurrentProcess().MainModule.FileName;
               _sw.WriteLine(String.Format("This is a logging file of application '{0}'", str));
               _sw.WriteLine("Created Time:" + DateTime.Now.ToString("F"));
               _sw.WriteLine("Author:" + "Lynne Chan");
               _sw.WriteLine();
           }
       }
       ~MyLogger()
       {
           Dispose(false);
       }

       /// <summary>
       /// 手动关闭资源流
       /// </summary>
       public void Close()
       {
           Dispose();
       }

       /// <summary>
       /// 异步打印日志，良好
       /// </summary>
       /// <param name="n">指定调用栈的层数</param>
       public async void Fine( int n =1)
       {
           var info = GetInfo(n);
           string str1 = info[0];
           await _sw.WriteLineAsync("Fine in "+str1);
       }

       /// <summary>
       /// 异步打印日志，信息
       /// </summary>
       /// <param name="n">指定调用栈的层数</param>
       public async void Info(int n = 1)
       {
           var info = GetInfo(n);
           string str1 = info[0];
           await _sw.WriteLineAsync("Info in " + str1);
       }

       /// <summary>
       /// 异步打印日志，警告
       /// </summary>
       /// <param name="n">指定调用栈的层数</param>
       public async void Warn(int n = 1)
       {
           var info = GetInfo(n);
           string str1 = info[0];
           await _sw.WriteLineAsync("Warn in " + str1);
       }

       /// <summary>
       /// 异步打印日志，错误
       /// </summary>
       /// <param name="n">指定调用栈的层数</param>
       public async void Error(int n = 1)
       {
           var info = GetInfo(n);
           string str1 = info[0];
           await _sw.WriteLineAsync("Error in " + str1+"trace:"+info[1]);
       }

       /// <summary>
       /// 异步打印日志，输出指定信息
       /// </summary>
       /// <param name="mess">指定的信息</param>
       public async void Message(String mess)
       {
           string str = string.Format("Message:{0}; Time:{1}\n", mess, DateTime.Now.ToString("F"));
          await _sw.WriteLineAsync(str);
       }

       /// <summary>
       /// 开启读写锁，写入日志
       /// </summary>
       /// <param name="logType">写入日志的类型</param>
       /// /// <param name="n">调用栈的层数</param>
       /// /// <param name="message">若为Message类型，需要指定该参数</param>
       /// <returns>返回是否写入成功</returns>
       public Boolean WirteLog(LogType logType,int n = 1,string message = null)
       {
           try
           {
               rwls.EnterReadLock();
               string write = "";
               switch (logType)
               {
                   case LogType.Test:
                       write = "Test";
                       break;
                   case LogType.Fine:
                       write = GetFine(n);
                       break;
                   case LogType.Info:
                       write = GetInfo2(n);
                       break;
                   case LogType.Warn:
                       write = GetWarn(n);
                       break;
                   case LogType.Error:
                       write = GetError(n);
                       break;
                   case LogType.Message:
                       write = GetMessage(message);
                       break;
                   default:
                       break;
               }
               _sw.WriteLine(write);
           }
           finally
           {
               rwls.ExitWriteLock();

           }

           return true;
       }


       /// <summary>
       /// 释放非托管资源
       /// </summary>
       public void Dispose()
       {
           Dispose(true);
           //通知垃圾回收器不再调用终结器
           GC.SuppressFinalize(this);
       }

       /// <summary>
       /// 释放资源
       /// </summary>
       /// <param name="disposing">是否已清理过非托管资源</param>
       protected virtual void Dispose(Boolean disposing)
       {
           if (disposed) return;
           //清理托管资源
           if (disposing)
           {
               if (_sw != null)
                   _sw.Flush();
                   _sw.Close();
           }
           //标记为已释放
           disposed = true;
 
       }

       /// <summary>
       /// 获取调用函数的基本信息，包括行号，文件名，方法名
       /// </summary>
       /// <param name="n">指定输出调用栈的层数，默认为1</param>
       /// <returns></returns>
       private String[] GetInfo(int n = 1)
       {
           string sTrace = stackTrace.ToString();
           StackFrame sf = stackTrace.GetFrame(n);
           //获取当前所在的函数的函数名，行号，文件名
           string str = string.Format("Line:{0}; Method Name:{1}; File Name:{2}; Time:{3}\n",
                sf.GetFileLineNumber(),
               sf.GetMethod().Name,
               sf.GetFileName(),
               DateTime.Now.ToString("F")
               );

           return new string[] { str, sTrace };
       }

       /// <summary>
       /// 打印日志，良好
       /// </summary>
       /// <param name="n">指定调用栈的层数</param>
       private string GetFine(int n = 1)
       {
           var info = GetInfo(n);
           string str1 = info[0];
           return "Fine in " + str1;
       }

       /// <summary>
       /// 打印日志，信息
       /// </summary>
       /// <param name="n">指定调用栈的层数</param>
       public  string GetInfo2(int n = 1)
       {
           var info = GetInfo(n);
           string str1 = info[0];
           return "Info in " + str1;
       }

       /// <summary>
       /// 打印日志，警告
       /// </summary>
       /// <param name="n">指定调用栈的层数</param>
       public  string GetWarn(int n = 1)
       {
           var info = GetInfo(n);
           string str1 = info[0];
           return "Warn in " + str1;
       }

       /// <summary>
       /// 打印日志，错误
       /// </summary>
       /// <param name="n">指定调用栈的层数</param>
       public string GetError(int n = 1)
       {
           var info = GetInfo(n);
           string str1 = info[0];
           return "Error in " + str1 + "trace:" + info[1];
       }

       /// <summary>
       /// 打印日志，输出指定信息
       /// </summary>
       /// <param name="mess">指定的信息</param>
       public string GetMessage(String mess)
       {
           string str = string.Format("Message:{0}; Time:{1}\n", mess, DateTime.Now.ToString("F"));
           return str;
       }

    }


}
