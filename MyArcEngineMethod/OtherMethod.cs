using System.Collections.Generic;
using System.Data.OleDb;
using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.AnalysisTools;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geoprocessor;
using ESRI.ArcGIS.Carto;

namespace MyArcEngineMethod
{
    [Guid("d6268b8d-0e92-4522-beef-c3f1e3dea586")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyArcEngineMethod.OtherMethod")]
    public class OtherMethod
    {
        /// <summary>
        /// 将Excel文件读取出来存储在datatable中
        /// </summary>
        /// <param name="excelpath">Excel文件路径</param>
        /// <param name="sheetName">要导出的工作表名称</param>
        /// <param name="column">需要导出的列，形如：A2:D5</param>
        /// <returns></returns>
        public static System.Data.DataTable ReadExcel(string excelpath, string sheetName, string column)
        {
            //新建一个数据表
            System.Data.DataTable dt = new System.Data.DataTable();
            string connectionStr = string.Empty;
            FileInfo file0 = new FileInfo(excelpath);
            if (!file0.Exists)
                throw new Exception("文件不存在");
            string extension0 = file0.Extension;
            switch (extension0)
            {
                case ".xls":
                    connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelpath
                        + ";Extended Properties='Excel 8.0;HDR=no;IMEX=1;'";
                    break;
                case ".xlsx":
                    connectionStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelpath
                        + ";Extended Properties='Excel 12.0;HDR=no;IMEX=1;'";
                    break;
                default:
                    connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelpath
                        + ";Extended Properties='Excel 8.0;HDR=no;IMEX=1;'";
                    break;
            }
            OleDbConnection oldbcon = new OleDbConnection(connectionStr);
            try
            {
                oldbcon.Open();
                System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand("select * from ["
                    + sheetName + "$" + column + "] where F1 is not null ", oldbcon);
                System.Data.OleDb.OleDbDataAdapter apt = new System.Data.OleDb.OleDbDataAdapter(cmd);
                try
                {
                    apt.Fill(dt);
                }
                catch (Exception ex)
                { throw new Exception("该Excel文件中未找到指定工作表名," + ex.Message); }
                dt.TableName = sheetName;

                if (dt.Rows.Count < 2)
                    throw new Exception("表中必须包含数据！");

                //dt.Rows.RemoveAt(0);
                System.Data.DataRow headRow = dt.Rows[0];
                foreach (System.Data.DataColumn c in dt.Columns)
                {
                    string headValue = (headRow[c.ColumnName] == DBNull.Value ||
                        headRow[c.ColumnName] == null) ? "" : headRow[c.ColumnName].ToString().Trim();
                    if (headValue.Length == 0)
                    { throw new Exception("必须输入列标题"); }
                    if (dt.Columns.Contains(headValue))
                    { throw new Exception("不能用重复的列标题：" + headValue); }
                    c.ColumnName = headValue;
                }
                dt.Rows.RemoveAt(0);

                return dt;

            }
            catch (Exception ee)
            {
                //Console.WriteLine( ee.StackTrace);
                // Console.ReadLine();
                throw ee;
                //return null;
            }
            finally
            { oldbcon.Close(); }


        }

       /// <summary>
       /// 复制所有同名文件到另一目录中
       /// </summary>
       /// <param name="sourcePath">待复制文件路径</param>
       /// <param name="targetDirectory">目标目录</param>
       /// <param name="newName">复制后文件的名称（可重命名）</param>
       /// <returns></returns>
        public static bool CopySameNameFiles(string sourcePath, string targetDirectory, string newName)
        {
            string sorceDirectory = System.IO.Path.GetDirectoryName(sourcePath);
            string fileName = System.IO.Path.GetFileNameWithoutExtension(sourcePath);

            bool isOk = false;

            //检查输出路径是否以目录分隔符隔开字符
            if (targetDirectory[targetDirectory.Length - 1] != System.IO.Path.DirectorySeparatorChar)
                targetDirectory += System.IO.Path.DirectorySeparatorChar;
            //如果目标目录不存在，则创建它
            if (!System.IO.Directory.Exists(targetDirectory))
                System.IO.Directory.CreateDirectory(targetDirectory);


            //删除目标路径下的同名文件
            DirectoryInfo root = new DirectoryInfo(targetDirectory);
            foreach (FileInfo f in root.GetFiles())
            {
                if (f.Name.Split('.')[0] == newName)
                {
                    string filepath = f.FullName;
                    File.Delete(filepath);
                }
            }

            //获取文件夹下的所有文件，开始复制
            string[] allFile = System.IO.Directory.GetFiles(sorceDirectory);
            //System.IO.File.Copy();

            for (int i = 0; i < allFile.Length; i++)
            {
                string curFile = allFile[i];
                string curFileName = System.IO.Path.GetFileNameWithoutExtension(curFile);
                if (curFileName == fileName)
                {
                    string outPatth = System.IO.Path.Combine(targetDirectory, newName) + System.IO.Path.GetExtension(curFile);
                    System.IO.File.Copy(curFile, outPatth, true);
                    isOk = true;
                }
            }

            return isOk;

        }

        /// <summary>
        /// 删除默认目录下所有同名文件
        /// </summary>
        /// <param name="filePath">待删除文件路径</param>
        public static void DeleteSameNameFiles(string filePath)
        {
            if (File.Exists(filePath)) //存在该路径，则在文件夹中找到文件并删除
            {
                string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath); //文件名称
                DirectoryInfo root = new DirectoryInfo(System.IO.Path.GetDirectoryName(filePath));
                foreach (FileInfo f in root.GetFiles())
                {
                    if (f.Name.Split('.')[0] == fileName)
                    {
                        string filepath = f.FullName;
                        File.Delete(filepath);
                    }
                }
            }
        }

        /// <summary>
        /// 获取到图层的路径
        /// </summary>
        /// <param name="curLayer"></param>
        /// <returns></returns>
        public static string GetFeatureLayerPath(IFeatureLayer curLayer)
        {
            string shpName0 = curLayer.FeatureClass.AliasName;
            IDataLayer dly = curLayer as IDataLayer;
            IDatasetName dn = dly.DataSourceName as IDatasetName;
            IWorkspaceName wn = dn.WorkspaceName;
            string currentShpPath = wn.PathName + "\\" + shpName0 + ".shp";//shp文件路径
            return currentShpPath;
        }

        
        /// <summary>
        /// 读取所有的txt文件内容到string list中
        /// </summary>
        /// <param name="txtPath">txt路径</param>
        /// <returns></returns>
        [Obsolete("推荐直接使用System.IO.File类的读取方法")]
        public static List<string> ReadTxt2StrList(string txtPath)
        {
            List<string> readArrayList = new List<string>();
           StreamReader streamReader = new StreamReader(txtPath, Encoding.UTF8);
            string line = "";
            while ((line = streamReader.ReadLine()) != null)
            {
                readArrayList.Add(line);
            }
            streamReader.Close();
            return readArrayList;
        }

       
        /// <summary>
        /// 将list中的内容写入txt文件中
        /// </summary>
        /// <param name="arrayList">string list</param>
        /// <param name="txtPath">txt路径，若不存在则自动创建</param>
        /// <param name="isAppend">是否为追加模式</param>
        /// <param name="encoding">编码方式</param>
        /// <returns></returns>
         [Obsolete("推荐直接使用System.IO.File类的读取方法")]
        public static bool WriteArryList2Txt(List<string> arrayList, string txtPath, bool isAppend, Encoding encoding)
        {
            StreamWriter streamWriter = new StreamWriter(txtPath, isAppend, encoding);
            foreach (var item in arrayList)
            {
                streamWriter.WriteLine(item);
            }
            streamWriter.Flush();
            streamWriter.Close();
            return true;
        }

        /// <summary>
        /// 得到路径下的所有文件夹下的某后缀文件
        /// </summary>
        /// <param name="targetDir">目标目录</param>
        /// <param name="allPath">所有的路径的list</param>
        /// <param name="filter">过滤器，默认为“.txt”</param>
        /// <returns></returns>
        public  static bool GetAllFilePath(string targetDir, List<string> allPath, string filter = "*.txt")
        {
            //检查输入是不是路径
            if (!Directory.Exists(targetDir)) throw new ArgumentException("The target directory is invalid!");
            var dirInfo = new DirectoryInfo(targetDir);
            var allDirs = dirInfo.GetDirectories();
            foreach (var dir in allDirs)
            {
                GetAllFilePath(dir.FullName, allPath);
            }
            var allFiles = dirInfo.GetFiles(filter);
            foreach (var fileInfo in allFiles)
            {
                allPath.Add(fileInfo.FullName);
            }

            return true; ;
        }


    }
}
