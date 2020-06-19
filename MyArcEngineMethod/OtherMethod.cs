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
        /// ��Excel�ļ���ȡ�����洢��datatable��
        /// </summary>
        /// <param name="excelpath">Excel�ļ�·��</param>
        /// <param name="sheetName">Ҫ�����Ĺ���������</param>
        /// <param name="column">��Ҫ�������У����磺A2:D5</param>
        /// <returns></returns>
        public static System.Data.DataTable ReadExcel(string excelpath, string sheetName, string column)
        {
            //�½�һ�����ݱ�
            System.Data.DataTable dt = new System.Data.DataTable();
            string connectionStr = string.Empty;
            FileInfo file0 = new FileInfo(excelpath);
            if (!file0.Exists)
                throw new Exception("�ļ�������");
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
                { throw new Exception("��Excel�ļ���δ�ҵ�ָ����������," + ex.Message); }
                dt.TableName = sheetName;

                if (dt.Rows.Count < 2)
                    throw new Exception("���б���������ݣ�");

                //dt.Rows.RemoveAt(0);
                System.Data.DataRow headRow = dt.Rows[0];
                foreach (System.Data.DataColumn c in dt.Columns)
                {
                    string headValue = (headRow[c.ColumnName] == DBNull.Value ||
                        headRow[c.ColumnName] == null) ? "" : headRow[c.ColumnName].ToString().Trim();
                    if (headValue.Length == 0)
                    { throw new Exception("���������б���"); }
                    if (dt.Columns.Contains(headValue))
                    { throw new Exception("�������ظ����б��⣺" + headValue); }
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
       /// ��������ͬ���ļ�����һĿ¼��
       /// </summary>
       /// <param name="sourcePath">�������ļ�·��</param>
       /// <param name="targetDirectory">Ŀ��Ŀ¼</param>
       /// <param name="newName">���ƺ��ļ������ƣ�����������</param>
       /// <returns></returns>
        public static bool CopySameNameFiles(string sourcePath, string targetDirectory, string newName)
        {
            string sorceDirectory = System.IO.Path.GetDirectoryName(sourcePath);
            string fileName = System.IO.Path.GetFileNameWithoutExtension(sourcePath);

            bool isOk = false;

            //������·���Ƿ���Ŀ¼�ָ��������ַ�
            if (targetDirectory[targetDirectory.Length - 1] != System.IO.Path.DirectorySeparatorChar)
                targetDirectory += System.IO.Path.DirectorySeparatorChar;
            //���Ŀ��Ŀ¼�����ڣ��򴴽���
            if (!System.IO.Directory.Exists(targetDirectory))
                System.IO.Directory.CreateDirectory(targetDirectory);


            //ɾ��Ŀ��·���µ�ͬ���ļ�
            DirectoryInfo root = new DirectoryInfo(targetDirectory);
            foreach (FileInfo f in root.GetFiles())
            {
                if (f.Name.Split('.')[0] == newName)
                {
                    string filepath = f.FullName;
                    File.Delete(filepath);
                }
            }

            //��ȡ�ļ����µ������ļ�����ʼ����
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
        /// ɾ��Ĭ��Ŀ¼������ͬ���ļ�
        /// </summary>
        /// <param name="filePath">��ɾ���ļ�·��</param>
        public static void DeleteSameNameFiles(string filePath)
        {
            if (File.Exists(filePath)) //���ڸ�·���������ļ������ҵ��ļ���ɾ��
            {
                string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath); //�ļ�����
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
        /// ��ȡ��ͼ���·��
        /// </summary>
        /// <param name="curLayer"></param>
        /// <returns></returns>
        public static string GetFeatureLayerPath(IFeatureLayer curLayer)
        {
            string shpName0 = curLayer.FeatureClass.AliasName;
            IDataLayer dly = curLayer as IDataLayer;
            IDatasetName dn = dly.DataSourceName as IDatasetName;
            IWorkspaceName wn = dn.WorkspaceName;
            string currentShpPath = wn.PathName + "\\" + shpName0 + ".shp";//shp�ļ�·��
            return currentShpPath;
        }

        
        /// <summary>
        /// ��ȡ���е�txt�ļ����ݵ�string list��
        /// </summary>
        /// <param name="txtPath">txt·��</param>
        /// <returns></returns>
        [Obsolete("�Ƽ�ֱ��ʹ��System.IO.File��Ķ�ȡ����")]
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
        /// ��list�е�����д��txt�ļ���
        /// </summary>
        /// <param name="arrayList">string list</param>
        /// <param name="txtPath">txt·���������������Զ�����</param>
        /// <param name="isAppend">�Ƿ�Ϊ׷��ģʽ</param>
        /// <param name="encoding">���뷽ʽ</param>
        /// <returns></returns>
         [Obsolete("�Ƽ�ֱ��ʹ��System.IO.File��Ķ�ȡ����")]
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
        /// �õ�·���µ������ļ����µ�ĳ��׺�ļ�
        /// </summary>
        /// <param name="targetDir">Ŀ��Ŀ¼</param>
        /// <param name="allPath">���е�·����list</param>
        /// <param name="filter">��������Ĭ��Ϊ��.txt��</param>
        /// <returns></returns>
        public  static bool GetAllFilePath(string targetDir, List<string> allPath, string filter = "*.txt")
        {
            //��������ǲ���·��
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
