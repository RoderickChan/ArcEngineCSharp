using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Collections.Generic;
using ESRI.ArcGIS.AnalysisTools;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geoprocessor;
using Microsoft.Office.Interop.Excel;

namespace MyArcEngineMethod
{
    [Guid("bd65d0e7-b782-4dcf-aa27-199a562f2c96")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyArcEngineMethod.FieldOperation")]
    public class FieldOperation
    {

        /// <summary>
        /// 根据字段名为要素类创建字段
        /// </summary>
        /// <param name="featClass">输入要素类</param>
        /// <param name="fieldType">字段类型</param>
        /// <param name="fieldsname">字段名称，可变参量</param>
        public static void CreateFieldsByName(IFeatureClass featClass, esriFieldType fieldType, params string[] fieldsname)
        {
            int index1 = -1;
            for (int i = 0; i < fieldsname.Length; i++)
            {
                index1 = featClass.Fields.FindField(fieldsname[i]);
                if (index1 == -1)
                {
                    IField pField = new FieldClass();
                    IFieldEdit pFieldEdit = pField as IFieldEdit;
                    pFieldEdit.Type_2 = fieldType;
                    pFieldEdit.Name_2 = fieldsname[i];
                    featClass.AddField(pField);
                }
            }
        }

        /// <summary>
        /// 创建字段并且得到字段的索引
        /// </summary>
        /// <param name="featClass">待创建要素类的字段</param>
        /// <param name="fieldType">字段类型</param>
        /// <param name="fieldsname">字段名称</param>
        /// <returns>返回新建字段的索引，若存在，返回其固有的索引</returns>
        public static int[] CreateFieldsAndGetIndex(IFeatureClass featClass, esriFieldType fieldType, params string[] fieldsname)
        {
            int[] indexArray = new int[fieldsname.Length];
            for (int i = 0; i < fieldsname.Length; i++)
            {
                int index1 = featClass.Fields.FindField(fieldsname[i]);
                if (index1 == -1)
                {
                    IField pField = new FieldClass();
                    IFieldEdit pFieldEdit = pField as IFieldEdit;
                    pFieldEdit.Type_2 = fieldType;
                    pFieldEdit.Name_2 = fieldsname[i];
                    featClass.AddField(pField);
                }
                index1 = featClass.Fields.FindField(fieldsname[i]);
                indexArray[i] = index1;
            }
            return indexArray;
        }

        /// <summary>
        /// 创建一个字段并返回索引
        /// </summary>
        /// <param name="featClass">要素类</param>
        /// <param name="fieldType">字段类型</param>
        /// <param name="fieldsname">字段名称</param>
        /// <returns>字段的索引</returns>
        public static int CreateOneFieldAndGetIndex(IFeatureClass featClass, esriFieldType fieldType, string fieldsname)
        {
            int index1 = -1;
            index1 = featClass.Fields.FindField(fieldsname);
            if (index1 == -1)
            {
                IField pField = new FieldClass();
                IFieldEdit pFieldEdit = pField as IFieldEdit;
                pFieldEdit.Type_2 = fieldType;
                pFieldEdit.Name_2 = fieldsname;
                featClass.AddField(pField);
            }
            index1 = featClass.Fields.FindField(fieldsname);
            return index1;
        }

        /// <summary>
        /// 将要素类该字段的值均重置为空
        /// </summary>
        /// <param name="feaClass">要素类</param>
        /// <param name="fieldNames">字段名称</param>
        public static void RefreshFieldValue(IFeatureClass feaClass, params string[] fieldNames)
        {
            IFeatureCursor update = feaClass.Update(null, false);
            IFeature fea = null;
            var indexes = new int[fieldNames.Length];
            int index = -1;
            for (int i = 0; i < fieldNames.Length; i++)
            {
                index = feaClass.Fields.FindField(fieldNames[i]);
                indexes[i] = index;               
            }
            while ((fea = update.NextFeature()) != null)
            {
                foreach (var cIndex in indexes)
                {                   
                    if ( cIndex == -1)
                        continue;
                    fea.Value[cIndex] = null;                
                }
                update.UpdateFeature(fea);
            }

            Marshal.FinalReleaseComObject(update);
 
        }

        /// <summary>
        /// 删除要素类中不必要的字段
        /// </summary>
        /// <param name="feaClass">输入的要素类</param>
        /// <param name="fieldName">删除的字段名,可变参量</param>
        public static bool DeleteFieldByName(IFeatureClass feaClass, params string[] fieldName)
        {
            if (feaClass == null || fieldName.Length < 1)
                return false;

            int index = -1;
            for (int i = 0; i < fieldName.Length; i++)
            {
                index = feaClass.Fields.FindField(fieldName[i]);
                if (index == -1)
                    continue;
                IField field = feaClass.Fields.get_Field(index);
                feaClass.DeleteField(field);
            }

            return true;
        }

        
        /// <summary>
        /// 获取到要素类的某一字段枚举值，返回字符串类型的列表
        /// </summary>
        /// <param name="fc">输入的要素类</param>
        /// <param name="fieldName">待查询的字段名称</param>
        /// <returns></returns>
        [ObsoleteAttribute("不推荐采用这个方法，推荐使用下面的Table")]
        public static List<string> GetFeatureClassUniqueFieldValue(IFeatureClass fc, string fieldName)
        {
            List<string> allValues = new List<string>();
            IQueryFilter pQuery = new QueryFilterClass();
            pQuery.SubFields = fieldName;
            IFeatureCursor pCursor = fc.Search(pQuery, false);

            IDataStatistics pDataStatistic = new DataStatisticsClass();
            pDataStatistic.Field = fieldName;
            pDataStatistic.Cursor = pCursor as ICursor;

            IEnumerator pEnumerator = pDataStatistic.UniqueValues;
            pEnumerator.Reset();
            while (pEnumerator.MoveNext())
            {
                object ob1 = pEnumerator.Current;
                allValues.Add(ob1.ToString());
            }
            Marshal.FinalReleaseComObject(pCursor);
            return allValues;
        }

        /// <summary>
        /// 获取到属性表的某字段唯一值
        /// </summary>
        /// <param name="table">属性表</param>
        /// <param name="fieldName">字段名</param>
        /// <returns></returns>
       public static List<string> GetTableFieldUniqueValue(ITable table, string fieldName)
        {
            List<string> allValues = new List<string>();
            //IField f = fc.Fields.Field[fc.FindField(name)];
            //IFeatureCursor cursor = fc.Search(null, false);
            ICursor cursor = table.Search(null, false);
            IDataStatistics ds = new DataStatisticsClass();
            ds.Field = fieldName;
            ds.Cursor = cursor;
            var enumer = ds.UniqueValues;
            enumer.Reset();
            while (enumer.MoveNext())
            {
                var s = enumer.Current.ToString();
                allValues.Add(s);
            }
            Marshal.FinalReleaseComObject(cursor);
            return allValues;
        }

        /// <summary>
        /// 复制要素的字段属性
        /// </summary>
        /// <param name="sourceFea">源要素</param>
        /// <param name="targetFea">待复制要素</param>
        /// <param name="geo">几何形状，可为空</param>
        /// <param name="isStore">是否调用fea的store方法</param>
        /// <returns></returns>
        public static bool CopyFeatureFields(IFeature sourceFea, IFeature targetFea, IGeometry geo, bool isStore)
        {
            if (sourceFea == null || targetFea == null)
                return false;

            IFields copyFields = sourceFea.Fields;
            for (int i = 0; i < copyFields.FieldCount; i++)
            {
                IField curField = copyFields.get_Field(i);
                //在带插入要素中寻找该字段
                int indexPaste = targetFea.Fields.FindField(curField.Name);
                if (indexPaste != -1)
                {
                    IField pasteField = targetFea.Fields.get_Field(indexPaste);
                    //若字段可编辑，则复制字段值
                    if (pasteField.Editable && pasteField.Type != esriFieldType.esriFieldTypeGeometry)
                    {
                        targetFea.Value[indexPaste] = sourceFea.Value[i];
                    }
                }
            }
            if (geo != null)
                targetFea.Shape = geo;
            if(isStore == true)
                targetFea.Store();
            return true;
 
        }

        /// <summary>
        /// 将一个要素类的所有字段集复制到另一个要素类（除了shape字段与不可编辑的字段）
        /// </summary>
        /// <param name="source">源要素类</param>
        /// <param name="target">目标要素类</param>
        /// <returns>是否复制成功</returns>
        public static bool CopyFeatureClassFields(IFeatureClass source, IFeatureClass target)
        {
            IFields fs = source.Fields;
            for (int i = 0; i < fs.FieldCount; i++)
            {
                IField field = fs.get_Field(i);
                if ((!field.Editable) || field.Type == esriFieldType.esriFieldTypeGeometry || field.Name.ToLower() == "shape")
                {
                    continue;
                }
                else
                {
                    if (target.FindField(field.Name) == -1)
                    {
                        target.AddField(field);
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// 判断字段集中是否含有某些字段
        /// </summary>
        /// <param name="fields">字段集合</param>
        /// <param name="fieldName">字段名称</param>
        /// <returns></returns>
        public static bool[] isContainFields(IFields fields, params string[] fieldName)
        {

            bool[] isContain = new bool[fieldName.Length];
            for (int i = 0; i < fieldName.Length; i++)
            {
                if (fields.FindField(fieldName[i]) != -1)
                    isContain[i] = true;
                else
                    isContain[i] = false;
            }
            return isContain;
        }

        #region 导出要素类的字段值到excel
        //新建excel的类
         class ExcelOper
        {
            private Microsoft.Office.Interop.Excel.Application excelApp = null;
            public ExcelOper()
            {

            }

            private Microsoft.Office.Interop.Excel.Application GetExcelApplication()
            {
                Microsoft.Office.Interop.Excel.Application excelA = new Microsoft.Office.Interop.Excel.Application();
                excelA.Application.Workbooks.Add(true);
                excelApp = excelA;
                return excelA;
            }

            public bool WriteExcel(string excelPath, Hashtable htDT)
            {
                if (htDT == null || htDT.Count == 0)
                {
                    return false;
                }
                bool isWrite = false;
                //写入excel
                try
                {

                    if (excelApp == null)
                    {
                        GetExcelApplication();
                    }

                    //依次写入sheet表
                    int sheetNum = 1;
                    foreach (DictionaryEntry de in htDT)
                    {
                        string sheetName = de.Key.ToString();
                        System.Data.DataTable dTable = (System.Data.DataTable)de.Value;
                        Microsoft.Office.Interop.Excel.Worksheet pworkSheet = null;
                        if (sheetNum == 1)
                        {
                            pworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.Worksheets[sheetNum];
                            //Microsoft.Office.Tools.Excel.Worksheet
                        }
                        else
                        {
                            pworkSheet = excelApp.Worksheets.Add(Type.Missing, Type.Missing, 1, Type.Missing)
                                as Microsoft.Office.Interop.Excel.Worksheet;

                        }
                        pworkSheet.Name = sheetName;
                        bool sheetRs = WriteSheet(pworkSheet, dTable);
                        if (!sheetRs)
                        {
                            return false;
                        }

                        sheetNum++;
                    }

                    //保存Excel
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    excelApp.AlertBeforeOverwriting = false;
                    excelApp.ActiveWorkbook.SaveAs(excelPath, Type.Missing, null, null, false, false,
                                                      XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                    isWrite = true;

                }
                catch (Exception e)
                {
                    isWrite = false;
                    MessageBox.Show(e.StackTrace);
                }
                finally
                {
                    object missing = System.Reflection.Missing.Value;
                    //excelApp.ActiveWorkbook.Close(missing, missing, missing);
                    excelApp.Quit();
                    excelApp = null;
                    GC.Collect();
                }

                return isWrite;
            }

            public bool WriteSheet(Microsoft.Office.Interop.Excel.Worksheet pworkSheet, System.Data.DataTable dat)
            {
                if (pworkSheet == null || dat == null)
                {
                    return false;
                }

                for (int i = 0; i < dat.Columns.Count; i++)
                {
                    DataColumn dc = dat.Columns[i];
                    string caption = dc.Caption;
                    pworkSheet.Cells[1, i + 1] = caption;
                }

                for (int i = 0; i < dat.Rows.Count; i++)
                {
                    for (int j = 0; j < dat.Columns.Count; j++)
                    {
                        object ob = dat.Rows[i][j];
                        pworkSheet.Cells[2 + i, j + 1] = ob;
                    }

                }

                pworkSheet.Columns.AutoFit();
                return true;

            }
        }

        /// <summary>
        /// 将要素类的table表中的字段导出到excel
        /// </summary>
        /// <param name="featClass">要导出的要素类</param>
        /// <param name="excelPath">导出excel的路径</param>
        /// <param name="sheetName">写入excel的第一页的名称</param>
        /// <param name="fields2Export">要导出的字段名称，可变参量</param>
        /// <returns></returns>
        public static bool ExportFieldsValue2Excel(IFeatureClass featClass, string excelPath, string sheetName, params string[] fields2Export)
        {
            //bool isOK = false;

            //检查文件路径
            if (!excelPath.EndsWith(".xls") && !excelPath.EndsWith(".xlsx"))
            {
                MessageBox.Show("excel输入路径格式有误，请检查！");
                return false;
            }
            if (File.Exists(excelPath))
            {
                File.Delete(excelPath);
            }

            //读取要素类中的数据
            int[] index = new int[fields2Export.Length];
            for (int i = 0; i < fields2Export.Length; i++)
            {
                index[i] = featClass.Fields.FindField(fields2Export[i]);
            }
            if (!(System.Array.IndexOf(index, -1) == -1))
            {
                MessageBox.Show("shp文件中不存在所输入的字段，请检查参数是否正确！");
                return false;
            }

            //新建数据表,创建列
            System.Data.DataTable dt = new System.Data.DataTable();
            for (int i = 0; i < fields2Export.Length; i++)
            {
                DataColumn dc = new DataColumn();
                dc.ColumnName = fields2Export[i];
                dc.DataType = typeof(double);
                dt.Columns.Add(dc);
            }

            string value0 = null;

            IFeatureCursor ser1 = featClass.Search(null, false);
            IFeature fea = null;
            while ((fea = ser1.NextFeature()) != null)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < index.Length; j++)
                {
                    value0 = fea.get_Value(index[j]).ToString();
                    dr[fields2Export[j]] = Convert.ToDouble(value0);
                }

                dt.Rows.Add(dr);

            }
            //新建excel
            ExcelOper excelOpe = new ExcelOper();
            Hashtable hata = new Hashtable();
            hata.Add(sheetName, dt);

            if (excelOpe.WriteExcel(excelPath, hata))
            {
                return true;

            }
            else
            {
                MessageBox.Show("创建excel失败，请检查！");
                return false;
            }

        }

        #endregion
    }
}
