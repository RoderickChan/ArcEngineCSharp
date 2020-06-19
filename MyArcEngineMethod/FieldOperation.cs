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
        /// �����ֶ���ΪҪ���ഴ���ֶ�
        /// </summary>
        /// <param name="featClass">����Ҫ����</param>
        /// <param name="fieldType">�ֶ�����</param>
        /// <param name="fieldsname">�ֶ����ƣ��ɱ����</param>
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
        /// �����ֶβ��ҵõ��ֶε�����
        /// </summary>
        /// <param name="featClass">������Ҫ������ֶ�</param>
        /// <param name="fieldType">�ֶ�����</param>
        /// <param name="fieldsname">�ֶ�����</param>
        /// <returns>�����½��ֶε������������ڣ���������е�����</returns>
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
        /// ����һ���ֶβ���������
        /// </summary>
        /// <param name="featClass">Ҫ����</param>
        /// <param name="fieldType">�ֶ�����</param>
        /// <param name="fieldsname">�ֶ�����</param>
        /// <returns>�ֶε�����</returns>
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
        /// ��Ҫ������ֶε�ֵ������Ϊ��
        /// </summary>
        /// <param name="feaClass">Ҫ����</param>
        /// <param name="fieldNames">�ֶ�����</param>
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
        /// ɾ��Ҫ�����в���Ҫ���ֶ�
        /// </summary>
        /// <param name="feaClass">�����Ҫ����</param>
        /// <param name="fieldName">ɾ�����ֶ���,�ɱ����</param>
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
        /// ��ȡ��Ҫ�����ĳһ�ֶ�ö��ֵ�������ַ������͵��б�
        /// </summary>
        /// <param name="fc">�����Ҫ����</param>
        /// <param name="fieldName">����ѯ���ֶ�����</param>
        /// <returns></returns>
        [ObsoleteAttribute("���Ƽ���������������Ƽ�ʹ�������Table")]
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
        /// ��ȡ�����Ա��ĳ�ֶ�Ψһֵ
        /// </summary>
        /// <param name="table">���Ա�</param>
        /// <param name="fieldName">�ֶ���</param>
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
        /// ����Ҫ�ص��ֶ�����
        /// </summary>
        /// <param name="sourceFea">ԴҪ��</param>
        /// <param name="targetFea">������Ҫ��</param>
        /// <param name="geo">������״����Ϊ��</param>
        /// <param name="isStore">�Ƿ����fea��store����</param>
        /// <returns></returns>
        public static bool CopyFeatureFields(IFeature sourceFea, IFeature targetFea, IGeometry geo, bool isStore)
        {
            if (sourceFea == null || targetFea == null)
                return false;

            IFields copyFields = sourceFea.Fields;
            for (int i = 0; i < copyFields.FieldCount; i++)
            {
                IField curField = copyFields.get_Field(i);
                //�ڴ�����Ҫ����Ѱ�Ҹ��ֶ�
                int indexPaste = targetFea.Fields.FindField(curField.Name);
                if (indexPaste != -1)
                {
                    IField pasteField = targetFea.Fields.get_Field(indexPaste);
                    //���ֶοɱ༭�������ֶ�ֵ
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
        /// ��һ��Ҫ����������ֶμ����Ƶ���һ��Ҫ���ࣨ����shape�ֶ��벻�ɱ༭���ֶΣ�
        /// </summary>
        /// <param name="source">ԴҪ����</param>
        /// <param name="target">Ŀ��Ҫ����</param>
        /// <returns>�Ƿ��Ƴɹ�</returns>
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
        /// �ж��ֶμ����Ƿ���ĳЩ�ֶ�
        /// </summary>
        /// <param name="fields">�ֶμ���</param>
        /// <param name="fieldName">�ֶ�����</param>
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

        #region ����Ҫ������ֶ�ֵ��excel
        //�½�excel����
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
                //д��excel
                try
                {

                    if (excelApp == null)
                    {
                        GetExcelApplication();
                    }

                    //����д��sheet��
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

                    //����Excel
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
        /// ��Ҫ�����table���е��ֶε�����excel
        /// </summary>
        /// <param name="featClass">Ҫ������Ҫ����</param>
        /// <param name="excelPath">����excel��·��</param>
        /// <param name="sheetName">д��excel�ĵ�һҳ������</param>
        /// <param name="fields2Export">Ҫ�������ֶ����ƣ��ɱ����</param>
        /// <returns></returns>
        public static bool ExportFieldsValue2Excel(IFeatureClass featClass, string excelPath, string sheetName, params string[] fields2Export)
        {
            //bool isOK = false;

            //����ļ�·��
            if (!excelPath.EndsWith(".xls") && !excelPath.EndsWith(".xlsx"))
            {
                MessageBox.Show("excel����·����ʽ�������飡");
                return false;
            }
            if (File.Exists(excelPath))
            {
                File.Delete(excelPath);
            }

            //��ȡҪ�����е�����
            int[] index = new int[fields2Export.Length];
            for (int i = 0; i < fields2Export.Length; i++)
            {
                index[i] = featClass.Fields.FindField(fields2Export[i]);
            }
            if (!(System.Array.IndexOf(index, -1) == -1))
            {
                MessageBox.Show("shp�ļ��в�������������ֶΣ���������Ƿ���ȷ��");
                return false;
            }

            //�½����ݱ�,������
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
            //�½�excel
            ExcelOper excelOpe = new ExcelOper();
            Hashtable hata = new Hashtable();
            hata.Add(sheetName, dt);

            if (excelOpe.WriteExcel(excelPath, hata))
            {
                return true;

            }
            else
            {
                MessageBox.Show("����excelʧ�ܣ����飡");
                return false;
            }

        }

        #endregion
    }
}
