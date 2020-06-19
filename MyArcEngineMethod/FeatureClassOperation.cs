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
    [Guid("85513b51-3a35-4ffe-8020-a03bb9950388")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyArcEngineMethod.Class1")]
    public class FeatureClassOperation
    {
        /// <summary>
        /// ��Ҫ����
        /// </summary>
        /// <param name="shpFilePath">�ļ�λ��</param>
        public static IFeatureClass OpenFeatureClass(string shpFilePath)
        {
            if (!File.Exists(shpFilePath)||!shpFilePath.EndsWith(".shp"))            
                return null;
           
            IWorkspace workspace;
            IFeatureWorkspace featureWorkspace;
            IFeatureClass featureClass=null;

            IWorkspaceFactory workspaceFactory = new ShapefileWorkspaceFactoryClass();
            //�ر�ͼ������
            IWorkspaceFactoryLockControl iwsflc = workspaceFactory as IWorkspaceFactoryLockControl;
            if (iwsflc.SchemaLockingEnabled)
                iwsflc.DisableSchemaLocking();

            workspace = workspaceFactory.OpenFromFile(System.IO.Path.GetDirectoryName(shpFilePath), 0);
            featureWorkspace = workspace as IFeatureWorkspace;
            string filename = System.IO.Path.GetFileNameWithoutExtension(shpFilePath);

            featureClass = featureWorkspace.OpenFeatureClass(filename);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(featureWorkspace);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workspaceFactory);                 
            
            return featureClass;          
       
        }

        /// <summary>
        /// ������Ҫ����
        /// </summary>
        /// <param name="shpPath">ָ�����Ҫ�����·��</param>
        /// <param name="geoType">��Ҫ����ļ�������</param>
        /// <param name="sr">��Ҫ����Ŀռ�ο�</param>
        /// <returns></returns>
        public static IFeatureClass CreateNewFeatureClass(string shpPath, esriGeometryType geoType, ISpatialReference sr)
        {
            if(!shpPath.EndsWith(".shp"))
                return null;

            string shpName = System.IO.Path.GetFileNameWithoutExtension(shpPath);
            string filenam = System.IO.Path.GetDirectoryName(shpPath);
            //ɾ��ͬ���ļ�
            DirectoryInfo root = new DirectoryInfo(filenam);
            foreach (FileInfo f in root.GetFiles())
            {
                if (f.Name.Split('.')[0] == shpName)
                {
                    string filepath = f.FullName;
                    try
                    {
                        File.Delete(filepath);
                    }
                    catch
                    {
 
                    }
                }
            }

            IWorkspaceFactory iwSF = new ShapefileWorkspaceFactory();
            //�ر�ͼ������
            IWorkspaceFactoryLockControl iwsflc = iwSF as IWorkspaceFactoryLockControl;
            if (iwsflc.SchemaLockingEnabled)
                iwsflc.DisableSchemaLocking();

            IFeatureWorkspace fWor = iwSF.OpenFromFile(filenam, 0) as IFeatureWorkspace;

            IFields pFields = new FieldsClass();
            IFieldsEdit pFieldsEdit = (IFieldsEdit)pFields;

            IField mField = new FieldClass();
            IFieldEdit mFieldEdit = mField as IFieldEdit;

            mFieldEdit.Name_2 = "Shape";
            mFieldEdit.Type_2 = esriFieldType.esriFieldTypeGeometry;
            IGeometryDef pGeoDef = new GeometryDefClass();
            IGeometryDefEdit pGeDEdit = pGeoDef as IGeometryDefEdit;
            pGeDEdit.GeometryType_2 = geoType;
            //��������ϵ
            pGeDEdit.SpatialReference_2 = sr;
            mFieldEdit.GeometryDef_2 = pGeoDef;
            pFieldsEdit.AddField(mField);

            IFeatureClass newFeatClass = fWor.CreateFeatureClass(shpName, pFields, null, null, esriFeatureType.esriFTSimple, "Shape", "");

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iwSF);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(fWor);

            return newFeatClass;

        }

        /// <summary>
        /// ������Ҫ����,������ԴҪ����������ֶ���ռ�ο�
        /// </summary>
        /// <param name="shpPath">ָ�����Ҫ�����·��</param>
        /// <param name="geoType">��Ҫ����ļ�������</param>
        /// <param name="sourceClass">ԴҪ����</param>
        /// <param name="copyRs">�Ƿ��ƿռ�ο�, ��Ϊ����������sr</param>
        /// <returns>��Ҫ���࣬�����κ�Ҫ��</returns>
        public static IFeatureClass CreateNewFeatureClass(string shpPath, esriGeometryType geoType, IFeatureClass sourceClass, bool copyRs, ISpatialReference sr = null)
        {
            if (copyRs)
            {
                sr = (sourceClass as IGeoDataset).SpatialReference;
            }
            IFeatureClass fc = CreateNewFeatureClass(shpPath, geoType, sr);
            //����ֶ�
            IFields allFields = sourceClass.Fields;
            for (int i = 0; i < allFields.FieldCount; i++)
            {
                IField curField = allFields.get_Field(i);
                if ((!curField.Editable) || (curField.Type == esriFieldType.esriFieldTypeGeometry)
                    || fc.FindField(curField.Name) != -1)
                    continue;
                fc.AddField(curField);
            }
            return fc;

        }

        /// <summary>
        /// �ж�Ҫ���������ϵ���ͣ������귵��0��ͶӰ����ϵ����1����������ϵ����-1
        /// </summary>
        /// <param name="feaClass"></param>
        /// <returns></returns>
        public static int JudgeFeatureClassCoordinateSystem(IFeatureClass feaClass)
        {
            IGeoDataset dataSet = feaClass as IGeoDataset;
            if(dataSet==null)
                return 0;

            ISpatialReference sR = dataSet.SpatialReference;
            if (sR == null)
                return 0;
            else if ((sR as IProjectedCoordinateSystem) != null)

                return 1;
            else
                return -1;
        }

        /// <summary>
        /// ����TIN
        /// </summary>
        /// <param name="pointColl">�㼯</param>
        /// <param name="featureClass">�㼯���ڵ�Ҫ����</param>
        /// <returns></returns>
       public static TinClass QueryTIN(IPointCollection pointColl, IFeatureClass featureClass)
        {
            TinClass tin = new TinClass();//����������������
            IGeoDataset pGeoDataset = featureClass as IGeoDataset;
            IEnvelope envelope = pGeoDataset.Extent;//������ݼ�Ŀ��ֲ��ķ�Χ��Ϣ
            // ����TIN
            tin.InitNew(envelope);//�����ʼ��
            for (int i = 0; i < pointColl.PointCount - 1; i++)
            {
                IPoint point = pointColl.get_Point(i);
                point.Z = 0;//Zֵ��0
                int tagValue = i;//tagValue�洢�õ��ڶ����IPointCollection���±�
                ITinNode node = new TinNodeClass();
                tin.AddPointZ(point, tagValue, node);//���ν�����뵽TIN��
            }
            return tin;
        }

        /// <summary>
        /// ������Ҫ�����������
        /// </summary>
        /// <param name="pointFeaClass">����ĵ�Ҫ����</param>
        /// <returns></returns>
        public static TinClass QueryTIN(IFeatureClass pointFeaClass)
        {
            IGeoDataset pGeoDataset = pointFeaClass as IGeoDataset;
            IEnvelope envelope = pGeoDataset.Extent;//������ݼ�Ŀ��ֲ��ķ�Χ��Ϣ

            if (pointFeaClass.ShapeType != esriGeometryType.esriGeometryPoint)
                return null;

            TinClass tin = new TinClass();//����������������
            // ����TIN
            tin.InitNew(envelope);//�����ʼ��

            IFeatureCursor cur1 = pointFeaClass.Search(null, false);
            IFeature mFea = null;
            //int i = 0;
            while ((mFea = cur1.NextFeature()) != null)
            {
                IPoint point = mFea.ShapeCopy as IPoint;
                point.Z = 0;//Zֵ��0
                (point as ITopologicalOperator).Simplify();
                int tagValue = mFea.OID;//tagValue�洢�õ���Ҫ�����FID���±�
                ITinNode node = new TinNodeClass();
                tin.AddPointZ(point, tagValue, node);//���ν�����뵽TIN
            }
            Marshal.FinalReleaseComObject(cur1);
            return tin;
        }

        /// <summary>
        /// ����״��װ��Ҫ�ز������Ҫ����
        /// </summary>
        /// <param name="featureClass">Ҫ����</param>
        /// <param name="geometry">�����</param>
        /// <param name="copyFea">�ṩ��Ҫ�ظ������Ե�ԭҪ��</param>
        public static IFeature AddGeometryToFeatureClass_Store(IFeatureClass featureClass, IGeometry geometry, IFeature copyFea)
        {
            // �½�Ҫ��
            IFeature feature2Insert = featureClass.CreateFeature();
            feature2Insert.Shape = geometry;

            IFields copyFields = copyFea.Fields;
            for (int i = 0; i < copyFields.FieldCount; i++)
            {
                IField curField = copyFields.get_Field(i);
                //�ڴ�����Ҫ����Ѱ�Ҹ��ֶ�
                int indexPaste = feature2Insert.Fields.FindField(curField.Name);
                if (indexPaste != -1)
                {
                    IField pasteField = feature2Insert.Fields.get_Field(indexPaste);
                    //���ֶοɱ༭�������ֶ�ֵ
                    if (pasteField.Editable && pasteField.Type != esriFieldType.esriFieldTypeGeometry)
                    {
                        feature2Insert.Value[indexPaste] = copyFea.Value[i];
                    }
                }
            }
            
            feature2Insert.Store();

            return feature2Insert;
        }

        /// <summary>
        /// ʹ���α�ķ�����Ҫ�����в���Ҫ��
        /// </summary>
        /// <param name="fc">Ҫ�������ݿ�</param>
        /// <param name="geo">��Ҫ�صļ�����״</param>
        /// <param name="copyFea">�����ֶ����Ե�Ҫ��</param>
        /// <param name="isGetFeature">�Ƿ񷵻�Ҫ�أ�trueʱ�õ���Ҫ�أ�falseʱ���ؿ�ֵ</param>
        /// <returns></returns>
        public static IFeature AddGeometryToFeatureClass_Cursor(IFeatureClass fc, IGeometry geo, IFeature copyFea,bool isGetFeature)
        {
            IFeatureCursor insertCursor = fc.Insert(true);
            IFeatureBuffer featBuf = fc.CreateFeatureBuffer();
            featBuf.Shape = geo;

            IFields copyFields = copyFea.Fields;
            for (int i = 0; i < copyFields.FieldCount; i++)
            {
                IField curField = copyFields.get_Field(i);
                //�ڴ�����Ҫ����Ѱ�Ҹ��ֶ�
                int indexPaste = featBuf.Fields.FindField(curField.Name);
                if (indexPaste != -1)
                {
                    IField pasteField = featBuf.Fields.get_Field(indexPaste);
                    //���ֶοɱ༭�������ֶ�ֵ
                    if (pasteField.Editable)
                    {
                        featBuf.Value[indexPaste] = copyFea.Value[i];
                    }
                }
            }
           object ob1=insertCursor.InsertFeature(featBuf);
            insertCursor.Flush();//д���ڴ�
            Marshal.FinalReleaseComObject(insertCursor);

            if (isGetFeature)
            {
                int feaId = Convert.ToInt32(ob1);
                IFeature curFea = fc.GetFeature(feaId);
                return curFea;
            }
            else
                return null;
        }

    }
}
