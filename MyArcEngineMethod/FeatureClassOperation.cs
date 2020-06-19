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
        /// 打开要素类
        /// </summary>
        /// <param name="shpFilePath">文件位置</param>
        public static IFeatureClass OpenFeatureClass(string shpFilePath)
        {
            if (!File.Exists(shpFilePath)||!shpFilePath.EndsWith(".shp"))            
                return null;
           
            IWorkspace workspace;
            IFeatureWorkspace featureWorkspace;
            IFeatureClass featureClass=null;

            IWorkspaceFactory workspaceFactory = new ShapefileWorkspaceFactoryClass();
            //关闭图层锁定
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
        /// 创建新要素类
        /// </summary>
        /// <param name="shpPath">指定输出要素类的路径</param>
        /// <param name="geoType">新要素类的几何类型</param>
        /// <param name="sr">新要素类的空间参考</param>
        /// <returns></returns>
        public static IFeatureClass CreateNewFeatureClass(string shpPath, esriGeometryType geoType, ISpatialReference sr)
        {
            if(!shpPath.EndsWith(".shp"))
                return null;

            string shpName = System.IO.Path.GetFileNameWithoutExtension(shpPath);
            string filenam = System.IO.Path.GetDirectoryName(shpPath);
            //删掉同名文件
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
            //关闭图层锁定
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
            //定义坐标系
            pGeDEdit.SpatialReference_2 = sr;
            mFieldEdit.GeometryDef_2 = pGeoDef;
            pFieldsEdit.AddField(mField);

            IFeatureClass newFeatClass = fWor.CreateFeatureClass(shpName, pFields, null, null, esriFeatureType.esriFTSimple, "Shape", "");

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iwSF);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(fWor);

            return newFeatClass;

        }

        /// <summary>
        /// 创建新要素类,并复制源要素类的所有字段与空间参考
        /// </summary>
        /// <param name="shpPath">指定输出要素类的路径</param>
        /// <param name="geoType">新要素类的几何类型</param>
        /// <param name="sourceClass">源要素类</param>
        /// <param name="copyRs">是否复制空间参考, 若为否，则必须给出sr</param>
        /// <returns>新要素类，不含任何要素</returns>
        public static IFeatureClass CreateNewFeatureClass(string shpPath, esriGeometryType geoType, IFeatureClass sourceClass, bool copyRs, ISpatialReference sr = null)
        {
            if (copyRs)
            {
                sr = (sourceClass as IGeoDataset).SpatialReference;
            }
            IFeatureClass fc = CreateNewFeatureClass(shpPath, geoType, sr);
            //添加字段
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
        /// 判断要素类的坐标系类型，无坐标返回0，投影坐标系返回1，地理坐标系返回-1
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
        /// 创建TIN
        /// </summary>
        /// <param name="pointColl">点集</param>
        /// <param name="featureClass">点集所在的要素类</param>
        /// <returns></returns>
       public static TinClass QueryTIN(IPointCollection pointColl, IFeatureClass featureClass)
        {
            TinClass tin = new TinClass();//创建的三角网对象
            IGeoDataset pGeoDataset = featureClass as IGeoDataset;
            IEnvelope envelope = pGeoDataset.Extent;//获得数据集目标分布的范围信息
            // 构建TIN
            tin.InitNew(envelope);//对象初始化
            for (int i = 0; i < pointColl.PointCount - 1; i++)
            {
                IPoint point = pointColl.get_Point(i);
                point.Z = 0;//Z值赋0
                int tagValue = i;//tagValue存储该点在多边形IPointCollection的下标
                ITinNode node = new TinNodeClass();
                tin.AddPointZ(point, tagValue, node);//依次将点加入到TIN中
            }
            return tin;
        }

        /// <summary>
        /// 构建点要素类的三角网
        /// </summary>
        /// <param name="pointFeaClass">输入的点要素类</param>
        /// <returns></returns>
        public static TinClass QueryTIN(IFeatureClass pointFeaClass)
        {
            IGeoDataset pGeoDataset = pointFeaClass as IGeoDataset;
            IEnvelope envelope = pGeoDataset.Extent;//获得数据集目标分布的范围信息

            if (pointFeaClass.ShapeType != esriGeometryType.esriGeometryPoint)
                return null;

            TinClass tin = new TinClass();//创建的三角网对象
            // 构建TIN
            tin.InitNew(envelope);//对象初始化

            IFeatureCursor cur1 = pointFeaClass.Search(null, false);
            IFeature mFea = null;
            //int i = 0;
            while ((mFea = cur1.NextFeature()) != null)
            {
                IPoint point = mFea.ShapeCopy as IPoint;
                point.Z = 0;//Z值赋0
                (point as ITopologicalOperator).Simplify();
                int tagValue = mFea.OID;//tagValue存储该点在要素类的FID的下标
                ITinNode node = new TinNodeClass();
                tin.AddPointZ(point, tagValue, node);//依次将点加入到TIN
            }
            Marshal.FinalReleaseComObject(cur1);
            return tin;
        }

        /// <summary>
        /// 将形状包装成要素并添加至要素类
        /// </summary>
        /// <param name="featureClass">要素类</param>
        /// <param name="geometry">多边形</param>
        /// <param name="copyFea">提供新要素各种属性的原要素</param>
        public static IFeature AddGeometryToFeatureClass_Store(IFeatureClass featureClass, IGeometry geometry, IFeature copyFea)
        {
            // 新建要素
            IFeature feature2Insert = featureClass.CreateFeature();
            feature2Insert.Shape = geometry;

            IFields copyFields = copyFea.Fields;
            for (int i = 0; i < copyFields.FieldCount; i++)
            {
                IField curField = copyFields.get_Field(i);
                //在带插入要素中寻找该字段
                int indexPaste = feature2Insert.Fields.FindField(curField.Name);
                if (indexPaste != -1)
                {
                    IField pasteField = feature2Insert.Fields.get_Field(indexPaste);
                    //若字段可编辑，则复制字段值
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
        /// 使用游标的方法向要素类中插入要素
        /// </summary>
        /// <param name="fc">要素类数据库</param>
        /// <param name="geo">新要素的几何形状</param>
        /// <param name="copyFea">复制字段属性的要素</param>
        /// <param name="isGetFeature">是否返回要素，true时得到新要素，false时返回空值</param>
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
                //在带插入要素中寻找该字段
                int indexPaste = featBuf.Fields.FindField(curField.Name);
                if (indexPaste != -1)
                {
                    IField pasteField = featBuf.Fields.get_Field(indexPaste);
                    //若字段可编辑，则复制字段值
                    if (pasteField.Editable)
                    {
                        featBuf.Value[indexPaste] = copyFea.Value[i];
                    }
                }
            }
           object ob1=insertCursor.InsertFeature(featBuf);
            insertCursor.Flush();//写入内存
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
