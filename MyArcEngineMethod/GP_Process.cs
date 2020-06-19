using System;
using System.IO;
using System.Runtime.InteropServices;
using MyArcEngineMethod.MyAttribute;
using ESRI.ArcGIS.AnalysisTools;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.Geoprocessor;

namespace MyArcEngineMethod
{
    [Guid("94dd641a-4690-47bf-940b-34ee219de29c")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyArcEngineMethod.GP_Process")]
    public class GP_Process
    {       
        /// <summary>
        /// Ҫ����ת�㹤��
        /// </summary>
        /// <param name="inFeaClassPath">����Ҫ����·��</param>
        /// <param name="outFeaClassPath">���Ҫ����·��</param>
        public static void FeaturesToPoint(string inFeaClassPath, string outFeaClassPath)
        {
            if (!File.Exists(inFeaClassPath))
                return;
      
            if(!inFeaClassPath.EndsWith(".shp"))
                return;

            if (File.Exists(outFeaClassPath))
            {
                string nname = System.IO.Path.GetFileNameWithoutExtension(outFeaClassPath);
                DirectoryInfo root = new System.IO.DirectoryInfo(System.IO.Path.GetDirectoryName(outFeaClassPath));
                foreach (FileInfo f in root.GetFiles())
                {
                    if (f.Name.Split('.')[0] == nname)
                    {
                        string fullPath = f.FullName;
                        File.Delete(fullPath);
                    }
                }
            }

            Geoprocessor gp = new Geoprocessor();
            FeatureToPoint ftp = new FeatureToPoint();
            ftp.in_features = inFeaClassPath;
            ftp.out_feature_class = outFeaClassPath;
            try
            {
                gp.Execute(ftp, null);
            }
            catch (Exception e)
            { throw e; }
        }

        /// <summary>
        /// ͶӰת������
        /// </summary>
        /// <param name="inFeatClassPath">����Ҫ����·��</param>
        /// <param name="outFeatClassPath">���Ҫ����·��</param>
        /// <param name="projectPath">ͶӰ�ļ�·����ͶӰ�ļ��磺CGCS2000 3 Degree GK CM 117E.prj</param>
        public static void Projection(string inFeatClassPath, string outFeatClassPath, string projectPath)
        {
            if (!File.Exists(inFeatClassPath) || !File.Exists(projectPath))            
                return;
            
            if (!inFeatClassPath.EndsWith(".shp") || !projectPath.EndsWith(".prj"))
                return;

            if (File.Exists(outFeatClassPath))
            {
                string nname = System.IO.Path.GetFileNameWithoutExtension(outFeatClassPath);
                DirectoryInfo root = new System.IO.DirectoryInfo(System.IO.Path.GetDirectoryName(outFeatClassPath));
                foreach (FileInfo f in root.GetFiles())
                {
                    if (f.Name.Split('.')[0] == nname)
                    {
                        string fullPath = f.FullName;
                        File.Delete(fullPath);
                    }
                }
            }

            Geoprocessor gp = new Geoprocessor();
            Project ppr = new Project();
            ppr.in_dataset = inFeatClassPath;
            ppr.out_dataset = outFeatClassPath;
            ppr.out_coor_system = projectPath;
            try
            { gp.Execute(ppr, null); }
            catch (Exception e) { throw e; }

        }

      /// <summary>
      /// ����Voronoiͼ
      /// </summary>
      /// <param name="inFeaClassPath">����Ҫ����·��</param>
      /// <param name="outFeaClassPath">���VoronoiҪ����·��</param>
        public static void CreateVoronoi(string inFeaClassPath, string outFeaClassPath)
        {
            if (!File.Exists(inFeaClassPath))
                return;

            if (!inFeaClassPath.EndsWith(".shp"))
                return;

            
            //������shp�ļ�������ɾ��
            if (File.Exists(outFeaClassPath))
            {
                string nname = System.IO.Path.GetFileNameWithoutExtension(outFeaClassPath);
                DirectoryInfo root = new System.IO.DirectoryInfo(System.IO.Path.GetDirectoryName(outFeaClassPath));
                foreach (FileInfo f in root.GetFiles())
                {
                    if (f.Name.Split('.')[0] == nname)
                    {
                        string fullPath = f.FullName;
                        File.Delete(fullPath);
                    }
                }
            }

            Geoprocessor gp = new Geoprocessor();
            CreateThiessenPolygons ctp = new CreateThiessenPolygons();
            ctp.in_features = inFeaClassPath;
            ctp.fields_to_copy = "ONLY_FID";
            ctp.out_feature_class = outFeaClassPath;
            try
            {
                gp.Execute(ctp, null);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        /// <summary>
        /// ����Ҫ����
        /// </summary>
        /// <param name="sourcePath">������Ҫ����·��</param>
        /// <param name="targetPath">���Ҫ����·��</param>
        public static void CopyFeatureClass(string sourcePath, string targetPath)
        {
            if (!File.Exists(sourcePath))
                return;
            if (!sourcePath.EndsWith(".shp"))
                return;

            string path_dir = System.IO.Path.GetDirectoryName(targetPath);
            if (File.Exists(targetPath)) //���ڸ�·���������ļ������ҵ��ļ���ɾ��
            {
                string fileName = System.IO.Path.GetFileNameWithoutExtension(targetPath); //�ļ�����
                DirectoryInfo root = new DirectoryInfo(path_dir);
                foreach (FileInfo f in root.GetFiles())
                {
                    if (f.Name.Split('.')[0] == fileName)
                    {
                        string filepath = f.FullName;
                        File.Delete(filepath);
                    }
                }
            }

            //����Ҫ����
            IGeoProcessor2 gp = new GeoProcessorClass();
            gp.OverwriteOutput = true;
            IGeoProcessorResult result = new GeoProcessorResultClass();
            IVariantArray parameters = new VarArrayClass();
            try
            {
                parameters.Add(sourcePath);
                parameters.Add(targetPath);
                result = gp.Execute("Copy_management", parameters, null);

            }
            catch (Exception ex)
            {
                throw ex;

            }
        }

        /// <summary>
        /// �õ�ָ���Ҫ��
        /// </summary>
        /// <param name="lineFeaClassPath">��Ҫ��·��</param>
        /// <param name="pointFeaClassPath">��Ҫ��·��</param>
        /// <param name="outFeaClass">�����Ҫ��·��</param>
        /// <param name="searchRadius">�����ݲ�ֵ��Ĭ��Ϊ"1 Meters"</param>
        public static void SplitPolylineByPointFeaClass(string lineFeaClassPath, string pointFeaClassPath, string outFeaClass, string searchRadius = "1 Meters")
        {
            if (!File.Exists(lineFeaClassPath) || !File.Exists(pointFeaClassPath))
                return;
            //ɾ��ͬ���ļ�
            if (File.Exists(outFeaClass))
            {
                string nname = System.IO.Path.GetFileNameWithoutExtension(outFeaClass);
                DirectoryInfo root = new System.IO.DirectoryInfo(System.IO.Path.GetDirectoryName(outFeaClass));
                foreach (FileInfo f in root.GetFiles())
                {
                    if (f.Name.Split('.')[0] == nname)
                    {
                        string fullPath = f.FullName;
                        File.Delete(fullPath);
                    }
                }
            }

            Geoprocessor gp = new Geoprocessor();
            SplitLineAtPoint slap = new SplitLineAtPoint();
            slap.in_features = lineFeaClassPath;
            slap.out_feature_class = outFeaClass;
            slap.point_features = pointFeaClassPath;
            slap.search_radius = searchRadius;
            try
            {
                gp.Execute(slap, null);
            }
            catch (Exception e)
            { throw e; }
        }


        /// <summary>
        /// ��������Ҫ��
        /// </summary>
        /// <param name="shpPath">ģ��Ҫ�����·��</param>
        /// <param name="savePath">����洢·��</param>
        /// <param name="columsNum">����</param>
        /// <param name="rowsNum">����</param>
        /// <param name="geometryType">�������ͣ�ֻ����"polyline"��"polygon"</param>
        [UnFinished("chenhuan","2020-03-07","�÷�������ʹ�ã�")]
        public static void CreateFishnetFeaClass(string shpPath, string savePath,int columsNum, int rowsNum, string geometryType)
        {
            Geoprocessor gp = new Geoprocessor();
            IFeatureClass fc = MyArcEngineMethod.FeatureClassOperation.OpenFeatureClass(shpPath);
            CreateFishnet cf = new CreateFishnet();
            cf.origin_coord = null;
            
            cf.number_columns = columsNum; //����
            cf.number_rows = rowsNum;//����
            cf.out_feature_class = savePath;
            cf.out_label = false; //����ʾ��ע
            cf.geometry_type = geometryType;
            cf.template = (fc as IGeoDataset);
            //cf.
            try
            {
                gp.Execute(cf, null);

            }
            catch 
            {
                throw;
            }
        }

    }
}
