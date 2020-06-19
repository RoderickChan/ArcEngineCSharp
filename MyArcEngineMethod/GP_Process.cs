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
        /// 要素类转点工具
        /// </summary>
        /// <param name="inFeaClassPath">输入要素类路径</param>
        /// <param name="outFeaClassPath">输出要素类路径</param>
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
        /// 投影转换工具
        /// </summary>
        /// <param name="inFeatClassPath">输入要素类路径</param>
        /// <param name="outFeatClassPath">输出要素类路径</param>
        /// <param name="projectPath">投影文件路径，投影文件如：CGCS2000 3 Degree GK CM 117E.prj</param>
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
      /// 生成Voronoi图
      /// </summary>
      /// <param name="inFeaClassPath">输入要素类路径</param>
      /// <param name="outFeaClassPath">输出Voronoi要素类路径</param>
        public static void CreateVoronoi(string inFeaClassPath, string outFeaClassPath)
        {
            if (!File.Exists(inFeaClassPath))
                return;

            if (!inFeaClassPath.EndsWith(".shp"))
                return;

            
            //若存在shp文件，则将其删除
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
        /// 复制要素类
        /// </summary>
        /// <param name="sourcePath">待复制要素类路径</param>
        /// <param name="targetPath">输出要素类路径</param>
        public static void CopyFeatureClass(string sourcePath, string targetPath)
        {
            if (!File.Exists(sourcePath))
                return;
            if (!sourcePath.EndsWith(".shp"))
                return;

            string path_dir = System.IO.Path.GetDirectoryName(targetPath);
            if (File.Exists(targetPath)) //存在该路径，则在文件夹中找到文件并删除
            {
                string fileName = System.IO.Path.GetFileNameWithoutExtension(targetPath); //文件名称
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

            //拷贝要素类
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
        /// 用点分割线要素
        /// </summary>
        /// <param name="lineFeaClassPath">线要素路径</param>
        /// <param name="pointFeaClassPath">点要素路径</param>
        /// <param name="outFeaClass">输出线要素路径</param>
        /// <param name="searchRadius">设置容差值，默认为"1 Meters"</param>
        public static void SplitPolylineByPointFeaClass(string lineFeaClassPath, string pointFeaClassPath, string outFeaClass, string searchRadius = "1 Meters")
        {
            if (!File.Exists(lineFeaClassPath) || !File.Exists(pointFeaClassPath))
                return;
            //删除同名文件
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
        /// 创建渔网要素
        /// </summary>
        /// <param name="shpPath">模板要素类的路径</param>
        /// <param name="savePath">结果存储路径</param>
        /// <param name="columsNum">列数</param>
        /// <param name="rowsNum">行数</param>
        /// <param name="geometryType">几何类型，只能是"polyline"或"polygon"</param>
        [UnFinished("chenhuan","2020-03-07","该方法不可使用！")]
        public static void CreateFishnetFeaClass(string shpPath, string savePath,int columsNum, int rowsNum, string geometryType)
        {
            Geoprocessor gp = new Geoprocessor();
            IFeatureClass fc = MyArcEngineMethod.FeatureClassOperation.OpenFeatureClass(shpPath);
            CreateFishnet cf = new CreateFishnet();
            cf.origin_coord = null;
            
            cf.number_columns = columsNum; //列数
            cf.number_rows = rowsNum;//行数
            cf.out_feature_class = savePath;
            cf.out_label = false; //不显示标注
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
