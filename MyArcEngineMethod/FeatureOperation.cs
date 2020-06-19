using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;
using System.Windows.Forms;
using ESRI.ArcGIS.AnalysisTools;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geoprocessor;
using ESRI.ArcGIS.Carto;
using MyArcEngineMethod.MyAttribute;

namespace MyArcEngineMethod
{
    [Guid("0b6b97d4-9b62-4311-9ff8-ecca36ecb458")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyArcEngineMethod.FeatureOperation")]
    [Author("Lyn", "1.0", "2020-03-31")]
    public class FeatureOperation
    {
        /// <summary>
        /// 获取到当前地图上的所有选择的要素
        /// </summary>
        /// <param name="mMap">当前地图 IMap</param>
        /// <returns></returns>   
        public static IFeature[]  GetAllSelectedFeaturesArray(IMap mMap)
        {
            if (mMap == null || mMap.SelectionCount == 0)
                return null;

            IFeature[] feaArray = new IFeature[mMap.SelectionCount];
            ISelection selection = mMap.FeatureSelection;            
            IEnumFeatureSetup feaSetup = selection as IEnumFeatureSetup;
            feaSetup.AllFields = true;
            IEnumFeature eFea = feaSetup as IEnumFeature;
            eFea.Reset();
            IFeature mFea = null;
            int index = 0;
            while ((mFea = eFea.Next()) != null)
            {
                feaArray[index++] = mFea;
            }
            return feaArray;
        }

        /// <summary>
        /// 获取到当前地图上的所有选择的要素
        /// </summary>
        /// <param name="mMap">当前地图 IMap</param>
        /// <returns></returns>
        public static List<IFeature> GetAllSelectedFeaturesList(IMap mMap)
        {
            List<IFeature> feaList = new List<IFeature>();

            var allFeas = GetAllSelectedFeaturesArray(mMap);
            if (allFeas != null)
                feaList.AddRange(allFeas);
            return feaList;
        }

        /// <summary>
        /// 获取在当前地图上选中了的属于某一要素类的所有要素
        /// </summary>
        /// <param name="mMap">当前地图IMap</param>
        /// <param name="objectClassID">某一要素类的ObjectClassID</param>
        /// <returns></returns>
        public static List<IFeature> GetSelectedFeaListInCurrentFeaclass(IMap mMap, int objectClassID)
        {           
            var allFeas = GetAllSelectedFeaturesArray(mMap);
            if (allFeas == null)
                return null;
            List<IFeature> feaList = new List<IFeature>();
            for (int i = 0; i < allFeas.Length; i++)
            {
                if (allFeas[i].Class.ObjectClassID == objectClassID)
                    feaList.Add(allFeas[i]);
            }
            return feaList;
        }

        /// <summary>
        /// 获取在当前地图上选中了的属于某一要素类的所有要素
        /// </summary>
        /// <param name="mMap">当前地图IMap</param>
        /// <param name="objectClassID">某一要素类的ObjectClassID</param>
        /// <returns></returns>
        public static IFeature[] GetASelectedFeaArrayInCurrentFeaclass(IMap mMap, int objectClassID)
        {            
            var list = GetSelectedFeaListInCurrentFeaclass(mMap, objectClassID);
            if (list ==null || list.Count == 0)
                return null;
            IFeature[] feaArray = new IFeature[list.Count];
            for (int i = 0; i < list.Count; i++)
            {
                feaArray[i] = list[i];
            }
            return feaArray;
        }
    }
}
