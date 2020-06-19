using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.Geometry;
using System.Threading.Tasks;

namespace MyArcEngineMethod.MyExtensionMethod
{
    public static class IPointArrayExtension
    {
        /// <summary>
        /// 判断点数组里面是否包含某一点
        /// </summary>
        /// <param name="pArray">ipointarray实例</param>
        /// <param name="p">待查询点</param>
        /// <param name="deci">判断精度，默认为5位小数</param>
        /// <returns></returns>
        public static bool Contains(this IPointArray pArray, ESRI.ArcGIS.Geometry.IPoint p, int deci = 5)
        {
            if (pArray == null)
            {
                return false;
            }
            for (int i = 0; i < pArray.Count; i++)
            {
                if (Math.Round(p.X, deci) == Math.Round(pArray.Element[i].X, deci)
                    && Math.Round(p.Y, deci) == Math.Round(pArray.Element[i].Y, deci))
                    return true;
            }
            return false;

        }

        /// <summary>
        /// 用点数组构造多边形，数组中最后一个元素与第一个元素为同一个点
        /// </summary>
        /// <param name="pArray">点数组</param>
        /// <returns></returns>
        public static IPolygon GetPolygonByPArray(this IPointArray pArray)
        {
            if (pArray.Count < 3)
                return null;

            IPolygon mPolygon = new PolygonClass();
            ISegmentCollection pSegCol = new RingClass();
            object missing1 = Type.Missing; object missing2 = Type.Missing;
            for (int i = 0; i < pArray.Count - 1; i++)
            {
                ILine pLine = new LineClass();
                pLine.PutCoords(pArray.get_Element(i), pArray.get_Element(i + 1));
                pSegCol.AddSegment(pLine as ISegment, ref missing1, ref missing2);
            }

            //用环去构造矩形
            IRing pRing = pSegCol as IRing;
            pRing.Close();
            IGeometryCollection pPolygon = new PolygonClass();
            pPolygon.AddGeometry(pRing, ref missing1, ref missing2);

            mPolygon = pPolygon as IPolygon;
            (mPolygon as ITopologicalOperator).Simplify();

            return mPolygon;
        }


    }
}
