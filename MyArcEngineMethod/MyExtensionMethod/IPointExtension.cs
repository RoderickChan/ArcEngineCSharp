using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.Geometry;
using System.Threading.Tasks;

namespace MyArcEngineMethod.MyExtensionMethod
{
  public static  class IPointExtension
    {

      /// <summary>
      /// 计算该点到另一个点的距离
      /// </summary>
      /// <param name="p1"></param>
      /// <param name="targetPoint">目标点</param>
      /// <returns>距离</returns>
      public static double GetDistance(this IPoint p1, IPoint targetPoint)
      {
          if (targetPoint == null)
              return -1;
          double dis = (p1.X - targetPoint.X) * (p1.X - targetPoint.X) +
              (p1.Y - targetPoint.Y) * (p1.Y - targetPoint.Y);
          return Math.Sqrt(dis);
      }

      /// <summary>
      /// 获取两点之间的中间点
      /// </summary>
      /// <param name="p">操作点</param>
      /// <param name="targetPoint">目标点</param>
      /// <returns>中间点</returns>
      public static IPoint GetMidpoint(this IPoint p, IPoint targetPoint)
      {
          if (targetPoint == null)
              throw new ArgumentNullException("Error: The target point is null!");
          IPoint midP = new PointClass();
          midP.PutCoords((p.X + targetPoint.X) / 2,  (p.Y + targetPoint.Y) / 2);
          return midP;
      }
  }
}
