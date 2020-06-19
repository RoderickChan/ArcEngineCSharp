using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;
using System.Windows.Forms;
using ESRI.ArcGIS.AnalysisTools;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geoprocessor;

namespace MyArcEngineMethod
{
    [Guid("d74a5912-00aa-481e-8089-d3bf0bff33cb")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyArcEngineMethod.GeometryOperation")]
    public class GeometryOperation
    {
        /// <summary>
        /// 将包络线转化为多边形
        /// </summary>
        /// <param name="pEnvelope">包络线参数</param>
        /// <returns></returns>
        public static IPolygon EnvelopeToPolygon(IEnvelope pEnvelope)
        {
            IGeometryCollection pPolygon = new PolygonClass();
            ISegmentCollection seGmeCo = new RingClass();

            ESRI.ArcGIS.Geometry.IPoint p1 = null;
            ESRI.ArcGIS.Geometry.IPoint p2 = null;

            ESRI.ArcGIS.Geometry.ILine pLine = new ESRI.ArcGIS.Geometry.LineClass();
            p1 = pEnvelope.UpperLeft; p2 = pEnvelope.UpperRight;
            pLine.PutCoords(p1, p2);
            seGmeCo.AddSegment(pLine as ISegment, Type.Missing, Type.Missing);

            pLine = new ESRI.ArcGIS.Geometry.LineClass();
            p1 = pEnvelope.UpperRight; p2 = pEnvelope.LowerRight;
            pLine.PutCoords(p1, p2);
            seGmeCo.AddSegment(pLine as ISegment, Type.Missing, Type.Missing);

            pLine = new ESRI.ArcGIS.Geometry.LineClass();
            p1 = pEnvelope.LowerRight; p2 = pEnvelope.LowerLeft;
            pLine.PutCoords(p1, p2);
            seGmeCo.AddSegment(pLine as ISegment, Type.Missing, Type.Missing);

            pLine = new ESRI.ArcGIS.Geometry.LineClass();
            p1 = pEnvelope.LowerLeft; p2 = pEnvelope.UpperLeft;
            pLine.PutCoords(p1, p2);
            seGmeCo.AddSegment(pLine as ISegment, Type.Missing, Type.Missing);

            IRing pRing = new RingClass();
            pRing = seGmeCo as IRing;
            pRing.Close();

            pPolygon.AddGeometry(pRing as IGeometry, Type.Missing, Type.Missing);

            return pPolygon as IPolygon;
        }

        /// <summary>
        /// 用点数组构造多边形，数组中最后一个元素与第一个元素为同一个点
        /// </summary>
        /// <param name="pArray">点数组</param>
        /// <returns></returns>
        public static IPolygon GetPolygonByPArray(IPointArray pArray)
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

        /// <summary>
        /// 用一组有顺序的点构造polyline
        /// </summary>
        /// <param name="pArray"></param>
        /// <returns></returns>
        public static IPolyline GetPolylineByPArray(IPointArray pArray, ISpatialReference sr = null)
        {
            if (pArray.Count < 2)
                return null;
            
            IGeometryCollection pPolyline = new PolylineClass();
            ISegmentCollection pPath = new PathClass();
            object missing1 = Type.Missing; object missing2 = Type.Missing;
            for (int i = 0; i < pArray.Count - 1; i++)
            {
                ILine pLine = new LineClass();
                pLine.PutCoords(pArray.get_Element(i), pArray.get_Element(i + 1));
                pPath.AddSegment(pLine as ISegment, ref missing1, ref missing2);
            }
            //将路径添加到线上
            pPolyline.AddGeometry(pPath as IGeometry, ref missing1, ref missing2);

            IPolyline py = null;
            py = pPolyline as IPolyline;
            if (sr != null)
                py.SpatialReference = sr;
            if (py != null || py.Length > 0)
            {
                return py;           
            }
            else
                return null;
        }

        /// <summary>
        /// 根据点获取边与XY轴平行的矩形
        /// </summary>
        /// <param name="mp">控制点</param>
        /// <param name="length">矩形长</param>
        /// <param name="width">矩形宽</param>
        /// <param name="isCenter">如果是true，则点mp是矩形的中心点；如果为false，则点mp为矩形的顶点之一</param>
        /// <returns></returns>
        public static IPolygon GetRectangleByPoint(IPoint mp, double length, double width, bool isCenter)
        {
            IPolygon mPolygon = new PolygonClass();

            //构造好矩形的四个点，初始矩形两个边与X和Y轴平行,点mp为矩形中心点
            double mp_X = mp.X; double mp_Y = mp.Y;
            double L_half = length / 2.0; double W_half = width / 2.0;
            IPoint p1 = new PointClass(); IPoint p2 = new PointClass();
            IPoint p3 = new PointClass(); IPoint p4 = new PointClass();
            if (isCenter)
            {
                p1.PutCoords(mp_X - L_half, mp_Y + W_half); p2.PutCoords(mp_X + L_half, mp_Y + W_half);
                p4.PutCoords(mp_X - L_half, mp_Y - W_half); p3.PutCoords(mp_X + L_half, mp_Y - W_half);
            }
            else
            {
                p1.PutCoords(mp_X, mp_Y + width); p2.PutCoords(mp_X + length, mp_Y + width);
                p4.PutCoords(mp_X, mp_Y); p3.PutCoords(mp_X + length, mp_Y);
            }

            //开始构造环

            ISegmentCollection pSegCol = new RingClass();
            object missing1 = Type.Missing; object missing2 = Type.Missing;

            ILine pLine = new LineClass();
            pLine.PutCoords(p1, p2);
            pSegCol.AddSegment(pLine as ISegment, ref missing1, ref missing2);

            pLine = new LineClass();
            pLine.PutCoords(p2, p3);
            pSegCol.AddSegment(pLine as ISegment, ref missing1, ref missing2);

            pLine = new LineClass();
            pLine.PutCoords(p3, p4);
            pSegCol.AddSegment(pLine as ISegment, ref missing1, ref missing2);

            pLine = new LineClass();
            pLine.PutCoords(p4, p1);
            pSegCol.AddSegment(pLine as ISegment, ref missing1, ref missing2);

            //用环去构造矩形
            IRing pRing = pSegCol as IRing;
            pRing.Close();
            IGeometryCollection pPolygon = new PolygonClass();
            pPolygon.AddGeometry(pRing, ref missing1, ref missing2);

            mPolygon = pPolygon as IPolygon;
            (mPolygon as ITopologicalOperator).Simplify();

            return mPolygon;
        }

        /// <summary>
        /// 获取一个四分之三的边与XY轴平行的矩形
        /// </summary>
        /// <param name="p">控制点</param>
        /// <param name="length">矩形长</param>
        /// <param name="width">矩形宽</param>
        /// <returns></returns>
        public static IPolygon GetRectangleByPoint_Cut(IPoint p, double length, double width)
        {
            IPolygon mPolygon = new PolygonClass();
            //先创建一个与X,Y轴平行的长与宽指定的矩形
            mPolygon = GetRectangleByPoint(p, length, width, true);
            IPolygon mPolygon2 = new PolygonClass();
            //构造一个小矩形，然后切开一小半
            mPolygon2 = GetRectangleByPoint(p, length / 2.0, width / 2.0, false);
            mPolygon = (mPolygon as ITopologicalOperator).Difference(mPolygon2) as IPolygon;
            (mPolygon as ITopologicalOperator).Simplify();
            return mPolygon;
        }

        /// <summary>
        /// 将矩形绕着某点逆时针旋转一定角度，输出旋转后的矩形
        /// </summary>
        /// <param name="mPoly">待旋转矩形</param>
        /// <param name="p">绕旋转点</param>
        /// <param name="angle">旋转角度，弧度</param>
        /// <returns></returns>
        public IPolygon RotateRectAntiCWByPoint(IPolygon mPoly, IPoint p, double angle = 0)//angle为弧度
        {
            IPolygon poly = null;
            double p_x = p.X; double p_y = p.Y;
            double sin_angle = Math.Sin(angle); double cos_angle = Math.Cos(angle);
            //计算每个点的相对坐标,并旋转
            IPointCollection pc = mPoly as IPointCollection;
            List<IPoint> pco = new List<IPoint>();
            for (int i = 0; i < pc.PointCount - 1; i++)
            {
                IPoint curP = pc.get_Point(i);
                //drawpoint(curP,getColor(255,0,0),5);
                double new_x = curP.X - p_x;
                double new_y = curP.Y - p_y;
                double new_2x = new_x * cos_angle + new_y * sin_angle;
                double new_2y = new_y * cos_angle - new_x * sin_angle;
                new_x = new_2x + p_x; new_y = new_2y + p_y;

                IPoint now_p = new PointClass();
                now_p.PutCoords(new_x, new_y);
                (now_p as ITopologicalOperator).Simplify();
                //drawpoint(now_p,getColor(),5);
                pco.Add(now_p);
            }

            ISegmentCollection pSegCol = new RingClass();
            object missing1 = Type.Missing; object missing2 = Type.Missing;
            for (int j = 0; j < pco.Count; j++)
            {
                ILine pLine = new LineClass();
                if ((j + 1) == pco.Count)
                {
                    pLine.PutCoords(pco[j], pco[0]);
                    pSegCol.AddSegment(pLine as ISegment, ref missing1, ref missing2);
                }
                else
                {
                    pLine.PutCoords(pco[j], pco[j + 1]);
                    pSegCol.AddSegment(pLine as ISegment, ref missing1, ref missing2);
                }
            }

            //用环去构造矩形
            IRing pRing = pSegCol as IRing;
            pRing.Close();
            IGeometryCollection pPolygon = new PolygonClass();
            pPolygon.AddGeometry(pRing, ref missing1, ref missing2);

            poly = pPolygon as IPolygon;
            (poly as ITopologicalOperator).Simplify();
            return poly;
        }

        /// <summary>
        /// 切割线切割尖角，得到碎片多边形
        /// </summary>>
        /// <param name="Geo">尖角</param>
        /// <param name="cutLineSet">切割线集合</param>
        public static List<IGeometry> CutGeo(IGeometry Geo, List<IPolyline> cutLineSet)
        {
            List<IGeometry> fragPolygonSet = new List<IGeometry>() { Geo };
            if (cutLineSet.Count == 0)
                return fragPolygonSet;
            for (int i = 0; i < fragPolygonSet.Count; i++)
            {
                IGeometry fragPolygon = fragPolygonSet[i];
                int firstCutIndex = FirstCut(fragPolygon, cutLineSet);
                if (firstCutIndex == -1) //该碎片多边形很单纯，不能再被切割
                    continue;
                IGeometryCollection geoColl = null;
                try
                {
                    geoColl = (fragPolygon as ITopologicalOperator4).Cut2(cutLineSet[firstCutIndex]);
                }
                catch
                {
                    continue;
                }
                fragPolygonSet.RemoveAt(i);
                i = i - 1;

                for (int j = 0; j < geoColl.GeometryCount; j++)
                {
                    fragPolygonSet.Add(geoColl.get_Geometry(j) as IPolygon);
                }
            }
            return fragPolygonSet;
        }

        /// <summary>
        /// 返回切割碎片多边形第一条找到的切割线索引
        /// </summary>>
        /// <param name="fragPolygon">碎片多边形</param>
        /// <param name="cutLineSet">切割线集合</param>
        private static int FirstCut(IGeometry fragPolygon, List<IPolyline> cutLineSet)
        {
            int index = -1;
            for (int i = 0; i < cutLineSet.Count; i++)
            {
                IPolyline cutLine = cutLineSet[i];
                if ((fragPolygon as IRelationalOperator).Crosses(cutLine) &&
                    !(fragPolygon as IRelationalOperator).Touches(cutLine))
                {
                    index = i;
                    break;
                }
            }
            return index;
        }

        /// <summary>
        /// 判断多边形是否为普通多边形，即不含有内环或多个外环，普通返回-1，有内环返回0，有外环返回1
        /// </summary>
        /// <param name="geo">几何形状</param>
        /// <returns></returns>
        public static int JudgeGeometryNormal(IGeometry geo)
        {
            IPolygon polygon = geo as IPolygon;
            IPolygon4 poly = polygon as IPolygon4;
            if (poly.ExteriorRingCount > 1) return 0;//若外环数大于1，则返回0
            else
            {
                //获取到多边形的内环
                IGeometryBag bag = poly.ExteriorRingBag;
                IEnumGeometry enumGeo = bag as IEnumGeometry;
                enumGeo.Reset();
                IRing exRing = null;
                int innerRingCount = 0;
                while ((exRing = enumGeo.Next() as IRing) != null)
                {
                    innerRingCount = innerRingCount + poly.get_InteriorRingCount(exRing);
                    break;
                }
                //若多边形内环数为0，则返回-1
                if (innerRingCount > 0) return 1;
                else //普通多边形返回-1
                    return -1;
            }
        }

        /// <summary>
        /// 计算两点之间的距离
        /// </summary>
        /// <param name="p1">点1</param>
        /// <param name="p2">点2</param>
        /// <returns></returns>
        public static double Points2Distance(IPoint p1, IPoint p2)
        {
            double k = (p1.X - p2.X) * (p1.X - p2.X) + (p1.Y - p2.Y) * (p1.Y - p2.Y);
            double dis = Math.Sqrt(k);
            return dis;
        }

        /// <summary>
        /// 获取点到直线段的对称点
        /// </summary>
        /// <param name="point">查询点</param>
        /// <param name="line">直线段</param>
        /// <returns></returns>
        public static IPoint QuerySymmetricPoint(IPoint point, ILine line)
        {
            // 对称点
            IPoint symmetricPoint = new PointClass();
            ISegment seg = line as ISegment;
            // 写出seg所在直线的方程式（一般式 Ax+By+C=0）
            double A = seg.ToPoint.Y - seg.FromPoint.Y;
            double B = seg.FromPoint.X - seg.ToPoint.X;
            double C = (seg.FromPoint.Y - seg.ToPoint.Y) * seg.FromPoint.X + (seg.ToPoint.X - seg.FromPoint.X) * seg.FromPoint.Y;
            // 计算对称点的坐标
            double numeratorX = B * B * point.X - 2 * A * B * point.Y - 2 * A * C - A * A * point.X;
            double numeratorY = A * A * point.Y - 2 * A * B * point.X - 2 * B * C - B * B * point.Y;
            double denominator = A * A + B * B;
            symmetricPoint.X = numeratorX / denominator;
            symmetricPoint.Y = numeratorY / denominator;

            return symmetricPoint;
        }


        /// <summary>
        /// 获取点到直线段的垂点
        /// </summary>
        /// <param name="point">查询点</param>
        /// <param name="line">直线段</param>
        /// <param name="extendLength">直线段向两端延长的距离，默认为10米</param>
        /// <returns></returns>
        public static IPoint QueryFootPoint(IPoint point, ILine line, double extendLength = 10)
        {
            IPoint footPoint = new PointClass();
            IPoint symmetricPoint = QuerySymmetricPoint(point, line); //找point相对于line的对称点
            IPolyline perpPolyline = new PolylineClass();
            //连接point和对称点作为垂线    
            perpPolyline.FromPoint = point;
            perpPolyline.ToPoint = symmetricPoint;
            (perpPolyline as ITopologicalOperator).Simplify();
            //将line包装成IPolyline
            IPolyline polyline = new PolylineClass();
            polyline.FromPoint = line.FromPoint;
            polyline.ToPoint = line.ToPoint;
            (polyline as ITopologicalOperator).Simplify();
            polyline = QueryExtendLine(polyline, 3, extendLength); //从左右两端10米开外延长polyline
            //获得垂线与polyline的交点
            IGeometry geo = (polyline as ITopologicalOperator).Intersect(perpPolyline, esriGeometryDimension.esriGeometry0Dimension);
            if (!geo.IsEmpty)
            {
                IPointCollection pointCollection = geo as IPointCollection;
                footPoint = pointCollection.get_Point(0);
            }
            return footPoint;
        }

        /// <summary>
        /// 延长线段。模式：1为从FromPoint处延长，2为从ToPint处延长，3为两端延长
        /// </summary>
        /// <param name="inputPolyline">传入去的线</param>
        /// <param name="mode">模式，1为从FromPoint处延长，2为从ToPint处延长，3为两端延长</param>
        /// <param name="dis">延长的距离</param>
        /// <returns></returns>
        public static IPolyline QueryExtendLine(IPolyline inputPolyline, int mode, double dis)
        {
            IPointCollection pPointCol = inputPolyline as IPointCollection;
            switch (mode)
            {
                case 1:
                    IPoint fromPoint = new PointClass();
                    inputPolyline.QueryPoint(esriSegmentExtension.esriExtendAtFrom, -1 * dis, false, fromPoint);
                    pPointCol.InsertPoints(0, 1, ref fromPoint);
                    break;
                case 2:
                    IPoint endPoint = new PointClass();
                    object missing = Type.Missing;
                    inputPolyline.QueryPoint(esriSegmentExtension.esriExtendAtTo, dis + inputPolyline.Length, false, endPoint);
                    pPointCol.AddPoint(endPoint, ref missing, ref missing);
                    break;
                case 3:
                    IPoint fPoint = new PointClass();
                    IPoint ePoint = new PointClass();
                    object missing2 = Type.Missing;
                    inputPolyline.QueryPoint(esriSegmentExtension.esriExtendAtFrom, -1 * dis, false, fPoint);
                    pPointCol.InsertPoints(0, 1, ref fPoint);
                    inputPolyline.QueryPoint(esriSegmentExtension.esriExtendAtTo, dis + inputPolyline.Length, false, ePoint);
                    pPointCol.AddPoint(ePoint, ref missing2, ref missing2);
                    break;
                default:
                    return pPointCol as IPolyline;
            }
            return pPointCol as IPolyline;
        }

        /// <summary>
        /// 建立最小外接矩形
        /// </summary>
        /// <param name="pPoly">经多边形建立的凸包</param>
        public static IPolygon MBR(IPolygon pPoly)
        {
            IPolygon m_Poly = new PolygonClass();
            IPointCollection pPtInPolygon = pPoly as IPointCollection;
            IPointCollection pPt = new PolygonClass();
            IPointCollection pPtInRect = new PolygonClass();
            List<double> terminal = new List<double>();//存储矩形Xmin，Xmax，Ymin，Ymax四端点
            List<double> Ang = new List<double>();
            IPoint Org, temp;
            Org = new PointClass();//原点
            temp = new PointClass();
            Org.X = 0; Org.Y = 0;
            Terminal(pPtInPolygon, terminal);
            double min_Area = (terminal[2] - terminal[0]) * (terminal[3] - terminal[1]);
            for (int i = 0; i < pPtInPolygon.PointCount - 1; i++)//求出旋转角
            {
                temp.X = pPtInPolygon.get_Point(i + 1).X - pPtInPolygon.get_Point(i).X;
                temp.Y = pPtInPolygon.get_Point(i + 1).Y - pPtInPolygon.get_Point(i).Y;
                if (temp.X * 0 + temp.Y * 1 > 0 && temp.X > 0)
                    Ang.Add(System.Math.Acos((temp.X * 0 + temp.Y * 1) / Points2Distance(Org, temp)));
                if (temp.X * 0 + temp.Y * -1 > 0 && temp.X < 0)
                    Ang.Add(System.Math.Acos((temp.X * 0 + temp.Y * -1) / Points2Distance(Org, temp)));
                if (temp.X * -1 + temp.Y * 0 > 0 && temp.Y > 0)
                    Ang.Add(System.Math.Acos((temp.X * -1 + temp.Y * 0) / Points2Distance(Org, temp)));
                if (temp.X * 1 + temp.Y * 0 > 0 && temp.Y < 0)
                    Ang.Add(System.Math.Acos((temp.X * 1 + temp.Y * 0) / Points2Distance(Org, temp)));
            }
            temp.X = terminal[0]; temp.Y = terminal[1];
            pPtInRect.AddPoint(temp);
            temp.X = terminal[0]; temp.Y = terminal[3];
            pPtInRect.AddPoint(temp);
            temp.X = terminal[2]; temp.Y = terminal[3];
            pPtInRect.AddPoint(temp);
            temp.X = terminal[2]; temp.Y = terminal[1];
            pPtInRect.AddPoint(temp);
            for (int i = 0; i < Ang.Count; i++)
            {
                for (int j = 0; j < pPtInPolygon.PointCount; j++)//旋转多边形至卡壳贴边
                {
                    temp.X = pPtInPolygon.get_Point(j).X * System.Math.Cos(Ang[i]) - pPtInPolygon.get_Point(j).Y * System.Math.Sin(Ang[i]);
                    temp.Y = pPtInPolygon.get_Point(j).X * System.Math.Sin(Ang[i]) + pPtInPolygon.get_Point(j).Y * System.Math.Cos(Ang[i]);
                    pPt.AddPoint(temp);
                }
                terminal.Clear();
                Terminal(pPt, terminal);
                pPt.RemovePoints(0, pPt.PointCount);
                if (min_Area > (terminal[2] - terminal[0]) * (terminal[3] - terminal[1]))
                {
                    min_Area = (terminal[2] - terminal[0]) * (terminal[3] - terminal[1]);
                    pPtInRect.RemovePoints(0, 4);//清空，矩形只存4个点
                    temp.X = terminal[0]; temp.Y = terminal[1];
                    pPtInRect.AddPoint(temp);
                    temp.X = terminal[0]; temp.Y = terminal[3];
                    pPtInRect.AddPoint(temp);
                    temp.X = terminal[2]; temp.Y = terminal[3];
                    pPtInRect.AddPoint(temp);
                    temp.X = terminal[2]; temp.Y = terminal[1];
                    pPtInRect.AddPoint(temp);
                    for (int k = 0; k < 4; k++)//将矩形反旋转至原始位置
                    {
                        temp.X = pPtInRect.get_Point(k).X * System.Math.Cos(-Ang[i]) - pPtInRect.get_Point(k).Y * System.Math.Sin(-Ang[i]);
                        temp.Y = pPtInRect.get_Point(k).X * System.Math.Sin(-Ang[i]) + pPtInRect.get_Point(k).Y * System.Math.Cos(-Ang[i]);
                        pPtInRect.AddPoint(temp);
                    }
                    pPtInRect.RemovePoints(0, 4);
                }
            }
            pPtInRect.AddPoint(pPtInRect.get_Point(0));
            m_Poly = pPtInRect as IPolygon;
            return m_Poly;
        }

        /// <summary>
        /// 求多边形四端点
        /// </summary>
        /// <param name="pPtInPolygon">待求端点的多边形</param>
        /// <param name="terminal">端点存储集合</param>
        public static void Terminal(IPointCollection pPtInPolygon, List<double> terminal)
        {
            double Xmin = double.MaxValue;
            double Ymin = double.MaxValue;
            double Xmax = double.MinValue;
            double Ymax = double.MinValue;
            for (int i = 0; i < pPtInPolygon.PointCount; i++)
            {
                if (Xmin > pPtInPolygon.get_Point(i).X)
                    Xmin = pPtInPolygon.get_Point(i).X;
                if (Ymin > pPtInPolygon.get_Point(i).Y)
                    Ymin = pPtInPolygon.get_Point(i).Y;
                if (Xmax < pPtInPolygon.get_Point(i).X)
                    Xmax = pPtInPolygon.get_Point(i).X;
                if (Ymax < pPtInPolygon.get_Point(i).Y)
                    Ymax = pPtInPolygon.get_Point(i).Y;
            }
            terminal.Add(Xmin);
            terminal.Add(Ymin);
            terminal.Add(Xmax);
            terminal.Add(Ymax);
        }

        /// <summary>
        /// 寻找外接矩形长边
        /// </summary>
        /// <param name="pPlg">矩形</param>
        public static double FindLongAngle(IPolygon pPlg)
        {
            double angle;//长边角度(限制在0-PI之间)
            IPointCollection pPtInPolygon = pPlg as IPointCollection;
            double length1 = Points2Distance(pPtInPolygon.get_Point(0), pPtInPolygon.get_Point(1));
            double length2 = Points2Distance(pPtInPolygon.get_Point(1), pPtInPolygon.get_Point(2));
            if (length1 > length2)
            {
                if ((pPtInPolygon.get_Point(1).X - pPtInPolygon.get_Point(0).X) == 0)
                    angle = System.Math.PI / 2;
                else
                {
                    angle = System.Math.Atan((pPtInPolygon.get_Point(1).Y - pPtInPolygon.get_Point(0).Y) / (pPtInPolygon.get_Point(1).X - pPtInPolygon.get_Point(0).X));
                    if (angle < 0)
                        angle = angle + System.Math.PI;
                }
            }
            else
            {
                if ((pPtInPolygon.get_Point(2).X - pPtInPolygon.get_Point(1).X) == 0)
                    angle = System.Math.PI / 2;
                else
                {
                    angle = System.Math.Atan((pPtInPolygon.get_Point(2).Y - pPtInPolygon.get_Point(1).Y) / (pPtInPolygon.get_Point(2).X - pPtInPolygon.get_Point(1).X));
                    if (angle < 0)
                        angle = angle + System.Math.PI;
                }
            }

            return angle;
        }

        /// <summary>
        /// 根据prj文件创建空间参考
        /// </summary>
        /// <param name="strProFile">空间参照文件</param>
        /// <returns></returns>
        public static ISpatialReference CreateSpatialReference(string strProFile)
        {
            ISpatialReferenceFactory pSpatialReferenceFactory = new SpatialReferenceEnvironmentClass();
            ISpatialReference pSpatialReference = pSpatialReferenceFactory.CreateESRISpatialReferenceFromPRJFile(strProFile);
            return pSpatialReference;
        }

        /// <summary>
        /// 得到折线的矩形缓冲区
        /// </summary>
        /// <param name="poly"></param>
        /// <param name="bufferDis2">矩形宽度的二分之一</param>
        /// <param name="cuts">输出两条切割线</param>
        /// <returns></returns>
        public static IGeometry CreatRectangleBufferByPolyline(IPolyline poly, double bufferDis2, List<IPolyline> cuts)
        {
            cuts.Clear();

            IGeometry simulateGeo = (poly as ITopologicalOperator).Buffer(bufferDis2);
            //drawpolygon(simulateGeo as IPolygon, getColor(), getColor(), 1);
            ILine line10 = new LineClass();
            ILine line11 = new LineClass();
            ILine line20 = new LineClass();
            ILine line21 = new LineClass();
            poly.QueryNormal(esriSegmentExtension.esriNoExtension, 0, true, bufferDis2 * 1.5, line10);
            poly.QueryNormal(esriSegmentExtension.esriNoExtension, 1, true, bufferDis2 * 1.5, line11);
            poly.ReverseOrientation();
            poly.QueryNormal(esriSegmentExtension.esriNoExtension, 0, true, bufferDis2 * 1.5, line20);
            poly.QueryNormal(esriSegmentExtension.esriNoExtension, 1, true, bufferDis2 * 1.5, line21);
            IGeometryCollection cutline1 = new PolylineClass();
            IGeometryCollection cutline2 = new PolylineClass();
            ISegmentCollection path1 = new PathClass();
            ISegmentCollection path2 = new PathClass();
            path1.AddSegment(line10 as ISegment);
            path1.AddSegment(line21 as ISegment);
            cutline1.AddGeometry(path1 as IGeometry);
            path2.AddSegment(line11 as ISegment);
            path2.AddSegment(line20 as ISegment);
            cutline2.AddGeometry(path2 as IGeometry);

            (cutline1 as ITopologicalOperator).Simplify();
            (cutline2 as ITopologicalOperator).Simplify();
            IPolyline cut1 = cutline1 as IPolyline; (cut1 as ITopologicalOperator).Simplify();
            IPolyline cut2 = cutline2 as IPolyline; (cut2 as ITopologicalOperator).Simplify();
            //drawpolyline(cut1, getColor(), 1.5);
            //drawpolyline(cut2, getColor(), 1.5);
            cuts.Add(cut1); cuts.Add(cut2);

            IGeometryCollection gc = null;
            try
            {
                gc = (simulateGeo as ITopologicalOperator4).Cut2(cut1);
            }
            catch 
            {
                return null;
            }
            if ((gc.get_Geometry(0) as IArea).Area > (gc.get_Geometry(1) as IArea).Area)
            {
                simulateGeo = gc.get_Geometry(0);
            }
            else
            {
                simulateGeo = gc.get_Geometry(1);
            }

            try
            {
                gc = (simulateGeo as ITopologicalOperator4).Cut2(cut2);
            }
            catch
            {
                return null;
            }
            if ((gc.get_Geometry(0) as IArea).Area > (gc.get_Geometry(1) as IArea).Area)
            {
                simulateGeo = gc.get_Geometry(0);
            }
            else
            {
                simulateGeo = gc.get_Geometry(1);
            }
            //drawpolygon(simulateGeo as IPolygon, getColor(), getColor(), 1);
            return simulateGeo;
        }


        

    }
}
