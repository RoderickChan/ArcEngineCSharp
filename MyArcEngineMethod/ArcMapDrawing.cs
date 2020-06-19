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
    [Guid("485df3d7-03ee-4df7-9048-33c0a4184342")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyArcEngineMethod.ArcMapDrawing")]
    public class ArcMapDrawing
    {

        #region 屏幕绘制函数

        /// <summary>
        /// 绘制点
        /// </summary>
        /// <param name="curView">当前视窗</param>
        /// <param name="point">点</param>
        /// <param name="color">颜色</param>
        /// <param name="radius">点的半径</param>
        public static void DrawPoint_Keep(IActiveView curView, IPoint point, IRgbColor color, double radius)
        {
            if ((point == null) || (curView == null))
            {
                return;
            }
            // 做点的缓冲区
            IPolygon buffer = (point as ITopologicalOperator).Buffer(radius) as IPolygon;
            // 绘制缓冲区
            DrawPolygon_Keep(curView, buffer, color, color, 0);
        }

        /// <summary>
        /// 绘制多边形
        /// </summary>
        /// <param name="curView">当前视窗</param>
        /// <param name="polygon">多边形</param>
        /// <param name="fillColor">填充色</param>
        /// <param name="outlineColor">轮廓色</param>
        /// <param name="outlineWidth">轮廓粗细</param>
        public static void DrawPolygon_Keep(IActiveView curView, IPolygon polygon, IRgbColor fillColor, IRgbColor outlineColor, double outlineWidth)
        {
            if ((polygon == null) || (curView == null))
            {
                return;
            }
            //边框样式
            ISimpleLineSymbol lineSymbol = new SimpleLineSymbolClass();
            lineSymbol = new SimpleLineSymbolClass();
            lineSymbol.Width = outlineWidth;
            lineSymbol.Color = outlineColor;
            lineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;
            //面域样式
            IFillSymbol fillSymbol = new SimpleFillSymbolClass();
            fillSymbol.Color = fillColor;
            fillSymbol.Outline = lineSymbol;
            //添加元素
            IFillShapeElement fillshapeElement = new PolygonElementClass() as IFillShapeElement;
            IElement element = fillshapeElement as IElement;
            fillshapeElement.Symbol = fillSymbol;
            element.Geometry = polygon as IGeometry;
            //绘制图形
            IGraphicsContainer graphicsContainer = curView.FocusMap as IGraphicsContainer;
            IActiveView activeView = graphicsContainer as IActiveView;
            graphicsContainer.AddElement(element, 0);
            //graphicsContainer.DeleteAllElements();//清空图层要素
            activeView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
        }

        /// <summary>
        /// 绘制线，需要手动清理
        /// </summary>
        /// <param name="curView">当前视窗</param>
        /// <param name="polyline">折线</param>
        /// <param name="lineColor">线颜色</param>
        /// <param name="lineWidth">线宽</param>
        public static void DrawPolyline_Keep(IActiveView curView, IPolyline polyline, IRgbColor lineColor, double lineWidth)
        {
            if ((polyline == null) || (curView == null))
            {
                return;
            }

            //线样式
            ISimpleLineSymbol lineSymbol = new SimpleLineSymbolClass();
            lineSymbol.Width = lineWidth;
            lineSymbol.Color = lineColor;
            lineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;
            //新建元素
            ILineElement lineElement = new LineElementClass();
            IElement element = lineElement as IElement;
            lineElement.Symbol = lineSymbol;
            element.Geometry = polyline;
            //绘制
            IGraphicsContainer graphicsContainer = curView.FocusMap as IGraphicsContainer;
            IActiveView activeView = graphicsContainer as IActiveView;
            graphicsContainer.AddElement(element, 0);

            activeView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);

        }

     /// <summary>
        /// 绘制多边形，并将text标在多边形的质心处
     /// </summary>
     /// <param name="curView">活动视窗</param>
     /// <param name="polygon">多边形</param>
     /// <param name="fillColor">填充色</param>
     /// <param name="outlineColor">轮廓色</param>
     /// <param name="outlineWidth">轮廓粗细</param>
     /// <param name="text">文字文本</param>
     /// <param name="textSize">文本大小</param>
        public static void DrawPolygonWithText_Keep(IActiveView curView, IPolygon polygon, IRgbColor fillColor, IRgbColor outlineColor, double outlineWidth, string text, int textSize)
        {
            if ((polygon == null) || (curView == null))
            {
                return;
            }

            ISimpleLineSymbol lineSymbol = new SimpleLineSymbolClass();
            lineSymbol.Width = outlineWidth;
            lineSymbol.Color = outlineColor;
            lineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;

            //填充样式
            ISimpleFillSymbol fillSymbol = new SimpleFillSymbolClass();
            fillSymbol.Color = fillColor;
            fillSymbol.Outline = lineSymbol;
            fillSymbol.Style = esriSimpleFillStyle.esriSFSSolid;

            //多边形区域
            IArea area = polygon as IArea;
            //多边形中心点
            IPoint textPoint = area.Centroid;
            //多边形面积
            double dArea = Math.Round(Math.Abs(area.Area), 2);

            //文本颜色
            IRgbColor textRGB;
            textRGB = new RgbColorClass();
            textRGB.Red = 10;
            textRGB.Green = 10;
            textRGB.Blue = 10;

            //文本样式
            ITextSymbol textSymbol;
            textSymbol = new ESRI.ArcGIS.Display.TextSymbolClass();
            textSymbol.HorizontalAlignment = esriTextHorizontalAlignment.esriTHACenter;
            textSymbol.VerticalAlignment = esriTextVerticalAlignment.esriTVACenter;
            textSymbol.Size = textSize;
            textSymbol.Color = textRGB;
            textSymbol.Text = text;

            curView.ScreenDisplay.StartDrawing(curView.ScreenDisplay.hDC, -1);
            curView.ScreenDisplay.SetSymbol(fillSymbol as ISymbol);
            curView.ScreenDisplay.DrawPolygon(polygon);
            curView.ScreenDisplay.SetSymbol(textSymbol as ISymbol);
            curView.ScreenDisplay.DrawText(textPoint, textSymbol.Text);
            curView.ScreenDisplay.FinishDrawing();

        }


        /// <summary>
        /// 获得随机颜色
        /// </summary>
        public static IRgbColor getColor()
        {
            Random random = new Random();
            IRgbColor color = new RgbColorClass();

            color.Red = random.Next(0, 255);
            color.Green = random.Next(0, 255);
            color.Blue = random.Next(0, 255);
            return color;
        }

        /// <summary>
        ///     获得一种RGB颜色
        /// </summary>
        /// <param name="R"></param>
        /// <param name="G"></param>
        /// <param name="B"></param>
        /// <returns></returns>
        public static IRgbColor getColor(int R, int G, int B)
        {
            IRgbColor color = new RgbColorClass();
            color.Red = R;
            color.Green = G;
            color.Blue = B;
            return color;
        }

        /// <summary>
        /// 清空绘制的内容
        /// </summary>      
        public static void ClearGraphics(IActiveView curView)
        {
            IGraphicsContainer graphicsContainer = (curView.FocusMap) as IGraphicsContainer;
            graphicsContainer.DeleteAllElements();
            curView.Refresh();
           
        }

        #endregion

        #region 屏幕绘制2

        /// <summary>
        /// 绘制多边形
        /// </summary>
        /// <param name="curView">当前视窗</param>
        /// <param name="polygon">多边形</param>
        /// <param name="fillcolor">填充色</param>
        /// <param name="outlineColor">轮廓色</param>
        /// <param name="outlineWidth">轮廓粗细</param>
        public static void Draw_Polygon(IActiveView curView, IPolygon polygon, IRgbColor fillcolor, IRgbColor outlineColor, double outlineWidth)
        {
            IPolygon m_polygon = polygon;
            IScreenDisplay screendisplay = curView.ScreenDisplay;

            ISimpleLineSymbol lineSymbol = new SimpleLineSymbolClass();
            lineSymbol.Width = outlineWidth;
            lineSymbol.Color = outlineColor;
            lineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;

            ISimpleFillSymbol fillsymbol = new SimpleFillSymbolClass();
            IRgbColor rgbColor = new RgbColorClass();
            rgbColor = fillcolor;
            fillsymbol.Color = rgbColor;
            fillsymbol.Outline = lineSymbol;

            screendisplay.StartDrawing(screendisplay.hDC, (short)esriScreenCache.esriNoScreenCache);
            screendisplay.SetSymbol((ISymbol)fillsymbol);
            screendisplay.DrawPolygon(m_polygon);
            screendisplay.FinishDrawing();
        }

        /// <summary>
        /// 绘制多义线
        /// </summary>
        /// <param name="curView">当前视窗</param>
        /// <param name="polyline">多义线</param>
        /// <param name="color">颜色</param>
        /// <param name="lineWidth">绘制线的粗细</param>
        public static void Draw_Polyline(IActiveView curView, IPolyline polyline, IRgbColor color, double lineWidth)
        {
            IPolyline m_polyline = polyline;
            IScreenDisplay screenDisplay = curView.ScreenDisplay;
            ISimpleLineSymbol lineSymbol = new SimpleLineSymbolClass();
            IRgbColor rgbColor = new RgbColorClass();
            rgbColor = color;
            lineSymbol.Color = rgbColor;
            lineSymbol.Width = lineWidth;
            screenDisplay.StartDrawing(screenDisplay.hDC, (short)esriScreenCache.esriNoScreenCache);
            screenDisplay.SetSymbol((ISymbol)lineSymbol);
            screenDisplay.DrawPolyline(m_polyline);
            screenDisplay.FinishDrawing();
        }

        /// <summary>
        /// 绘制点
        /// </summary>
        /// <param name="curView">当前视窗</param>
        /// <param name="point">点</param>
        /// <param name="color">颜色</param>
        /// <param name="radius">点的半径</param>
        public static void Draw_Point(IActiveView curView, IPoint point, IRgbColor color, int radius)
        {
            IPoint m_point = point;
            IScreenDisplay screenDisplay = curView.ScreenDisplay;
            ISimpleMarkerSymbol pMarkerSymbol = new SimpleMarkerSymbolClass();
            IRgbColor rgbColor = new RgbColorClass();
            rgbColor = color;
            pMarkerSymbol.Color = rgbColor;
            pMarkerSymbol.Size = radius;
            screenDisplay.StartDrawing(screenDisplay.hDC, (short)esriScreenCache.esriNoScreenCache);
            screenDisplay.SetSymbol((ISymbol)pMarkerSymbol);
            screenDisplay.DrawPoint(m_point);
            screenDisplay.FinishDrawing();
        }


        #endregion


        /// <summary>
        /// 对某一矩形区域进行刷白
        /// </summary>
        /// <param name="activeView">活动视窗</param>
        /// <param name="pEnvelope">要素类的包络</param>
        public static void RefreshArea(IActiveView activeView, IEnvelope pEnvelope)
        {
            IScreenDisplay pDisplay = activeView.ScreenDisplay;
            pDisplay.StartDrawing(pDisplay.hDC, 0);
            IRgbColor pColor = new RgbColorClass();
            pColor.Red = 255;
            pColor.Green = 255;
            pColor.Blue = 255;
            ISimpleLineSymbol simpleLineSymbol = new SimpleLineSymbolClass();
            simpleLineSymbol.Width = 10;
            simpleLineSymbol.Color = pColor;
            //simpleLineSymbol.Style = borderStyle;
            ISimpleFillSymbol simplePolySym = new SimpleFillSymbol();
            simplePolySym.Outline = simpleLineSymbol;
            // simplePolySym.Style = fillStyle;
            simplePolySym.Color = pColor;
            pDisplay.SetSymbol(simplePolySym as ISymbol);
            pDisplay.DrawRectangle(pEnvelope);
            pDisplay.FinishDrawing();

        }
    }
}
