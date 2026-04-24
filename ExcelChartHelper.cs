/* 
 * Author: YH CHIU
 */
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.OpenXmlFormats.Dml;
using System;
using System.Collections.Generic;

namespace DailyReportTool
{
    public static class ExcelChartHelper
    {
        public static void CreateParetoChart(ISheet sheet, string dataSheetName, int dataCount, string title, double maxVal)
        {
            if (dataCount == 0) return;
            if (maxVal <= 0) maxVal = 1.0;

            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 0, 10, 20);
            IChart chart = drawing.CreateChart(anchor);
            chart.SetTitle(title);

            if (chart is XSSFChart xssfChart)
            {
                CT_Chart ctChart = xssfChart.GetCTChart();
                CT_PlotArea ctPlotArea = ctChart.plotArea;

                if (ctPlotArea.barChart == null) ctPlotArea.barChart = new List<CT_BarChart>();
                if (ctPlotArea.lineChart == null) ctPlotArea.lineChart = new List<CT_LineChart>();
                if (ctPlotArea.valAx == null) ctPlotArea.valAx = new List<CT_ValAx>();
                if (ctPlotArea.catAx == null) ctPlotArea.catAx = new List<CT_CatAx>();

                if (ctPlotArea.layout == null) ctPlotArea.layout = new CT_Layout();
                ctPlotArea.layout.manualLayout = new CT_ManualLayout();
                ctPlotArea.layout.manualLayout.xMode = new CT_LayoutMode { val = ST_LayoutMode.edge };
                ctPlotArea.layout.manualLayout.yMode = new CT_LayoutMode { val = ST_LayoutMode.edge };
                ctPlotArea.layout.manualLayout.x = new CT_Double { val = 0.10 };
                ctPlotArea.layout.manualLayout.y = new CT_Double { val = 0.10 };
                ctPlotArea.layout.manualLayout.w = new CT_Double { val = 0.85 };
                ctPlotArea.layout.manualLayout.h = new CT_Double { val = 0.75 };

                uint catAxId = 1001; uint valAxId1 = 1002; uint valAxId2 = 1003;

                CT_CatAx catAx = new CT_CatAx { axId = new CT_UnsignedInt { val = catAxId } };
                catAx.scaling = new CT_Scaling { orientation = new CT_Orientation { val = ST_Orientation.minMax } };
                catAx.delete = new CT_Boolean { val = 0 };
                catAx.axPos = new CT_AxPos { val = ST_AxPos.b };
                catAx.majorTickMark = new CT_TickMark { val = ST_TickMark.@out };
                catAx.tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo };
                catAx.crossAx = new CT_UnsignedInt { val = valAxId1 };
                ctPlotArea.catAx.Add(catAx);

                CT_ValAx valAx1 = new CT_ValAx { axId = new CT_UnsignedInt { val = valAxId1 } };
                valAx1.scaling = new CT_Scaling { orientation = new CT_Orientation { val = ST_Orientation.minMax }, min = new CT_Double { val = 0.0 }, max = new CT_Double { val = maxVal } };
                valAx1.delete = new CT_Boolean { val = 0 };
                valAx1.axPos = new CT_AxPos { val = ST_AxPos.l };
                valAx1.crossAx = new CT_UnsignedInt { val = catAxId };
                valAx1.crosses = new CT_Crosses { val = ST_Crosses.autoZero };
                valAx1.numFmt = new CT_NumFmt { formatCode = "#,##0", sourceLinked = false };
                valAx1.majorGridlines = new CT_ChartLines();
                valAx1.majorTickMark = new CT_TickMark { val = ST_TickMark.@out };
                ctPlotArea.valAx.Add(valAx1);

                CT_ValAx valAx2 = new CT_ValAx { axId = new CT_UnsignedInt { val = valAxId2 } };
                valAx2.scaling = new CT_Scaling { orientation = new CT_Orientation { val = ST_Orientation.minMax }, min = new CT_Double { val = 0.0 }, max = new CT_Double { val = 1.0 } };
                valAx2.delete = new CT_Boolean { val = 0 };
                valAx2.axPos = new CT_AxPos { val = ST_AxPos.r };
                valAx2.crossAx = new CT_UnsignedInt { val = catAxId };
                valAx2.crosses = new CT_Crosses { val = ST_Crosses.max };
                valAx2.numFmt = new CT_NumFmt { formatCode = "0%", sourceLinked = false };
                valAx2.tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo };
                valAx2.majorTickMark = new CT_TickMark { val = ST_TickMark.@out };
                ctPlotArea.valAx.Add(valAx2);

                CT_BarChart bc = new CT_BarChart();
                bc.barDir = new CT_BarDir { val = ST_BarDir.col };
                bc.axId = new List<CT_UnsignedInt> { new CT_UnsignedInt { val = catAxId }, new CT_UnsignedInt { val = valAxId1 } };
                AddSer_Bar(bc, 0, "當機時數(小時)", GetRangeString(dataSheetName, 0, 1, dataCount), GetRangeString(dataSheetName, 1, 1, dataCount), "0F9ED5", true, "000000");
                ctPlotArea.barChart.Add(bc);

                CT_LineChart lc = new CT_LineChart();
                lc.axId = new List<CT_UnsignedInt> { new CT_UnsignedInt { val = catAxId }, new CT_UnsignedInt { val = valAxId2 } };
                AddSer_Line(lc, 1, "累積百分比", GetRangeString(dataSheetName, 0, 1, dataCount), GetRangeString(dataSheetName, 2, 1, dataCount), "000000", false, true, "000000");
                ctPlotArea.lineChart.Add(lc);
            }
        }

        private static string GetRangeString(string sheetName, int colIndex, int startRow, int endRow)
        {
            string colLetter = CellReference.ConvertNumToColString(colIndex);
            return $"'{sheetName}'!${colLetter}${startRow + 1}:${colLetter}${endRow + 1}";
        }

        public static void CreateCapacityBalanceChart(ISheet sheet, string dataSheetName, int dataCount, string chartTitle, double yMax)
        {
            if (dataCount == 0) return;
            if (yMax <= 0) yMax = 1000.0;

            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 0, 15, 18);
            IChart chart = drawing.CreateChart(anchor);
            chart.SetTitle(chartTitle);

            if (chart is XSSFChart xssfChart)
            {
                CT_Chart ctChart = xssfChart.GetCTChart();
                CT_PlotArea ctPlotArea = ctChart.plotArea;

                if (ctPlotArea.barChart == null) ctPlotArea.barChart = new List<CT_BarChart>();
                if (ctPlotArea.lineChart == null) ctPlotArea.lineChart = new List<CT_LineChart>();
                if (ctPlotArea.valAx == null) ctPlotArea.valAx = new List<CT_ValAx>();
                if (ctPlotArea.catAx == null) ctPlotArea.catAx = new List<CT_CatAx>();

                if (ctPlotArea.layout == null) ctPlotArea.layout = new CT_Layout();
                ctPlotArea.layout.manualLayout = new CT_ManualLayout();
                ctPlotArea.layout.manualLayout.xMode = new CT_LayoutMode { val = ST_LayoutMode.edge };
                ctPlotArea.layout.manualLayout.yMode = new CT_LayoutMode { val = ST_LayoutMode.edge };
                ctPlotArea.layout.manualLayout.x = new CT_Double { val = 0.10 };
                ctPlotArea.layout.manualLayout.y = new CT_Double { val = 0.10 };
                ctPlotArea.layout.manualLayout.w = new CT_Double { val = 0.85 };
                ctPlotArea.layout.manualLayout.h = new CT_Double { val = 0.75 };

                uint catId = 4001; uint valId = 4002;

                CT_CatAx catAx = new CT_CatAx { axId = new CT_UnsignedInt { val = catId } };
                catAx.scaling = new CT_Scaling { orientation = new CT_Orientation { val = ST_Orientation.minMax } };
                catAx.delete = new CT_Boolean { val = 0 };
                catAx.axPos = new CT_AxPos { val = ST_AxPos.b };
                catAx.majorTickMark = new CT_TickMark { val = ST_TickMark.@out };
                catAx.tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo };
                catAx.crossAx = new CT_UnsignedInt { val = valId };
                ctPlotArea.catAx.Add(catAx);

                CT_ValAx valAx = new CT_ValAx { axId = new CT_UnsignedInt { val = valId } };
                valAx.scaling = new CT_Scaling { orientation = new CT_Orientation { val = ST_Orientation.minMax }, min = new CT_Double { val = 0.0 }, max = new CT_Double { val = yMax } };
                valAx.delete = new CT_Boolean { val = 0 };
                valAx.axPos = new CT_AxPos { val = ST_AxPos.l };
                valAx.crossAx = new CT_UnsignedInt { val = catId };
                valAx.crosses = new CT_Crosses { val = ST_Crosses.autoZero };
                valAx.majorGridlines = new CT_ChartLines();
                valAx.numFmt = new CT_NumFmt { formatCode = "#,##0", sourceLinked = false };
                ctPlotArea.valAx.Add(valAx);

                CT_BarChart bc = new CT_BarChart();
                bc.grouping = new CT_BarGrouping { val = ST_BarGrouping.stacked };
                bc.barDir = new CT_BarDir { val = ST_BarDir.col };
                bc.overlap = new CT_Overlap { val = 100 };
                bc.axId = new List<CT_UnsignedInt> { new CT_UnsignedInt { val = catId }, new CT_UnsignedInt { val = valId } };
                
                string cR = $"'{dataSheetName}'!$B$21:${GetExcelColumnName(dataCount + 1)}$21";
                AddSer_Bar(bc, 0, "設備產能", cR, $"'{dataSheetName}'!$B$23:${GetExcelColumnName(dataCount + 1)}$23", "0F9ED5", true, "000000", ST_DLblPos.ctr);
                AddSer_Bar(bc, 1, "產速損失", cR, $"'{dataSheetName}'!$B$25:${GetExcelColumnName(dataCount + 1)}$25", "FFFF00", false, "000000", ST_DLblPos.ctr);
                AddSer_Bar(bc, 2, "機故損失", cR, $"'{dataSheetName}'!$B$24:${GetExcelColumnName(dataCount + 1)}$24", "FF0000", true, "FF0000", ST_DLblPos.inBase);
                ctPlotArea.barChart.Add(bc);

                CT_LineChart lc = new CT_LineChart();
                lc.grouping = new CT_Grouping { val = ST_Grouping.standard };
                lc.axId = new List<CT_UnsignedInt> { new CT_UnsignedInt { val = catId }, new CT_UnsignedInt { val = valId } };
                AddSer_Line(lc, 3, "目標產能", cR, $"'{dataSheetName}'!$B$22:${GetExcelColumnName(dataCount + 1)}$22", "00B0F0", true, false, "000000");
                ctPlotArea.lineChart.Add(lc);

                if (ctChart.legend == null) ctChart.legend = new CT_Legend();
                ctChart.legend.legendPos = new CT_LegendPos { val = ST_LegendPos.b };
                ctChart.legend.overlay = new CT_Boolean { val = 0 };
            }
        }

        private static void AddSer_Bar(CT_BarChart bc, int idx, string name, string catF, string valF, string rgb, bool showLbl, string lblColor, ST_DLblPos pos = ST_DLblPos.ctr)
        {
            CT_BarSer s = bc.AddNewSer();
            s.idx = new CT_UnsignedInt { val = (uint)idx };
            s.order = new CT_UnsignedInt { val = (uint)idx };
            s.tx = new CT_SerTx { v = name };
            s.cat = new CT_AxDataSource { strRef = new CT_StrRef { f = catF } };
            s.val = new CT_NumDataSource { numRef = new CT_NumRef { f = valF } };
            s.spPr = new NPOI.OpenXmlFormats.Dml.Chart.CT_ShapeProperties { solidFill = new CT_SolidColorFillProperties { srgbClr = new CT_SRgbColor { val = HexToBytes(rgb) } } };
            
            if (showLbl) {
                s.dLbls = new CT_DLbls();
                SetDataLabels(s.dLbls, lblColor, pos);
            } else s.dLbls = null;
        }

        private static void AddSer_Line(CT_LineChart lc, int idx, string name, string catF, string valF, string rgb, bool isDash, bool showLbl, string lblColor)
        {
            CT_LineSer s = lc.AddNewSer();
            s.idx = new CT_UnsignedInt { val = (uint)idx };
            s.order = new CT_UnsignedInt { val = (uint)idx };
            s.tx = new CT_SerTx { v = name };
            s.cat = new CT_AxDataSource { strRef = new CT_StrRef { f = catF } };
            s.val = new CT_NumDataSource { numRef = new CT_NumRef { f = valF } };
            CT_LineProperties ln = new CT_LineProperties { solidFill = new CT_SolidColorFillProperties { srgbClr = new CT_SRgbColor { val = HexToBytes(rgb) } } };
            if (isDash) ln.prstDash = new CT_PresetLineDashProperties { val = ST_PresetLineDashVal.dash };
            s.spPr = new NPOI.OpenXmlFormats.Dml.Chart.CT_ShapeProperties { ln = ln };
            s.marker = new CT_Marker { symbol = new CT_MarkerStyle { val = ST_MarkerStyle.none } };
            s.smooth = new CT_Boolean { val = 0 };

            if (showLbl) {
                s.dLbls = new CT_DLbls();
                SetDataLabels(s.dLbls, lblColor, ST_DLblPos.ctr);
            } else s.dLbls = null;
        }

        private static void SetDataLabels(CT_DLbls dLbls, string lblColor, ST_DLblPos pos)
        {
            dLbls.dLblPos = new CT_DLblPos { val = pos };
            dLbls.showVal = new CT_Boolean { val = 1 };
            dLbls.showCatName = new CT_Boolean { val = 0 };
            dLbls.showSerName = new CT_Boolean { val = 0 };
            dLbls.showPercent = new CT_Boolean { val = 0 };
            dLbls.showLegendKey = new CT_Boolean { val = 0 };
            dLbls.txPr = new NPOI.OpenXmlFormats.Dml.Chart.CT_TextBody {
                bodyPr = new CT_TextBodyProperties(),
                lstStyle = new CT_TextListStyle(),
                p = new List<CT_TextParagraph> {
                    new CT_TextParagraph {
                        pPr = new CT_TextParagraphProperties {
                            defRPr = new CT_TextCharacterProperties {
                                sz = 1000, solidFill = new CT_SolidColorFillProperties { srgbClr = new CT_SRgbColor { val = HexToBytes(lblColor) } },
                                latin = new CT_TextFont { typeface = "Arial" }
                            }
                        }
                    }
                }
            };
        }

        private static byte[] HexToBytes(string hex)
        {
            if (hex.StartsWith("#")) hex = hex.Substring(1);
            byte[] bytes = new byte[3];
            for (int i = 0; i < 3; i++) bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
            return bytes;
        }

        private static string GetExcelColumnName(int n)
        {
            int d = n; string c = "";
            while (d > 0) { int m = (d - 1) % 26; c = Convert.ToChar(65 + m) + c; d = (d - m) / 26; }
            return c;
        }
    }
}
