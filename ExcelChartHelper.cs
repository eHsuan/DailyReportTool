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

            // Ensure maxVal is at least a small positive number to avoid Excel errors
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

                // Manual Layout
                if (ctPlotArea.layout == null) ctPlotArea.layout = new CT_Layout();
                ctPlotArea.layout.manualLayout = new CT_ManualLayout();
                ctPlotArea.layout.manualLayout.xMode = new CT_LayoutMode { val = ST_LayoutMode.edge };
                ctPlotArea.layout.manualLayout.yMode = new CT_LayoutMode { val = ST_LayoutMode.edge };
                ctPlotArea.layout.manualLayout.x = new CT_Double { val = 0.10 };
                ctPlotArea.layout.manualLayout.y = new CT_Double { val = 0.10 };
                ctPlotArea.layout.manualLayout.w = new CT_Double { val = 0.85 };
                ctPlotArea.layout.manualLayout.h = new CT_Double { val = 0.75 };

                uint catAxId = 1001;
                uint valAxId1 = 1002;
                uint valAxId2 = 1003;

                // 1. Category Axis
                CT_CatAx catAx = new CT_CatAx { axId = new CT_UnsignedInt { val = catAxId } };
                catAx.scaling = new CT_Scaling { orientation = new CT_Orientation { val = ST_Orientation.minMax } };
                catAx.delete = new CT_Boolean { val = 0 };
                catAx.axPos = new CT_AxPos { val = ST_AxPos.b };
                catAx.majorTickMark = new CT_TickMark { val = ST_TickMark.@out };
                catAx.tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo };
                catAx.crossAx = new CT_UnsignedInt { val = valAxId1 };
                ctPlotArea.catAx.Add(catAx);

                // 2. Primary Value Axis (Left - Hours)
                CT_ValAx valAx1 = new CT_ValAx { axId = new CT_UnsignedInt { val = valAxId1 } };
                valAx1.scaling = new CT_Scaling();
                valAx1.scaling.orientation = new CT_Orientation { val = ST_Orientation.minMax };
                valAx1.scaling.min = new CT_Double { val = 0.000001 }; // Force fixed minimum 0
                valAx1.scaling.max = new CT_Double { val = maxVal };
                valAx1.delete = new CT_Boolean { val = 0 };
                valAx1.axPos = new CT_AxPos { val = ST_AxPos.l };
                valAx1.crossAx = new CT_UnsignedInt { val = catAxId };
                valAx1.crosses = new CT_Crosses { val = ST_Crosses.autoZero };
                valAx1.numFmt = new CT_NumFmt { formatCode = "#,##0.00", sourceLinked = false };
                valAx1.majorGridlines = new CT_ChartLines();
                valAx1.majorTickMark = new CT_TickMark { val = ST_TickMark.@out };
                ctPlotArea.valAx.Add(valAx1);

                // 3. Secondary Value Axis (Right - Percentage)
                CT_ValAx valAx2 = new CT_ValAx { axId = new CT_UnsignedInt { val = valAxId2 } };
                valAx2.scaling = new CT_Scaling();
                valAx2.scaling.orientation = new CT_Orientation { val = ST_Orientation.minMax };
                valAx2.scaling.min = new CT_Double { val = 0.000001 }; // Force fixed minimum 0
                valAx2.scaling.max = new CT_Double { val = 1.0 }; // Force fixed maximum 100%
                valAx2.delete = new CT_Boolean { val = 0 };
                valAx2.axPos = new CT_AxPos { val = ST_AxPos.r };
                valAx2.crossAx = new CT_UnsignedInt { val = catAxId };
                valAx2.crosses = new CT_Crosses { val = ST_Crosses.max }; // Force axis to right
                valAx2.numFmt = new CT_NumFmt { formatCode = "0%", sourceLinked = false };
                valAx2.tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo };
                valAx2.majorTickMark = new CT_TickMark { val = ST_TickMark.@out };
                ctPlotArea.valAx.Add(valAx2);

                // 4. Bar Chart
                CT_BarChart barChart = new CT_BarChart();
                barChart.barDir = new CT_BarDir { val = ST_BarDir.col };
                barChart.axId = new List<CT_UnsignedInt> { new CT_UnsignedInt { val = catAxId }, new CT_UnsignedInt { val = valAxId1 } };
                CT_BarSer barSer = barChart.AddNewSer();
                barSer.idx = new CT_UnsignedInt { val = 0 };
                barSer.order = new CT_UnsignedInt { val = 0 };
                barSer.tx = new CT_SerTx { v = "當機時數(小時)" };
                barSer.cat = new CT_AxDataSource { strRef = new CT_StrRef { f = GetRangeString(dataSheetName, 0, 1, dataCount) } };
                barSer.val = new CT_NumDataSource { numRef = new CT_NumRef { f = GetRangeString(dataSheetName, 1, 1, dataCount) } };
                ctPlotArea.barChart.Add(barChart);

                // 5. Line Chart
                CT_LineChart lineChart = new CT_LineChart();
                lineChart.axId = new List<CT_UnsignedInt> { new CT_UnsignedInt { val = catAxId }, new CT_UnsignedInt { val = valAxId2 } };
                CT_LineSer lineSer = lineChart.AddNewSer();
                lineSer.idx = new CT_UnsignedInt { val = 1 };
                lineSer.order = new CT_UnsignedInt { val = 1 };
                lineSer.tx = new CT_SerTx { v = "累積百分比" };
                lineSer.cat = new CT_AxDataSource { strRef = new CT_StrRef { f = GetRangeString(dataSheetName, 0, 1, dataCount) } };
                lineSer.val = new CT_NumDataSource { numRef = new CT_NumRef { f = GetRangeString(dataSheetName, 2, 1, dataCount) } };
                ctPlotArea.lineChart.Add(lineChart);
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
            // CRITICAL: Draw anchor in a safe area
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 0, 15, 18);
            IChart chart = drawing.CreateChart(anchor);
            chart.SetTitle(chartTitle);

            if (chart is XSSFChart xssfChart)
            {
                CT_Chart ctChart = xssfChart.GetCTChart();
                CT_PlotArea ctPlotArea = ctChart.plotArea;
                uint cId = 3001; uint vId = 3002;

                // 1. Stacked Bar Chart
                CT_BarChart bc = ctPlotArea.AddNewBarChart();
                bc.grouping = new CT_BarGrouping { val = ST_BarGrouping.stacked };
                bc.barDir = new CT_BarDir { val = ST_BarDir.col };
                bc.overlap = new CT_Overlap { val = 100 };
                bc.axId = new List<CT_UnsignedInt> { new CT_UnsignedInt { val = cId }, new CT_UnsignedInt { val = vId } };

                string catR = $"'{dataSheetName}'!$B$21:${GetExcelColumnName(dataCount + 1)}$21";
                AddSer_Bar(bc, 0, "有效產能", catR, $"'{dataSheetName}'!$B$23:${GetExcelColumnName(dataCount + 1)}$23", "C0C0C0");
                AddSer_Bar(bc, 1, "產速損失", catR, $"'{dataSheetName}'!$B$25:${GetExcelColumnName(dataCount + 1)}$25", "FFFF00");
                AddSer_Bar(bc, 2, "機故損失", catR, $"'{dataSheetName}'!$B$24:${GetExcelColumnName(dataCount + 1)}$24", "FF0000");

                // 2. Line Chart (Target)
                CT_LineChart lc = ctPlotArea.AddNewLineChart();
                lc.grouping = new CT_Grouping { val = ST_Grouping.standard };
                lc.axId = new List<CT_UnsignedInt> { new CT_UnsignedInt { val = cId }, new CT_UnsignedInt { val = vId } };
                AddSer_Line(lc, 3, "目標產能", catR, $"'{dataSheetName}'!$B$22:${GetExcelColumnName(dataCount + 1)}$22", "00B0F0", true);

                // 3. Category Axis
                CT_CatAx catAx = ctPlotArea.AddNewCatAx();
                catAx.axId = new CT_UnsignedInt { val = cId };
                catAx.scaling = new CT_Scaling { orientation = new CT_Orientation { val = ST_Orientation.minMax } };
                catAx.delete = new CT_Boolean { val = 0 };
                catAx.axPos = new CT_AxPos { val = ST_AxPos.b };
                catAx.crossAx = new CT_UnsignedInt { val = vId };
                catAx.tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo };

                // 4. Value Axis
                CT_ValAx valAx = ctPlotArea.AddNewValAx();
                valAx.axId = new CT_UnsignedInt { val = vId };
                valAx.scaling = new CT_Scaling();
                valAx.scaling.orientation = new CT_Orientation { val = ST_Orientation.minMax };
                valAx.scaling.max = new CT_Double { val = yMax };
                valAx.scaling.min = new CT_Double { val = 0.000001 };
                valAx.delete = new CT_Boolean { val = 0 };
                valAx.axPos = new CT_AxPos { val = ST_AxPos.l };
                valAx.crossAx = new CT_UnsignedInt { val = cId };
                valAx.majorGridlines = new CT_ChartLines();
                valAx.numFmt = new CT_NumFmt { formatCode = "#,##0", sourceLinked = false };

                // Legend
                CT_Legend legend = ctChart.AddNewLegend();
                legend.legendPos = new CT_LegendPos { val = ST_LegendPos.b };
                legend.overlay = new CT_Boolean { val = 0 };
            }
        }

        private static void AddSer_Bar(CT_BarChart bc, int idx, string name, string catF, string valF, string rgb)
        {
            CT_BarSer s = bc.AddNewSer();
            s.idx = new CT_UnsignedInt { val = (uint)idx };
            s.order = new CT_UnsignedInt { val = (uint)idx };
            s.tx = new CT_SerTx { v = name };
            s.cat = new CT_AxDataSource { strRef = new CT_StrRef { f = catF } };
            s.val = new CT_NumDataSource { numRef = new CT_NumRef { f = valF } };
            s.spPr = new NPOI.OpenXmlFormats.Dml.Chart.CT_ShapeProperties { 
                solidFill = new CT_SolidColorFillProperties { srgbClr = new CT_SRgbColor { val = System.Text.Encoding.ASCII.GetBytes(rgb) } } 
            };
            s.dLbls = new CT_DLbls { showVal = new CT_Boolean { val = 1 }, dLblPos = new CT_DLblPos { val = ST_DLblPos.ctr } };
        }

        private static void AddSer_Line(CT_LineChart lc, int idx, string name, string catF, string valF, string rgb, bool isDash)
        {
            CT_LineSer s = lc.AddNewSer();
            s.idx = new CT_UnsignedInt { val = (uint)idx };
            s.order = new CT_UnsignedInt { val = (uint)idx };
            s.tx = new CT_SerTx { v = name };
            s.cat = new CT_AxDataSource { strRef = new CT_StrRef { f = catF } };
            s.val = new CT_NumDataSource { numRef = new CT_NumRef { f = valF } };
            
            CT_LineProperties ln = new CT_LineProperties { 
                solidFill = new CT_SolidColorFillProperties { srgbClr = new CT_SRgbColor { val = System.Text.Encoding.ASCII.GetBytes(rgb) } } 
            };
            if (isDash) ln.prstDash = new CT_PresetLineDashProperties { val = ST_PresetLineDashVal.dash };
            
            s.spPr = new NPOI.OpenXmlFormats.Dml.Chart.CT_ShapeProperties { ln = ln };
            s.marker = new CT_Marker { symbol = new CT_MarkerStyle { val = ST_MarkerStyle.none } };
            s.smooth = new CT_Boolean { val = 0 };
        }

        private static string GetExcelColumnName(int n)
        {
            int d = n; string c = "";
            while (d > 0) { int m = (d - 1) % 26; c = Convert.ToChar(65 + m) + c; d = (d - m) / 26; }
            return c;
        }
    }
}
