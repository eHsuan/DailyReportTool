/* 
 * Author: YH CHIU
 */
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.OpenXmlFormats.Dml;
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

            // Data assumed to be in:
            // Col 0: Category Name
            // Col 1: Value (Hours)
            // Col 2: Cumulative Percentage

            int dataStartRow = 1;
            int dataEndRow = dataCount;

            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            // Anchor: Col 4 (E) to Col 14 (O), Row 0 to Row 20
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 0, 10, 20);

            IChart chart = drawing.CreateChart(anchor);
            chart.SetTitle(title);

            // NPOI 2.5.6 Chart interface is limited, we drop down to CT_Chart
            if (chart is XSSFChart xssfChart)
            {
                CT_Chart ctChart = xssfChart.GetCTChart();
                CT_PlotArea ctPlotArea = ctChart.plotArea;

                // Ensure lists exist
                if (ctPlotArea.barChart == null) ctPlotArea.barChart = new List<CT_BarChart>();
                if (ctPlotArea.lineChart == null) ctPlotArea.lineChart = new List<CT_LineChart>();
                if (ctPlotArea.valAx == null) ctPlotArea.valAx = new List<CT_ValAx>();
                if (ctPlotArea.catAx == null) ctPlotArea.catAx = new List<CT_CatAx>();

                // Manual Layout (Optional, mimics TPPlatoTool for better fit)
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

                // 1. Category Axis (Horizontal)
                CT_CatAx catAx = new CT_CatAx();
                catAx.axId = new CT_UnsignedInt { val = catAxId };
                catAx.scaling = new CT_Scaling();
                catAx.scaling.orientation = new CT_Orientation { val = ST_Orientation.minMax };
                catAx.delete = new CT_Boolean { val = 0 };
                catAx.axPos = new CT_AxPos { val = ST_AxPos.b };
                catAx.majorTickMark = new CT_TickMark { val = ST_TickMark.@out };
                catAx.tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo };
                catAx.crossAx = new CT_UnsignedInt { val = valAxId1 };
                ctPlotArea.catAx.Add(catAx);

                // 2. Primary Value Axis (Left - Hours)
                CT_ValAx valAx1 = new CT_ValAx();
                valAx1.axId = new CT_UnsignedInt { val = valAxId1 };
                valAx1.scaling = new CT_Scaling();
                valAx1.scaling.orientation = new CT_Orientation { val = ST_Orientation.minMax };
                valAx1.scaling.max = new CT_Double { val = maxVal };
                valAx1.scaling.min = new CT_Double { val = 0.0 }; // Force minimum to 0
                valAx1.delete = new CT_Boolean { val = 0 };
                valAx1.axPos = new CT_AxPos { val = ST_AxPos.l };
                valAx1.crossAx = new CT_UnsignedInt { val = catAxId };
                valAx1.majorGridlines = new CT_ChartLines();
                ctPlotArea.valAx.Add(valAx1);

                // 3. Val Axis 2 (Secondary - Right - Percentage)
                CT_ValAx valAx2 = new CT_ValAx();
                valAx2.axId = new CT_UnsignedInt { val = valAxId2 };
                valAx2.scaling = new CT_Scaling();
                valAx2.scaling.orientation = new CT_Orientation { val = ST_Orientation.minMax };
                valAx2.scaling.max = new CT_Double { val = 1.0 };
                valAx2.delete = new CT_Boolean { val = 0 };
                valAx2.axPos = new CT_AxPos { val = ST_AxPos.r };
                valAx2.crossAx = new CT_UnsignedInt { val = catAxId };
                valAx2.crosses = new CT_Crosses { val = ST_Crosses.max }; // Force axis to the right
                valAx2.numFmt = new CT_NumFmt { formatCode = "0%", sourceLinked = false };
                valAx2.tickLblPos = new CT_TickLblPos { val = ST_TickLblPos.nextTo };
                ctPlotArea.valAx.Add(valAx2);


                // 4. Bar Chart (Downtime Hours)
                CT_BarChart barChart = new CT_BarChart();
                barChart.barDir = new CT_BarDir { val = ST_BarDir.col };
                barChart.axId = new List<CT_UnsignedInt> {
                    new CT_UnsignedInt { val = catAxId },
                    new CT_UnsignedInt { val = valAxId1 }
                };

                CT_BarSer barSer = barChart.AddNewSer();
                barSer.idx = new CT_UnsignedInt { val = 0 };
                barSer.order = new CT_UnsignedInt { val = 0 };
                barSer.tx = new CT_SerTx { v = "當機時數(小時)" };
                barSer.cat = new CT_AxDataSource { strRef = new CT_StrRef { f = GetRangeString(dataSheetName, 0, 1, dataCount) } };
                barSer.val = new CT_NumDataSource { numRef = new CT_NumRef { f = GetRangeString(dataSheetName, 1, 1, dataCount) } };
                ctPlotArea.barChart.Add(barChart);

                // 5. Line Chart (Cumulative Percentage)
                CT_LineChart lineChart = new CT_LineChart();
                lineChart.axId = new List<CT_UnsignedInt> {
                    new CT_UnsignedInt { val = catAxId },
                    new CT_UnsignedInt { val = valAxId2 }
                };

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
            // Range: Sheet!$A$2:$A$10
            return $"'{sheetName}'!${colLetter}${startRow + 1}:${colLetter}${endRow + 1}";
        }
    }
}
