/* 
 * Author: YH CHIU
 * Description: Daily Report Tool for equipment maintenance aggregation and Pareto chart generation.
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using HtmlAgilityPack;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;

namespace DailyReportTool
{
    public partial class Form1 : Form
    {
        private JObject config;
        private Dictionary<string, List<string>> processMap;
        private Dictionary<string, string> equipCategoryMap; // Key: EquipName, Value: uPOL or uBMU

        public Form1()
        {
            InitializeComponent();
            this.Text = "DailyReportTool v1.0.1";
            LoadConfig();
        }

        private void LoadConfig()
        {
            try
            {
                if (File.Exists("config.json"))
                {
                    string json = File.ReadAllText("config.json");
                    config = JObject.Parse(json);
                    processMap = new Dictionary<string, List<string>>();
                    equipCategoryMap = new Dictionary<string, string>();

                    // 1. Load Process Map (Original Logic)
                    JToken processSection = config["Process"];
                    if (processSection != null)
                    {
                        foreach (JProperty proc in processSection.Children<JProperty>())
                        {
                            string processName = proc.Name;
                            string equipmentsStr = (string)proc.Value["Equipments"];
                            List<string> equipments = equipmentsStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                                   .Select(e => e.Trim())
                                                                   .ToList();
                            processMap[processName] = equipments;
                        }
                    }

                    // 2. Load Non-Common Category (New Logic)
                    JToken nonCommonSection = config["Non-Common"];
                    if (nonCommonSection != null)
                    {
                        foreach (JProperty type in nonCommonSection.Children<JProperty>())
                        {
                            string typeName = type.Name; // uPOL, uBMU
                            string equipmentsStr = (string)type.Value["Equipments"];
                            if (!string.IsNullOrEmpty(equipmentsStr))
                            {
                                string[] equipments = equipmentsStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                foreach (var eq in equipments)
                                {
                                    equipCategoryMap[eq.Trim()] = typeName;
                                }
                            }
                        }
                    }
                    Log("Config loaded successfully.");
                }
            }
            catch (Exception ex)
            {
                Log("Error loading config: " + ex.Message);
            }
        }

        private void Log(string message)
        {
            string logMsg = $"{DateTime.Now:yyyy/MM/dd/HH:mm:ss} {message}";
            txtLog.AppendText(logMsg + Environment.NewLine);
        }

        // --- Navigation Logic ---
        private void menuPareto_Click(object sender, EventArgs e)
        {
            pnlPareto.Visible = true;
            pnlBalance.Visible = false;
        }

        private void menuBalance_Click(object sender, EventArgs e)
        {
            pnlPareto.Visible = false;
            pnlBalance.Visible = true;
        }

        // --- Original Pareto Logic (Untouched) ---
        private void btnSelectMaintenance_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel/HTML Files|*.xlsx;*.xls;*.html;*.htm";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtMaintenancePath.Text = ofd.FileName;
                    Log("Selected Maintenance Record: " + ofd.FileName);
                }
            }
        }

        private void btnSelectConnection_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel/HTML Files|*.xlsx;*.xls;*.html;*.htm";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtConnectionPath.Text = ofd.FileName;
                    Log("Selected Connection Data: " + ofd.FileName);
                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtMaintenancePath.Text = "";
            txtConnectionPath.Text = "";
            Log("Inputs cleared.");
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtMaintenancePath.Text) && string.IsNullOrEmpty(txtConnectionPath.Text))
            {
                MessageBox.Show("Please select at least one input file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Dictionary<string, double> equipmentHours = new Dictionary<string, double>();
                if (!string.IsNullOrEmpty(txtMaintenancePath.Text) && File.Exists(txtMaintenancePath.Text)) ProcessMaintenanceFile(txtMaintenancePath.Text, equipmentHours);
                if (!string.IsNullOrEmpty(txtConnectionPath.Text) && File.Exists(txtConnectionPath.Text)) ProcessConnectionFile(txtConnectionPath.Text, equipmentHours);

                var processStats = new List<Tuple<string, double>>();
                foreach (var kvp in processMap)
                {
                    double totalHours = 0;
                    foreach (var eq in kvp.Value) { if (equipmentHours.ContainsKey(eq)) totalHours += equipmentHours[eq]; }
                    if (totalHours > 0) processStats.Add(Tuple.Create(kvp.Key, totalHours));
                }
                processStats.Sort((a, b) => b.Item2.CompareTo(a.Item2));
                GenerateReport(processStats, equipmentHours);
                Log("Report Generation Complete.");
            }
            catch (Exception ex)
            {
                Log("Error: " + ex.Message);
            }
        }

        private bool IsHtmlFile(string path)
        {
            try {
                using (StreamReader reader = new StreamReader(path)) {
                    char[] buffer = new char[512]; int charsRead = reader.Read(buffer, 0, buffer.Length);
                    string content = new string(buffer, 0, charsRead).ToLower();
                    return content.Contains("<html") || content.Contains("<table");
                }
            } catch { return false; }
        }

        private void ProcessMaintenanceFile(string path, Dictionary<string, double> equipmentHours)
        {
            if (IsHtmlFile(path)) ProcessMaintenanceHtml(path, equipmentHours);
            else ProcessMaintenanceExcel(path, equipmentHours);
        }

        private void ProcessMaintenanceExcel(string path, Dictionary<string, double> equipmentHours)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook workbook = WorkbookFactory.Create(fs);
                ISheet sheet = workbook.GetSheetAt(0); 
                for (int i = 1; i <= sheet.LastRowNum; i++) 
                {
                    IRow row = sheet.GetRow(i); if (row == null) continue;
                    string eqId = GetCellValue(row.GetCell(1));
                    string type = GetCellValue(row.GetCell(5));
                    if (type != null && type.Contains("A-維修"))
                    {
                        double hours = GetNumericValue(row.GetCell(8));
                        if (!equipmentHours.ContainsKey(eqId)) equipmentHours[eqId] = 0;
                        equipmentHours[eqId] += hours;
                    }
                }
                workbook.Close();
            }
        }

        private void ProcessMaintenanceHtml(string path, Dictionary<string, double> equipmentHours)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(path, Encoding.UTF8); 
            var rows = doc.DocumentNode.SelectNodes("//tr"); if (rows == null) return;
            foreach (var row in rows.Skip(1))
            {
                var cells = row.SelectNodes("td|th"); if (cells == null || cells.Count < 10) continue; 
                string eqId = System.Net.WebUtility.HtmlDecode(cells[2].InnerText).Trim();
                string type = System.Net.WebUtility.HtmlDecode(cells[6].InnerText).Trim();
                if (type != null && type.Contains("A-維修"))
                {
                    if (double.TryParse(cells[9].InnerText, out double hours))
                    {
                        if (!equipmentHours.ContainsKey(eqId)) equipmentHours[eqId] = 0;
                        equipmentHours[eqId] += hours;
                    }
                }
            }
        }

        private void ProcessConnectionFile(string path, Dictionary<string, double> equipmentHours)
        {
            if (IsHtmlFile(path)) ProcessConnectionHtml(path, equipmentHours);
            else ProcessConnectionExcel(path, equipmentHours);
        }

        private void ProcessConnectionExcel(string path, Dictionary<string, double> equipmentHours)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook workbook = WorkbookFactory.Create(fs);
                ISheet sheet = workbook.GetSheetAt(0);
                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i); if (row == null) continue;
                    string eqId = GetCellValue(row.GetCell(1));
                    DateTime? start = row.GetCell(6)?.DateCellValue;
                    DateTime? end = row.GetCell(7)?.DateCellValue;
                    if (start.HasValue && end.HasValue)
                    {
                        double hours = (end.Value - start.Value).TotalHours;
                        if (hours > 0) { if (!equipmentHours.ContainsKey(eqId)) equipmentHours[eqId] = 0; equipmentHours[eqId] += hours; }
                    }
                }
                workbook.Close();
            }
        }

        private void ProcessConnectionHtml(string path, Dictionary<string, double> equipmentHours)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(path, Encoding.UTF8);
            var rows = doc.DocumentNode.SelectNodes("//tr"); if (rows == null) return;
            foreach (var row in rows.Skip(1))
            {
                var cells = row.SelectNodes("td|th"); if (cells == null || cells.Count < 9) continue; 
                string eqId = System.Net.WebUtility.HtmlDecode(cells[2].InnerText).Trim();
                if (DateTime.TryParse(cells[7].InnerText, out DateTime start) && DateTime.TryParse(cells[8].InnerText, out DateTime end))
                {
                    double hours = (end - start).TotalHours;
                    if (hours > 0) { if (!equipmentHours.ContainsKey(eqId)) equipmentHours[eqId] = 0; equipmentHours[eqId] += hours; }
                }
            }
        }

        private string GetCellValue(ICell cell)
        {
            if (cell == null) return "";
            if (cell.CellType == CellType.String) return cell.StringCellValue.Trim();
            if (cell.CellType == CellType.Numeric) return cell.NumericCellValue.ToString();
            return cell.ToString().Trim();
        }

        private double GetNumericValue(ICell cell)
        {
            if (cell == null) return 0;
            if (cell.CellType == CellType.Numeric) return cell.NumericCellValue;
            if (cell.CellType == CellType.String) { double val; if (double.TryParse(cell.StringCellValue, out val)) return val; }
            return 0;
        }

        private void GenerateReport(List<Tuple<string, double>> processStats, Dictionary<string, double> equipmentHours)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Files|*.xlsx";
                sfd.FileName = $"DailyReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return;

                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet1 = workbook.CreateSheet("總機故時數");
                CreateDataSheet(sheet1, processStats, "製程站點");

                ISheet sheet2 = workbook.CreateSheet("總機故柏拉圖");
                double totalHours = processStats.Sum(x => x.Item2);
                ExcelChartHelper.CreateParetoChart(sheet2, sheet1.SheetName, processStats.Count, "總機故柏拉圖", totalHours);

                int topCount = Math.Min(3, processStats.Count);
                for (int i = 0; i < topCount; i++)
                {
                    string processName = processStats[i].Item1;
                    var eqs = new List<Tuple<string, double>>();
                    if (processMap.ContainsKey(processName)) { foreach (var eq in processMap[processName]) if (equipmentHours.ContainsKey(eq) && equipmentHours[eq] > 0) eqs.Add(Tuple.Create(eq, equipmentHours[eq])); }
                    eqs.Sort((a, b) => b.Item2.CompareTo(a.Item2));

                    string snD = processName + "機故時數"; string snC = processName + "機故柏拉圖";
                    if (snD.Length > 31) snD = snD.Substring(0, 31); if (snC.Length > 31) snC = snC.Substring(0, 31);

                    ISheet sd = workbook.CreateSheet(snD); CreateDataSheet(sd, eqs, "設備編號");
                    ISheet sc = workbook.CreateSheet(snC); ExcelChartHelper.CreateParetoChart(sc, sd.SheetName, eqs.Count, processName + "柏拉圖", eqs.Sum(x => x.Item2));
                }

                using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write)) { workbook.Write(fs); }
                Log($"Report saved to {sfd.FileName}");
                MessageBox.Show("Report generated successfully!");
            }
        }

        private void CreateDataSheet(ISheet sheet, List<Tuple<string, double>> data, string nameHeader)
        {
            IRow header = sheet.CreateRow(0); header.CreateCell(0).SetCellValue(nameHeader); header.CreateCell(1).SetCellValue("當機時數(小時)"); header.CreateCell(2).SetCellValue("累積百分比");
            ICellStyle hourStyle = sheet.Workbook.CreateCellStyle(); hourStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat("0.00");
            ICellStyle pctStyle = sheet.Workbook.CreateCellStyle(); pctStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat("0%");

            double total = data.Sum(x => x.Item2); double currentSum = 0;
            for (int i = 0; i < data.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(data[i].Item1);
                ICell c1 = row.CreateCell(1); c1.SetCellValue(data[i].Item2); c1.CellStyle = hourStyle;
                currentSum += data[i].Item2;
                ICell c2 = row.CreateCell(2); c2.SetCellValue(total == 0 ? 0 : (currentSum / total)); c2.CellStyle = pctStyle;
            }
        }

        // --- NEW Independent Logic for Capacity Balance ---

        private class Balance_DataRow
        {
            public string Name;
            public string Category; // uPOL, uBMU, Common
            public double TargetPcsPerDay;
            public double TheoreticalPcsPerDay;
            public double EquipCount;
            public double WeightedAvgCT;
            public double WeightedAvgPass;
            public List<Balance_ProductInfo> Products = new List<Balance_ProductInfo>();
        }

        private class Balance_ProductInfo
        {
            public double CT, PcsPerLot, DailyLotDemand, PassCount;
        }

        private void btnImportIE_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls" })
                if (ofd.ShowDialog() == DialogResult.OK) { txtIEPath.Text = ofd.FileName; Log("IE Data selected: " + ofd.FileName); }
        }

        private void btnGenerateBalance_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtIEPath.Text)) return;
            try { Balance_Execute(); }
            catch (Exception ex) { Log("Balance Error: " + ex.Message); MessageBox.Show(ex.Message); }
        }

        private void Balance_Execute()
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = $"BalanceReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx" })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;
                IWorkbook workbook = new XSSFWorkbook();
                
                // 1. Read & Calculate
                List<Balance_DataRow> data = Balance_ReadIEData(txtIEPath.Text);
                Balance_PerformCalculations(data, (double)numDailyHours.Value);

                // 2. Render Sheet (CRITICAL: Drawing patriarch FIRST)
                ISheet sheet = workbook.CreateSheet("目標片次與設備能力");
                sheet.CreateDrawingPatriarch();
                Balance_RenderReport(sheet, data, (int)numDataDays.Value, (double)numDailyHours.Value);

                using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create)) { workbook.Write(fs); }
                Log("Success: " + sfd.FileName);
                MessageBox.Show("Report generated!");
            }
        }

        private List<Balance_DataRow> Balance_ReadIEData(string path)
        {
            List<Balance_DataRow> rows = new List<Balance_DataRow>();
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook wb = WorkbookFactory.Create(fs);
                IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
                ISheet ws = wb.GetSheet("IE主表");
                
                // Products from Row 3
                int prodCount = 0; IRow r3 = ws.GetRow(2);
                for (int c = 1; c < r3.LastCellNum; c++) { if (string.IsNullOrEmpty(r3.GetCell(c)?.ToString())) break; prodCount++; }

                // Data from Row 15
                for (int r = 14; r <= ws.LastRowNum; r++)
                {
                    IRow row = ws.GetRow(r);
                    if (row == null || Balance_IsMerged(ws, r, 0)) continue;
                    string name = row.GetCell(0)?.ToString().Trim();
                    if (string.IsNullOrEmpty(name) || name.Contains("專線") || name.Contains("C/T")) continue;

                    Balance_DataRow dr = new Balance_DataRow { Name = name };
                    dr.Category = equipCategoryMap.ContainsKey(name) ? equipCategoryMap[name] : "Common";
                    
                    for (int p = 0; p < prodCount; p++) {
                        int bc = 1 + p * 4;
                        dr.Products.Add(new Balance_ProductInfo {
                            CT = Balance_GetNum(row.GetCell(bc), eval),
                            PcsPerLot = Balance_GetNum(row.GetCell(bc+1), eval),
                            DailyLotDemand = Balance_GetNum(row.GetCell(bc+2), eval),
                            PassCount = Balance_GetNum(row.GetCell(bc+3), eval)
                        });
                    }
                    dr.TargetPcsPerDay = Balance_GetNum(row.GetCell(10), eval); // Col K
                    dr.EquipCount = processMap.ContainsKey(name) ? processMap[name].Count : 0;
                    rows.Add(dr);
                    Log($"[Read] {name} Target: {dr.TargetPcsPerDay}");
                }
            }
            return rows;
        }

        private void Balance_PerformCalculations(List<Balance_DataRow> rows, double dailyHours)
        {
            foreach (var r in rows)
            {
                double totalPcsPass = 0, totalPcsBase = 0, weightedCT = 0, weightedPass = 0;
                foreach (var p in r.Products) {
                    double basePcs = p.DailyLotDemand * p.PcsPerLot;
                    totalPcsBase += basePcs;
                    totalPcsPass += (basePcs * p.PassCount);
                    weightedCT += (p.CT * basePcs * p.PassCount);
                    weightedPass += (p.PassCount * basePcs);
                }
                r.WeightedAvgCT = (totalPcsPass > 0) ? (weightedCT / totalPcsPass) : 0;
                r.WeightedAvgPass = (totalPcsBase > 0) ? (weightedPass / totalPcsBase) : 1;
                
                double rawTheory = (r.EquipCount * dailyHours * 3600) / (r.WeightedAvgCT > 0 ? r.WeightedAvgCT : 1);
                r.TheoreticalPcsPerDay = rawTheory / (r.WeightedAvgPass > 0 ? r.WeightedAvgPass : 1);
            }
        }

        private void Balance_RenderReport(ISheet s, List<Balance_DataRow> rows, int days, double hours)
        {
            ICellStyle commonStyle = s.Workbook.CreateCellStyle(); commonStyle.FillForegroundColor = IndexedColors.LemonChiffon.Index; commonStyle.FillPattern = FillPattern.SolidForeground;
            ICellStyle specialStyle = s.Workbook.CreateCellStyle(); specialStyle.FillForegroundColor = IndexedColors.LightTurquoise.Index; specialStyle.FillPattern = FillPattern.SolidForeground;

            string[] headers = { "機台/站別", "目標片次", "有效產能", "機故損失", "產速損失" };
            for (int j = 0; j < 5; j++) { IRow r = s.GetRow(20 + j) ?? s.CreateRow(20 + j); r.CreateCell(0).SetCellValue(headers[j]); }

            for (int i = 0; i < rows.Count; i++)
            {
                int c = i + 1;
                ICell nameC = s.GetRow(20).CreateCell(c); nameC.SetCellValue(rows[i].Name);
                nameC.CellStyle = (rows[i].Category == "Common") ? commonStyle : specialStyle;

                s.GetRow(21).CreateCell(c).SetCellValue(rows[i].TargetPcsPerDay * days);
                s.GetRow(22).CreateCell(c).SetCellValue(rows[i].TheoreticalPcsPerDay * days);
                s.GetRow(23).CreateCell(c).SetCellValue(0);
                s.GetRow(24).CreateCell(c).SetCellValue(0);
            }

            double maxV = (rows.Count > 0) ? rows.Max(r => Math.Max(r.TargetPcsPerDay, r.TheoreticalPcsPerDay)) * days * 1.2 : 1000;
            ExcelChartHelper.CreateCapacityBalanceChart(s, s.SheetName, rows.Count, $"產能平衡圖 ({days}天, {hours}hr/日)", maxV);
        }

        private bool Balance_IsMerged(ISheet s, int r, int c) { for (int i = 0; i < s.NumMergedRegions; i++) { if (s.GetMergedRegion(i).IsInRange(r, c)) return true; } return false; }
        
        private double Balance_GetNum(ICell c, IFormulaEvaluator e)
        {
            if (c == null) return 0;
            if (c.CellType == CellType.Formula) { try { return e.Evaluate(c).NumberValue; } catch { return c.NumericCellValue; } }
            if (c.CellType == CellType.Numeric) return c.NumericCellValue;
            double v; return double.TryParse(c.ToString(), out v) ? v : 0;
        }
    }
}
