/* 
 * Author: YH CHIU
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;

namespace DailyReportTool
{
    public partial class Form1 : Form
    {
        private JObject config;
        private Dictionary<string, List<string>> processMap;
        private Dictionary<string, string> equipCategoryMap;

        // 資料結構定義
        private class EquipDetailedStat {
            public string Id;
            public double MaintHrs, KPIHrs, MaxHrs;
        }

        public Form1()
        {
            InitializeComponent();
            this.Text = "DailyReportTool v1.1.1";
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
                    if (config["Process"] != null) {
                        foreach (JProperty proc in config["Process"].Children<JProperty>())
                            processMap[proc.Name] = ((string)proc.Value["Equipments"]).Split(',').Select(e => e.Trim()).ToList();
                    }
                    if (config["Non-Common"] != null) {
                        foreach (JProperty type in config["Non-Common"].Children<JProperty>())
                            foreach (var eq in ((string)type.Value["Equipments"]).Split(',')) equipCategoryMap[eq.Trim()] = type.Name;
                    }
                }
            }
            catch (Exception ex) { Log("Config Error: " + ex.Message); }
        }

        private void Log(string m) { txtLog.AppendText($"{DateTime.Now:HH:mm:ss} {m}{Environment.NewLine}"); }

        // --- Navigation ---
        private void menuPareto_Click(object sender, EventArgs e) { pnlPareto.Visible = true; pnlBalance.Visible = false; }
        private void menuBalance_Click(object sender, EventArgs e) { pnlPareto.Visible = false; pnlBalance.Visible = true; }

        // --- Shared Helpers ---
        private bool IsBinaryExcel(string path) {
            try {
                byte[] header = new byte[4];
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    fs.Read(header, 0, 4);
                }
                return (header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0) || 
                       (header[0] == 0x50 && header[1] == 0x4B && header[2] == 0x03 && header[3] == 0x04);
            } catch { return false; }
        }

        private string GetCellValue(ICell c) { if (c == null) return ""; if (c.CellType == CellType.String) return c.StringCellValue; return c.ToString().Trim(); }

        // --- Pareto Logic ---
        private void btnSelectMaintenance_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel/HTML|*.xlsx;*.xls;*.html" })
                if (ofd.ShowDialog() == DialogResult.OK) { txtMaintenancePath.Text = ofd.FileName; Log("Selected Record: " + ofd.FileName); }
        }
        private void btnSelectConnection_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel/HTML|*.xlsx;*.xls;*.html" })
                if (ofd.ShowDialog() == DialogResult.OK) { txtConnectionPath.Text = ofd.FileName; Log("Selected KPI: " + ofd.FileName); }
        }
        private void btnClear_Click(object sender, EventArgs e) { txtMaintenancePath.Clear(); txtConnectionPath.Clear(); Log("Cleared."); }
        
        private void btnGenerate_Click(object sender, EventArgs e) {
            if (string.IsNullOrEmpty(txtMaintenancePath.Text) && string.IsNullOrEmpty(txtConnectionPath.Text)) return;
            try {
                Dictionary<string, double> maintMap = new Dictionary<string, double>();
                if (!string.IsNullOrEmpty(txtMaintenancePath.Text)) ProcessFile(txtMaintenancePath.Text, maintMap, true);
                
                Dictionary<string, double> kpiRatioMap = new Dictionary<string, double>();
                if (!string.IsNullOrEmpty(txtConnectionPath.Text)) ProcessKPIFile(txtConnectionPath.Text, kpiRatioMap);

                double totalSchedHours = (double)numParetoDataDays.Value * (double)numParetoDailyHours.Value;

                var allEquips = maintMap.Keys.Union(kpiRatioMap.Keys).Distinct();
                var detailedStats = new List<EquipDetailedStat>();

                foreach (var id in allEquips) {
                    double mH = maintMap.ContainsKey(id) ? maintMap[id] : 0;
                    double kH = kpiRatioMap.ContainsKey(id) ? kpiRatioMap[id] * totalSchedHours : 0;
                    detailedStats.Add(new EquipDetailedStat { Id = id, MaintHrs = mH, KPIHrs = kH, MaxHrs = Math.Max(mH, kH) });
                }
                detailedStats = detailedStats.OrderByDescending(x => x.MaxHrs).ToList();

                var stats = new List<EquipDetailedStat>();
                foreach (var kvp in processMap) {
                    double h = kvp.Value.Sum(eq => detailedStats.FirstOrDefault(d => d.Id == eq)?.MaxHrs ?? 0);
                    if (h > 0) stats.Add(new EquipDetailedStat { Id = kvp.Key, MaxHrs = h });
                }
                stats = stats.OrderByDescending(x => x.MaxHrs).ToList();

                GenerateParetoReport(stats, detailedStats, (int)numParetoTopX.Value, (double)numParetoThresholdY.Value, totalSchedHours);
            } catch (Exception ex) { Log("Error: " + ex.Message); }
        }

        private void ProcessFile(string path, Dictionary<string, double> hours, bool isMaintenance) {
            if (!IsBinaryExcel(path)) {
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); doc.Load(path, true);
                var trs = doc.DocumentNode.SelectNodes("//tr"); if (trs == null) return;
                foreach (var tr in trs.Skip(1)) {
                    var tds = tr.SelectNodes("td|th"); if (tds == null || tds.Count < 10) continue;
                    string id = System.Net.WebUtility.HtmlDecode(tds[2].InnerText).Trim();
                    if (isMaintenance) {
                        if (tds[6].InnerText.Contains("A-維修") && double.TryParse(tds[9].InnerText, out double h)) { if (!hours.ContainsKey(id)) hours[id] = 0; hours[id] += h; }
                    } else {
                        if (DateTime.TryParse(tds[7].InnerText, out DateTime s) && DateTime.TryParse(tds[8].InnerText, out DateTime ed)) {
                            double h = (ed - s).TotalHours; if (h > 0) { if (!hours.ContainsKey(id)) hours[id] = 0; hours[id] += h; }
                        }
                    }
                }
            } else {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    IWorkbook wb = WorkbookFactory.Create(fs); ISheet ws = wb.GetSheetAt(0);
                    for (int i = 1; i <= ws.LastRowNum; i++) {
                        IRow r = ws.GetRow(i); if (r == null) continue;
                        string id = GetCellValue(r.GetCell(1));
                        if (isMaintenance) {
                            if (GetCellValue(r.GetCell(5)).Contains("A-維修")) { double h = r.GetCell(8)?.NumericCellValue ?? 0; if (!hours.ContainsKey(id)) hours[id] = 0; hours[id] += h; }
                        } else {
                            DateTime? s = r.GetCell(6)?.DateCellValue; DateTime? ed = r.GetCell(7)?.DateCellValue;
                            if (s.HasValue && ed.HasValue) { double h = (ed.Value - s.Value).TotalHours; if (h > 0) { if (!hours.ContainsKey(id)) hours[id] = 0; hours[id] += h; } }
                        }
                    }
                }
            }
        }

        private void GenerateParetoReport(List<EquipDetailedStat> stats, List<EquipDetailedStat> equipStats, int topX, double thresholdY, double schedHours) {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = $"DailyReport_{DateTime.Now:yyyyMMdd}.xlsx" })
            if (sfd.ShowDialog() == DialogResult.OK) {
                IWorkbook wb = new XSSFWorkbook();
                // 1 & 2: 總機故
                ISheet s1 = wb.CreateSheet("總機故時數"); CreateDataSheet(s1, stats, "製程站點", false);
                ISheet s2 = wb.CreateSheet("總機故柏拉圖");
                ExcelChartHelper.CreateParetoChart(s2, s1.SheetName, stats.Count, "總機故柏拉圖", stats.Sum(x => x.MaxHrs));

                // 3: 單機機故時數 (新 - 含原始數據欄位)
                ISheet s3 = wb.CreateSheet("單機機故時數"); CreateDataSheet(s3, equipStats, "設備編號", true);

                // 4: 單機機故柏拉圖 (前X大)
                var topXData = equipStats.Take(topX).ToList();
                ISheet s4 = wb.CreateSheet($"單機機故柏拉圖(前{topX}大)");
                CreateDataSheet(s4, topXData, "設備編號", false); 
                ExcelChartHelper.CreateParetoChart(s4, s4.SheetName, topXData.Count, $"單機機故柏拉圖(前{topX}大)", topXData.Sum(x => x.MaxHrs));

                // 5: 單機機故柏拉圖 (Y%以上)
                var thresholdData = equipStats.Where(x => schedHours > 0 && (x.MaxHrs / schedHours) * 100 >= thresholdY).ToList();
                ISheet s5 = wb.CreateSheet($"單機機故柏拉圖({thresholdY}%以上)");
                CreateDataSheet(s5, thresholdData, "設備編號", false);
                ExcelChartHelper.CreateParetoChart(s5, s5.SheetName, thresholdData.Count, $"單機機故柏拉圖({thresholdY}%以上)", thresholdData.Sum(x => x.MaxHrs));

                using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create)) { wb.Write(fs); }
                Log("Report Success.");
                if (MessageBox.Show("報表產出成功，是否直接開啟檔案？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
            }
        }

        private void CreateDataSheet(ISheet s, List<EquipDetailedStat> d, string n, bool detailed) {
            IRow hRow = s.CreateRow(0); hRow.CreateCell(0).SetCellValue(n); hRow.CreateCell(1).SetCellValue("當機時數(小時)"); hRow.CreateCell(2).SetCellValue("累積百分比");
            if (detailed) { hRow.CreateCell(3).SetCellValue("維修"); hRow.CreateCell(4).SetCellValue("KPI"); }

            ICellStyle hS = s.Workbook.CreateCellStyle(); hS.DataFormat = s.Workbook.CreateDataFormat().GetFormat("0.00");
            ICellStyle pS = s.Workbook.CreateCellStyle(); pS.DataFormat = s.Workbook.CreateDataFormat().GetFormat("0%");
            double total = d.Sum(x => x.MaxHrs), cur = 0;
            for (int i = 0; i < d.Count; i++) {
                IRow r = s.CreateRow(i + 1); r.CreateCell(0).SetCellValue(d[i].Id);
                ICell c1 = r.CreateCell(1); c1.SetCellValue(d[i].MaxHrs); c1.CellStyle = hS;
                cur += d[i].MaxHrs; ICell c2 = r.CreateCell(2); c2.SetCellValue(total == 0 ? 0 : cur / total); c2.CellStyle = pS;
                if (detailed) {
                    ICell c3 = r.CreateCell(3); c3.SetCellValue(d[i].MaintHrs); c3.CellStyle = hS;
                    ICell c4 = r.CreateCell(4); c4.SetCellValue(d[i].KPIHrs); c4.CellStyle = hS;
                }
            }
            for (int i = 0; i < (detailed ? 5 : 3); i++) s.AutoSizeColumn(i);
        }

        // --- Capacity Balance Logic ---
        private class Balance_DataRow {
            public string Name, Category;
            public double TargetPcsPerDay, TheoreticalPcsPerDay, BreakdownPcsPerDay, EquipCount, WeightedAvgCT, WeightedAvgPass;
            public List<Balance_ProductInfo> Products = new List<Balance_ProductInfo>();
        }
        private class Balance_ProductInfo { public double CT, PcsPerLot, DailyLotDemand, PassCount; }

        private void btnImportIE_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel|*.xlsx;*.xls" })
                if (ofd.ShowDialog() == DialogResult.OK) { txtIEPath.Text = ofd.FileName; Log("Selected IE: " + ofd.FileName); }
        }

        private void btnBalanceSelectMaintenance_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel/HTML|*.xlsx;*.xls;*.html" })
                if (ofd.ShowDialog() == DialogResult.OK) { txtBalanceMaintenancePath.Text = ofd.FileName; Log("Balance Record: " + ofd.FileName); }
        }

        private void btnBalanceSelectConnection_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel|*.xlsx;*.xls" })
                if (ofd.ShowDialog() == DialogResult.OK) { txtBalanceConnectionPath.Text = ofd.FileName; Log("KPI File: " + ofd.FileName); }
        }

        private void btnGenerateBalance_Click(object sender, EventArgs e) {
            if (string.IsNullOrEmpty(txtIEPath.Text)) return;
            try {
                using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = $"BalanceReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx" }) {
                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    Dictionary<string, double> equipHours = new Dictionary<string, double>();
                    if (!string.IsNullOrEmpty(txtBalanceMaintenancePath.Text)) ProcessFile(txtBalanceMaintenancePath.Text, equipHours, true);

                    Dictionary<string, double> equipKPI = new Dictionary<string, double>();
                    if (!string.IsNullOrEmpty(txtBalanceConnectionPath.Text)) ProcessKPIFile(txtBalanceConnectionPath.Text, equipKPI);

                    IWorkbook wb = new XSSFWorkbook();
                    List<Balance_DataRow> allData = Balance_ReadIEData(txtIEPath.Text);
                    Balance_PerformCalculations(allData, (double)numDailyHours.Value, (int)numDataDays.Value, equipHours, equipKPI);

                    var groups = allData.GroupBy(x => x.Category).OrderBy(g => g.Key == "Common" ? 0 : 1);

                    foreach (var group in groups) {
                        string catName = group.Key;
                        string displayCatName = catName == "Common" ? "共用設備" : catName;
                        var groupData = group.ToList();

                        ISheet sData = wb.CreateSheet($"數據({catName})");
                        Balance_RenderData(sData, groupData, (int)numDataDays.Value);

                        ISheet sChart = wb.CreateSheet($"平衡圖({catName})");
                        ExcelChartHelper.CreateCapacityBalanceChart(sChart, sData.SheetName, groupData.Count, $"產能平衡圖 - {displayCatName}", 
                            groupData.Max(r => Math.Max(r.TargetPcsPerDay, r.TheoreticalPcsPerDay)) * (int)numDataDays.Value * 1.1);
                    }

                    using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create)) { wb.Write(fs); }
                    Log("Balance Success!");
                    if (MessageBox.Show("平衡圖產出成功，是否直接開啟檔案？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                        System.Diagnostics.Process.Start(sfd.FileName);
                    }
                }
            } catch (Exception ex) { Log("Balance Error: " + ex.Message); MessageBox.Show("Balance Error: " + ex.Message); }
        }

        private void ProcessKPIFile(string path, Dictionary<string, double> kpiMap) {
            if (!IsBinaryExcel(path)) {
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); doc.Load(path, true);
                var trs = doc.DocumentNode.SelectNodes("//tr"); if (trs == null) return;
                foreach (var tr in trs.Skip(1)) {
                    var tds = tr.SelectNodes("td|th"); if (tds == null || tds.Count < 12) continue;
                    string id = System.Net.WebUtility.HtmlDecode(tds[0].InnerText).Trim();
                    if (string.IsNullOrEmpty(id) || id.Contains("合計") || id.ToUpper().Contains("TOTAL")) continue;
                    string valStr = tds[11].InnerText.Replace("%", "").Replace(",", "").Trim();
                    if (double.TryParse(valStr, out double d)) { kpiMap[id] = d / 100.0; } 
                }
            } else {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    IWorkbook wb = WorkbookFactory.Create(fs); ISheet ws = wb.GetSheetAt(0);
                    for (int i = 2; i <= ws.LastRowNum; i++) {
                        IRow r = ws.GetRow(i); if (r == null) continue;
                        string id = GetCellValue(r.GetCell(0)).Trim();
                        if (string.IsNullOrEmpty(id) || id.Contains("合計") || id.ToUpper().Contains("TOTAL")) continue;
                        double val = 0; ICell cL = r.GetCell(11);
                        if (cL != null) {
                            if (cL.CellType == CellType.Numeric) val = cL.NumericCellValue;
                            else if (double.TryParse(cL.ToString().Replace("%", "").Replace(",", ""), out double d)) { val = d / 100.0; }
                        }
                        kpiMap[id] = val;
                    }
                }
            }
        }

        private List<Balance_DataRow> Balance_ReadIEData(string p) {
            List<Balance_DataRow> rows = new List<Balance_DataRow>();
            using (FileStream fs = new FileStream(p, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                IWorkbook wb = WorkbookFactory.Create(fs); IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
                ISheet ws = wb.GetSheet("IE主表"); int prodCount = 0; IRow r3 = ws.GetRow(2);
                for (int c = 1; c < r3.LastCellNum; c++) { if (string.IsNullOrEmpty(r3.GetCell(c)?.ToString())) break; prodCount++; }
                for (int r = 14; r <= ws.LastRowNum; r++) {
                    IRow row = ws.GetRow(r); if (row == null || Balance_IsMerged(ws, r, 0)) continue;
                    string n = row.GetCell(0)?.ToString().Trim(); if (string.IsNullOrEmpty(n) || n.Contains("專線") || n.Contains("C/T")) continue;
                    string cat = "Common";
                    if (equipCategoryMap.ContainsKey(n)) cat = equipCategoryMap[n];
                    if (n.Contains("NiAu") || n.Contains("uPOL-DB") || n.Contains("uPOL-TP")) cat = "uPOL";
                    else if (n.Contains("De-Flux") || n.Contains("Molding") || n.Contains("Laser Marking") || n.Contains("uBMU-TP") || n.Contains("uBMU-DB")) cat = "uBMU";
                    Balance_DataRow dr = new Balance_DataRow { Name = n, Category = cat };
                    for (int j = 0; j < prodCount; j++) dr.Products.Add(new Balance_ProductInfo { CT = Balance_GetCleanNum(row.GetCell(1+j*4), eval), PcsPerLot = Balance_GetCleanNum(row.GetCell(2+j*4), eval), DailyLotDemand = Balance_GetCleanNum(row.GetCell(3+j*4), eval), PassCount = Balance_GetCleanNum(row.GetCell(4+j*4), eval) });
                    dr.WeightedAvgCT = Balance_GetCleanNum(row.GetCell(9), eval); // Col J
                    dr.TargetPcsPerDay = Balance_GetCleanNum(row.GetCell(10), eval); // Col K
                    dr.EquipCount = Balance_GetCleanNum(row.GetCell(13), eval); // Col N
                    rows.Add(dr);
                }
            }
            return rows;
        }

        private void Balance_PerformCalculations(List<Balance_DataRow> rows, double hr, int days, Dictionary<string, double> equipHours, Dictionary<string, double> equipKPI) {
            foreach (var r in rows) {
                // 產能計算直接使用從 Excel 讀取的 J 欄 (WeightedAvgCT) 與 N 欄 (EquipCount)
                double theory = (r.EquipCount * hr * 3600) / (r.WeightedAvgCT > 0 ? r.WeightedAvgCT : 1);
                r.TheoreticalPcsPerDay = theory;

                // 仍然計算加權良率，以供機故損失或其他邏輯參考（如果需要）
                double tP = 0, tB = 0, wP = 0;
                foreach (var p in r.Products) { 
                    double b = p.DailyLotDemand * p.PcsPerLot; 
                    tB += b; 
                    tP += (b * p.PassCount); 
                    wP += (p.PassCount * b); 
                }
                r.WeightedAvgPass = (tB > 0) ? (wP / tB) : 1;

                if (processMap.ContainsKey(r.Name)) {
                    double siteTotalBreakdownPcs = 0;
                    foreach (var eqId in processMap[r.Name]) {
                        double maintenancePcs = 0;
                        if (equipHours.ContainsKey(eqId) && hr > 0) 
                        {
                            double eqTheoreticalPcsPerDay = r.TheoreticalPcsPerDay / (r.EquipCount > 0 ? r.EquipCount : 1);
                            maintenancePcs = (equipHours[eqId] / hr) * eqTheoreticalPcsPerDay;
                        }
                        double kpiRatio = equipKPI.ContainsKey(eqId) ? equipKPI[eqId] : 0;
                        double kpiPcs = kpiRatio * (r.TheoreticalPcsPerDay / (r.EquipCount > 0 ? r.EquipCount : 1));
                        siteTotalBreakdownPcs += Math.Max(maintenancePcs, kpiPcs);
                    }
                    r.BreakdownPcsPerDay = siteTotalBreakdownPcs;
                }
            }
        }

        private void Balance_RenderData(ISheet s, List<Balance_DataRow> rows, int d) {
            string[] heads = { "機台/站別", "目標片次", "設備產能", "機故損失", "產速損失" };
            for (int j = 0; j < 5; j++) s.CreateRow(20 + j).CreateCell(0).SetCellValue(heads[j]);
            double maxTarget = rows.Count > 0 ? rows.Max(r => r.TargetPcsPerDay * d) : 0;
            long roundedMaxTarget = (long)Math.Round(maxTarget);
            for (int i = 0; i < rows.Count; i++) {
                int c = i + 1; s.GetRow(20).CreateCell(c).SetCellValue(rows[i].Name);
                s.GetRow(21).CreateCell(c).SetCellValue(roundedMaxTarget);
                s.GetRow(22).CreateCell(c).SetCellValue(Math.Round(rows[i].TheoreticalPcsPerDay * d));
                s.GetRow(23).CreateCell(c).SetCellValue(Math.Round(rows[i].BreakdownPcsPerDay));
                s.GetRow(24).CreateCell(c).SetCellValue(0);
            }
        }

        private bool Balance_IsMerged(ISheet s, int r, int c) { for (int i = 0; i < s.NumMergedRegions; i++) { if (s.GetMergedRegion(i).IsInRange(r, c)) return true; } return false; }
        
        private double Balance_GetCleanNum(ICell c, IFormulaEvaluator e) {
            if (c == null) return 0;
            try {
                if (c.CellType == CellType.Formula) {
                    try { CellValue v = e.Evaluate(c); return (v != null) ? v.NumberValue : 0; }
                    catch { return c.NumericCellValue; }
                }
                if (c.CellType == CellType.Numeric) return c.NumericCellValue;
                if (double.TryParse(c.ToString().Replace("%", "").Replace(",", ""), out double d)) return d / (c.ToString().Contains("%") ? 100.0 : 1.0);
            } catch { }
            return 0;
        }
    }
}
