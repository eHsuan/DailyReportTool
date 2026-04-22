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
                    Log("Config loaded successfully.");
                }
                else
                {
                    Log("Error: config.json not found.");
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

                // 1. Process Maintenance Record
                if (!string.IsNullOrEmpty(txtMaintenancePath.Text) && File.Exists(txtMaintenancePath.Text))
                {
                    Log("Processing Maintenance Record...");
                    ProcessMaintenanceFile(txtMaintenancePath.Text, equipmentHours);
                }

                // 2. Process Connection Data
                if (!string.IsNullOrEmpty(txtConnectionPath.Text) && File.Exists(txtConnectionPath.Text))
                {
                    Log("Processing Connection Data...");
                    ProcessConnectionFile(txtConnectionPath.Text, equipmentHours);
                }

                // 3. Aggregate by Process
                Log("Aggregating data by Process...");
                var processStats = new List<Tuple<string, double>>();
                
                foreach (var kvp in processMap)
                {
                    string processName = kvp.Key;
                    double totalHours = 0;
                    foreach (var eq in kvp.Value)
                    {
                        if (equipmentHours.ContainsKey(eq))
                        {
                            totalHours += equipmentHours[eq];
                        }
                    }
                    if (totalHours > 0)
                        processStats.Add(Tuple.Create(processName, totalHours));
                }

                // Sort descending
                processStats.Sort((a, b) => b.Item2.CompareTo(a.Item2));

                // 4. Generate Output Excel
                Log("Generating Report...");
                GenerateReport(processStats, equipmentHours);
                
                Log("Report Generation Complete.");
            }
            catch (Exception ex)
            {
                Log("Error generating report: " + ex.Message);
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool IsHtmlFile(string path)
        {
            try
            {
                using (StreamReader reader = new StreamReader(path))
                {
                    char[] buffer = new char[512];
                    int charsRead = reader.Read(buffer, 0, buffer.Length);
                    string content = new string(buffer, 0, charsRead).ToLower();
                    return content.Contains("<html") || content.Contains("<table") || content.Contains("<div") || content.Contains("<meta");
                }
            }
            catch
            {
                return false;
            }
        }

        private void ProcessMaintenanceFile(string path, Dictionary<string, double> equipmentHours)
        {
            if (IsHtmlFile(path))
            {
                ProcessMaintenanceHtml(path, equipmentHours);
            }
            else
            {
                ProcessMaintenanceExcel(path, equipmentHours);
            }
        }

        private void ProcessMaintenanceExcel(string path, Dictionary<string, double> equipmentHours)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook workbook = WorkbookFactory.Create(fs);
                ISheet sheet = workbook.GetSheetAt(0); 

                for (int i = 1; i <= sheet.LastRowNum; i++) 
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    string eqId = GetCellValue(row.GetCell(1));
                    string type = GetCellValue(row.GetCell(5));
                    
                    if (type != null && type.Contains("A-維修"))
                    {
                        double hours = GetNumericValue(row.GetCell(8));
                        if (!equipmentHours.ContainsKey(eqId))
                        {
                            equipmentHours[eqId] = 0;
                        }
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

            var rows = doc.DocumentNode.SelectNodes("//tr");
            if (rows == null) return;

            foreach (var row in rows.Skip(1))
            {
                var cells = row.SelectNodes("td|th");
                if (cells == null || cells.Count < 10) continue; 
                
                string eqId = System.Net.WebUtility.HtmlDecode(cells[2].InnerText).Trim();
                string type = System.Net.WebUtility.HtmlDecode(cells[6].InnerText).Trim();
                
                if (type != null && type.Contains("A-維修"))
                {
                    if (double.TryParse(cells[9].InnerText, out double hours))
                    {
                        if (!equipmentHours.ContainsKey(eqId))
                        {
                            equipmentHours[eqId] = 0;
                        }
                        equipmentHours[eqId] += hours;
                    }
                }
            }
        }

        private void ProcessConnectionFile(string path, Dictionary<string, double> equipmentHours)
        {
            if (IsHtmlFile(path))
            {
                ProcessConnectionHtml(path, equipmentHours);
            }
            else
            {
                ProcessConnectionExcel(path, equipmentHours);
            }
        }

        private void ProcessConnectionExcel(string path, Dictionary<string, double> equipmentHours)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook workbook = WorkbookFactory.Create(fs);
                ISheet sheet = workbook.GetSheetAt(0);

                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    string eqId = GetCellValue(row.GetCell(1));
                    DateTime? start = row.GetCell(6)?.DateCellValue;
                    DateTime? end = row.GetCell(7)?.DateCellValue;

                    if (start.HasValue && end.HasValue)
                    {
                        double hours = (end.Value - start.Value).TotalHours;
                        if (hours < 0) hours = 0; 

                        if (!equipmentHours.ContainsKey(eqId))
                        {
                            equipmentHours[eqId] = 0;
                        }
                        equipmentHours[eqId] += hours;
                    }
                }
                workbook.Close();
            }
        }

        private void ProcessConnectionHtml(string path, Dictionary<string, double> equipmentHours)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(path, Encoding.UTF8);

            var rows = doc.DocumentNode.SelectNodes("//tr");
            if (rows == null) return;

            foreach (var row in rows.Skip(1))
            {
                var cells = row.SelectNodes("td|th");
                if (cells == null || cells.Count < 9) continue; 

                string eqId = System.Net.WebUtility.HtmlDecode(cells[2].InnerText).Trim();
                if (DateTime.TryParse(cells[7].InnerText, out DateTime start) && DateTime.TryParse(cells[8].InnerText, out DateTime end))
                {
                    double hours = (end - start).TotalHours;
                    if (hours > 0)
                    {
                        if (!equipmentHours.ContainsKey(eqId))
                        {
                            equipmentHours[eqId] = 0;
                        }
                        equipmentHours[eqId] += hours;
                    }
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
            if (cell.CellType == CellType.String)
            {
                double val;
                if (double.TryParse(cell.StringCellValue, out val)) return val;
            }
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

                // Sheet 1: Process Totals
                ISheet sheet1 = workbook.CreateSheet("總機故時數");
                CreateDataSheet(sheet1, processStats, "製程站點");

                // Sheet 2: Process Pareto
                ISheet sheet2 = workbook.CreateSheet("總機故柏拉圖");
                double totalHours = processStats.Sum(x => x.Item2);
                ExcelChartHelper.CreateParetoChart(sheet2, sheet1.SheetName, processStats.Count, "總機故柏拉圖", totalHours);

                // Top 3 Processes
                int topCount = Math.Min(3, processStats.Count);
                for (int i = 0; i < topCount; i++)
                {
                    string processName = processStats[i].Item1;
                    var eqs = new List<Tuple<string, double>>();
                    if (processMap.ContainsKey(processName))
                    {
                        foreach (var eq in processMap[processName])
                        {
                            if (equipmentHours.ContainsKey(eq) && equipmentHours[eq] > 0)
                            {
                                eqs.Add(Tuple.Create(eq, equipmentHours[eq]));
                            }
                        }
                    }
                    eqs.Sort((a, b) => b.Item2.CompareTo(a.Item2));

                    string snD = processName + "機故時數";
                    string snC = processName + "機故柏拉圖";
                    if (snD.Length > 31) snD = snD.Substring(0, 31);
                    if (snC.Length > 31) snC = snC.Substring(0, 31);

                    ISheet sheetData = workbook.CreateSheet(snD);
                    CreateDataSheet(sheetData, eqs, "設備編號");

                    ISheet sheetChart = workbook.CreateSheet(snC);
                    ExcelChartHelper.CreateParetoChart(sheetChart, sheetData.SheetName, eqs.Count, processName + "柏拉圖", eqs.Sum(x => x.Item2));
                }

                using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }
                Log($"Report saved to {sfd.FileName}");
                MessageBox.Show("Report generated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CreateDataSheet(ISheet sheet, List<Tuple<string, double>> data, string nameHeader)
        {
            IRow header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue(nameHeader);
            header.CreateCell(1).SetCellValue("當機時數(小時)");
            header.CreateCell(2).SetCellValue("累積百分比");

            ICellStyle hourStyle = sheet.Workbook.CreateCellStyle();
            hourStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat("0.00");
            ICellStyle pctStyle = sheet.Workbook.CreateCellStyle();
            pctStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat("0%");

            double total = data.Sum(x => x.Item2);
            double currentSum = 0;

            for (int i = 0; i < data.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(data[i].Item1);
                ICell c1 = row.CreateCell(1);
                c1.SetCellValue(data[i].Item2);
                c1.CellStyle = hourStyle;
                currentSum += data[i].Item2;
                ICell c2 = row.CreateCell(2);
                c2.SetCellValue(total == 0 ? 0 : (currentSum / total));
                c2.CellStyle = pctStyle;
            }
        }
    }
}
