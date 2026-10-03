using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using OneAssist.Assistant.Core;

namespace OneAssist.Assistant.Services
{
    public static class ReportExportService
    {
        public static string ExportToHtml(QualityReport report, string presentationName = "Presentation")
        {
            if (report == null) return "<html><body><h1>No Report Generated</h1></body></html>";

            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset='utf-8'>");
            sb.AppendLine($"<title>Quality Audit: {presentationName}</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; margin: 30px; background: #f8fafc; color: #1e293b; }");
            sb.AppendLine(".card { background: white; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); padding: 24px; margin-bottom: 24px; }");
            sb.AppendLine(".header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #e2e8f0; padding-bottom: 16px; margin-bottom: 20px; }");
            sb.AppendLine(".score-badge { font-size: 28px; font-weight: bold; padding: 8px 18px; border-radius: 9999px; background: #e0f2fe; color: #0284c7; }");
            sb.AppendLine(".stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }");
            sb.AppendLine(".stat-box { background: #f1f5f9; padding: 16px; border-radius: 6px; text-align: center; }");
            sb.AppendLine(".stat-val { font-size: 24px; font-weight: bold; margin-bottom: 4px; }");
            sb.AppendLine(".stat-lbl { font-size: 12px; color: #64748b; text-transform: uppercase; }");
            sb.AppendLine("table { width: 100%; border-collapse: collapse; margin-top: 16px; }");
            sb.AppendLine("th, td { text-align: left; padding: 12px; border-bottom: 1px solid #e2e8f0; font-size: 13px; }");
            sb.AppendLine("th { background: #f8fafc; color: #475569; }");
            sb.AppendLine(".badge { padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; }");
            sb.AppendLine(".badge-Error { background: #fee2e2; color: #dc2626; }");
            sb.AppendLine(".badge-Warning { background: #fef3c7; color: #d97706; }");
            sb.AppendLine(".badge-Info { background: #e0f2fe; color: #0284c7; }");
            sb.AppendLine("</style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");

            sb.AppendLine("<div class='card'>");
            sb.AppendLine("<div class='header'>");
            sb.AppendLine("<div>");
            sb.AppendLine($"<h1>Presentation Quality Audit Report</h1>");
            sb.AppendLine($"<p style='color:#64748b; margin:0;'>Analyzed: {report.AnalyzedAt:yyyy-MM-dd HH:mm:ss} | Scope: {report.Scope}</p>");
            sb.AppendLine("</div>");
            sb.AppendLine($"<div class='score-badge'>{report.HealthScore:F0}% Health</div>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='stats-grid'>");
            sb.AppendLine($"<div class='stat-box'><div class='stat-val' style='color:#dc2626;'>{report.ErrorCount}</div><div class='stat-lbl'>Errors</div></div>");
            sb.AppendLine($"<div class='stat-box'><div class='stat-val' style='color:#d97706;'>{report.WarningCount}</div><div class='stat-lbl'>Warnings</div></div>");
            sb.AppendLine($"<div class='stat-box'><div class='stat-val' style='color:#16a34a;'>{report.PassedCount}</div><div class='stat-lbl'>Passed</div></div>");
            sb.AppendLine($"<div class='stat-box'><div class='stat-val' style='color:#0284c7;'>{report.TotalSlidesAnalyzed}</div><div class='stat-lbl'>Slides Analyzed</div></div>");
            sb.AppendLine("</div>");

            sb.AppendLine("<h2>Detailed Findings</h2>");
            sb.AppendLine("<table>");
            sb.AppendLine("<tr><th>Slide</th><th>Severity</th><th>Category</th><th>Rule</th><th>Problem</th><th>Expected</th><th>Actual</th><th>Fix Available</th></tr>");

            foreach (var iss in report.Issues.OrderBy(i => i.SlideIndex).ThenBy(i => i.Severity))
            {
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td>Slide {iss.SlideIndex}</td>");
                sb.AppendLine($"<td><span class='badge badge-{iss.Severity}'>{iss.Severity}</span></td>");
                sb.AppendLine($"<td>{iss.Category}</td>");
                sb.AppendLine($"<td><strong>{iss.RuleTitle}</strong></td>");
                sb.AppendLine($"<td>{iss.Problem}</td>");
                sb.AppendLine($"<td>{iss.Expected}</td>");
                sb.AppendLine($"<td>{iss.Actual}</td>");
                sb.AppendLine($"<td>{(iss.CanAutoFix ? "Fix Available" : "Manual Review")}</td>");
                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</table>");
            sb.AppendLine("</div>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            return sb.ToString();
        }

        public static string ExportToCsv(QualityReport report)
        {
            if (report == null) return string.Empty;
            var sb = new StringBuilder();
            sb.AppendLine("Slide,Severity,Category,Rule,Problem,Expected,Actual,CanAutoFix");

            foreach (var iss in report.Issues)
            {
                string p = (iss.Problem ?? "").Replace(",", ";").Replace("\r", " ").Replace("\n", " ");
                string e = (iss.Expected ?? "").Replace(",", ";");
                string a = (iss.Actual ?? "").Replace(",", ";");
                sb.AppendLine($"{iss.SlideIndex},{iss.Severity},{iss.Category},\"{iss.RuleTitle}\",\"{p}\",\"{e}\",\"{a}\",{iss.CanAutoFix}");
            }

            return sb.ToString();
        }

        public static string ExportToJson(QualityReport report, string presentationName = "Presentation")
        {
            if (report == null) return "{}";

            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine($"  \"presentation\": \"{EscapeJson(presentationName)}\",");
            sb.AppendLine($"  \"analyzedAt\": \"{report.AnalyzedAt:O}\",");
            sb.AppendLine($"  \"healthScore\": {report.HealthScore:F1},");
            sb.AppendLine($"  \"summary\": {{");
            sb.AppendLine($"    \"errors\": {report.ErrorCount},");
            sb.AppendLine($"    \"warnings\": {report.WarningCount},");
            sb.AppendLine($"    \"passed\": {report.PassedCount},");
            sb.AppendLine($"    \"fixed\": {report.FixedCount},");
            sb.AppendLine($"    \"totalSlidesAnalyzed\": {report.TotalSlidesAnalyzed},");
            sb.AppendLine($"    \"totalShapesAnalyzed\": {report.TotalShapesAnalyzed}");
            sb.AppendLine($"  }},");
            sb.AppendLine("  \"findings\": [");

            var issues = report.Issues.OrderBy(i => i.SlideIndex).ToList();
            for (int i = 0; i < issues.Count; i++)
            {
                var iss = issues[i];
                sb.AppendLine("    {");
                sb.AppendLine($"      \"slide\": {iss.SlideIndex},");
                sb.AppendLine($"      \"severity\": \"{iss.Severity}\",");
                sb.AppendLine($"      \"category\": \"{iss.Category}\",");
                sb.AppendLine($"      \"ruleId\": \"{EscapeJson(iss.RuleId)}\",");
                sb.AppendLine($"      \"ruleTitle\": \"{EscapeJson(iss.RuleTitle)}\",");
                sb.AppendLine($"      \"problem\": \"{EscapeJson(iss.Problem)}\",");
                sb.AppendLine($"      \"expected\": \"{EscapeJson(iss.Expected)}\",");
                sb.AppendLine($"      \"actual\": \"{EscapeJson(iss.Actual)}\",");
                sb.AppendLine($"      \"canAutoFix\": {iss.CanAutoFix.ToString().ToLowerInvariant()},");
                sb.AppendLine($"      \"safetyLevel\": \"{iss.SafetyLevel}\"");
                sb.Append("    }");
                if (i < issues.Count - 1) sb.Append(",");
                sb.AppendLine();
            }

            sb.AppendLine("  ]");
            sb.AppendLine("}");

            return sb.ToString();
        }

        public static string ExportToExcelXml(QualityReport report, string presentationName = "Presentation")
        {
            if (report == null) return string.Empty;

            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\"?>");
            sb.AppendLine("<?mso-application progid=\"Excel.Sheet\"?>");
            sb.AppendLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            sb.AppendLine(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
            sb.AppendLine(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"");
            sb.AppendLine(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
            sb.AppendLine(" <Styles>");
            sb.AppendLine("  <Style ss:ID=\"Header\"><Font ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#0F172A\" ss:Pattern=\"Solid\"/></Style>");
            sb.AppendLine("  <Style ss:ID=\"Error\"><Interior ss:Color=\"#FEE2E2\" ss:Pattern=\"Solid\"/><Font ss:Color=\"#DC2626\" ss:Bold=\"1\"/></Style>");
            sb.AppendLine("  <Style ss:ID=\"Warning\"><Interior ss:Color=\"#FEF3C7\" ss:Pattern=\"Solid\"/><Font ss:Color=\"#D97706\" ss:Bold=\"1\"/></Style>");
            sb.AppendLine(" </Styles>");
            sb.AppendLine(" <Worksheet ss:Name=\"Quality Report\">");
            sb.AppendLine("  <Table>");
            sb.AppendLine("   <Row ss:StyleID=\"Header\">");
            sb.AppendLine("    <Cell><Data ss:Type=\"String\">Slide</Data></Cell>");
            sb.AppendLine("    <Cell><Data ss:Type=\"String\">Severity</Data></Cell>");
            sb.AppendLine("    <Cell><Data ss:Type=\"String\">Category</Data></Cell>");
            sb.AppendLine("    <Cell><Data ss:Type=\"String\">Rule</Data></Cell>");
            sb.AppendLine("    <Cell><Data ss:Type=\"String\">Problem</Data></Cell>");
            sb.AppendLine("    <Cell><Data ss:Type=\"String\">Expected</Data></Cell>");
            sb.AppendLine("    <Cell><Data ss:Type=\"String\">Actual</Data></Cell>");
            sb.AppendLine("    <Cell><Data ss:Type=\"String\">Fix Available</Data></Cell>");
            sb.AppendLine("   </Row>");

            foreach (var iss in report.Issues)
            {
                string style = iss.Severity == IssueSeverity.Error ? " ss:StyleID=\"Error\"" : (iss.Severity == IssueSeverity.Warning ? " ss:StyleID=\"Warning\"" : "");
                sb.AppendLine($"   <Row{style}>");
                sb.AppendLine($"    <Cell><Data ss:Type=\"Number\">{iss.SlideIndex}</Data></Cell>");
                sb.AppendLine($"    <Cell><Data ss:Type=\"String\">{EscapeXml(iss.Severity.ToString())}</Data></Cell>");
                sb.AppendLine($"    <Cell><Data ss:Type=\"String\">{EscapeXml(iss.Category.ToString())}</Data></Cell>");
                sb.AppendLine($"    <Cell><Data ss:Type=\"String\">{EscapeXml(iss.RuleTitle)}</Data></Cell>");
                sb.AppendLine($"    <Cell><Data ss:Type=\"String\">{EscapeXml(iss.Problem)}</Data></Cell>");
                sb.AppendLine($"    <Cell><Data ss:Type=\"String\">{EscapeXml(iss.Expected)}</Data></Cell>");
                sb.AppendLine($"    <Cell><Data ss:Type=\"String\">{EscapeXml(iss.Actual)}</Data></Cell>");
                sb.AppendLine($"    <Cell><Data ss:Type=\"String\">{(iss.CanAutoFix ? "Yes" : "No")}</Data></Cell>");
                sb.AppendLine("   </Row>");
            }

            sb.AppendLine("  </Table>");
            sb.AppendLine(" </Worksheet>");
            sb.AppendLine("</Workbook>");

            return sb.ToString();
        }

        public static string ExportToTextReport(QualityReport report, string presentationName = "Presentation")
        {
            if (report == null) return "No Report";

            var sb = new StringBuilder();
            sb.AppendLine("Presentation Quality Report");
            sb.AppendLine($"Presentation: {presentationName}");
            sb.AppendLine();
            sb.AppendLine("Summary:");
            sb.AppendLine($"Errors: {report.ErrorCount}");
            sb.AppendLine($"Warnings: {report.WarningCount}");
            sb.AppendLine($"Passed: {report.PassedCount}");
            sb.AppendLine();
            sb.AppendLine("Detailed Findings:");

            foreach (var iss in report.Issues.OrderBy(i => i.SlideIndex))
            {
                sb.AppendLine($"Slide {iss.SlideIndex}");
                sb.AppendLine($"{iss.Category}");
                sb.AppendLine($"{iss.RuleTitle}");
                if (!string.IsNullOrWhiteSpace(iss.Expected))
                    sb.AppendLine($"Expected:\n{iss.Expected}");
                if (!string.IsNullOrWhiteSpace(iss.Actual))
                    sb.AppendLine($"Actual:\n{iss.Actual}");
                sb.AppendLine(iss.CanAutoFix ? "[Fix available]" : "[Manual review required]");
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private static string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
        }

        private static string EscapeXml(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
        }
    }
}
