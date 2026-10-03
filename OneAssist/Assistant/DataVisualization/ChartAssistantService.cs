using System;
using System.Drawing;
using OneAssist.Assistant.Services;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.DataVisualization
{
    public static class ChartAssistantService
    {
        public static string FormatDataLabel(double value, string formatType)
        {
            switch (formatType?.ToLower())
            {
                case "currency_millions":
                case "millions":
                    if (Math.Abs(value) >= 1_000_000)
                        return $"${value / 1_000_000:F1}M";
                    if (Math.Abs(value) >= 1_000)
                        return $"${value / 1_000:F1}K";
                    return $"${value:F0}";

                case "percentage":
                case "percent":
                    return $"{value * 100:F0}%";

                case "integer":
                    return $"{Math.Round(value):F0}";

                default:
                    return value.ToString("G");
            }
        }

        public static bool CleanChart(PowerPoint.Shape chartShape, Color brandPrimary, Color brandSecondary)
        {
            if (chartShape == null || !chartShape.HasChart.Equals(Office.MsoTriState.msoTrue)) return false;

            try
            {
                var chart = chartShape.Chart;

                // Remove major clutter gridlines if present
                try
                {
                    PowerPoint.Axis axis = chart.Axes(PowerPoint.XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlPrimary) as PowerPoint.Axis;
                    if (axis != null && axis.HasMajorGridlines)
                    {
                        axis.MajorGridlines.Format.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                }
                catch { }

                // Apply brand colors to series
                try
                {
                    PowerPoint.SeriesCollection seriesCol = chart.SeriesCollection() as PowerPoint.SeriesCollection;
                    if (seriesCol != null && seriesCol.Count > 0)
                    {
                        PowerPoint.Series s1 = seriesCol.Item(1) as PowerPoint.Series;
                        if (s1 != null)
                        {
                            s1.Format.Fill.Solid();
                            s1.Format.Fill.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(brandPrimary);
                        }

                        if (seriesCol.Count > 1)
                        {
                            PowerPoint.Series s2 = seriesCol.Item(2) as PowerPoint.Series;
                            if (s2 != null)
                            {
                                s2.Format.Fill.Solid();
                                s2.Format.Fill.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(brandSecondary);
                            }
                        }
                    }
                }
                catch { }

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
