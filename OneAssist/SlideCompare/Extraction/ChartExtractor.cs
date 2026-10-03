using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OneAssist.SlideCompare.Extraction
{
    /// <summary>
    /// Safely extracts chart metadata, titles, and series from PowerPoint charts.
    /// </summary>
    public static class ChartExtractor
    {
        public static bool ExtractChart(
            PowerPoint.Shape shape,
            out string chartTitle,
            out List<string> seriesNames,
            out List<string> categories)
        {
            chartTitle = string.Empty;
            seriesNames = new List<string>();
            categories = new List<string>();

            if (shape == null) return false;

            try
            {
                if (shape.HasChart == Office.MsoTriState.msoTrue)
                {
                    var chart = shape.Chart;
                    try
                    {
                        if (chart.HasTitle)
                        {
                            chartTitle = chart.ChartTitle.Text ?? string.Empty;
                        }
                    }
                    catch { }

                    try
                    {
                        var sc = (PowerPoint.SeriesCollection)chart.SeriesCollection();
                        int sCount = sc.Count;
                        for (int i = 1; i <= sCount; i++)
                        {
                            try
                            {
                                PowerPoint.Series s = sc.Item(i);
                                seriesNames.Add(s.Name ?? $"Series {i}");
                            }
                            catch { }
                        }
                    }
                    catch { }

                    return true;
                }
            }
            catch { }

            return false;
        }
    }
}
