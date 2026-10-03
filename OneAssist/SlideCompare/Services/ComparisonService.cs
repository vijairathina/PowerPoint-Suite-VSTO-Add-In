using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OneAssist.SlideCompare.Comparison;
using OneAssist.SlideCompare.Extraction;
using OneAssist.SlideCompare.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OneAssist.SlideCompare.Services
{
    /// <summary>
    /// Service that coordinates opening source file, extracting COM models, running comparison, and exporting results.
    /// </summary>
    public class ComparisonService
    {
        public async Task<PresentationComparisonResult> RunComparisonAsync(
            PowerPoint.Application app,
            PowerPoint.Presentation destinationPresentation,
            string sourceFilePath,
            CompareSettings settings,
            IProgress<Tuple<int, string>> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (app == null) throw new ArgumentNullException(nameof(app));
            if (destinationPresentation == null) throw new ArgumentNullException(nameof(destinationPresentation));
            if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
                throw new FileNotFoundException("Source presentation file not found.", sourceFilePath);

            settings = settings ?? new CompareSettings();

            progress?.Report(Tuple.Create(5, "Reading destination presentation..."));
            var destSnapshot = PresentationExtractor.ExtractPresentation(destinationPresentation, progress, cancellationToken);

            progress?.Report(Tuple.Create(25, "Opening source presentation in read-only mode..."));
            PowerPoint.Presentation sourcePresentation = null;
            PresentationSnapshot sourceSnapshot = null;

            try
            {
                sourcePresentation = app.Presentations.Open(
                    sourceFilePath,
                    ReadOnly: Office.MsoTriState.msoTrue,
                    Untitled: Office.MsoTriState.msoFalse,
                    WithWindow: Office.MsoTriState.msoFalse);

                progress?.Report(Tuple.Create(40, "Extracting source presentation snapshot..."));
                sourceSnapshot = PresentationExtractor.ExtractPresentation(sourcePresentation, progress, cancellationToken);
            }
            finally
            {
                if (sourcePresentation != null)
                {
                    try
                    {
                        sourcePresentation.Close();
                    }
                    catch { }
                    finally
                    {
                        try { Marshal.ReleaseComObject(sourcePresentation); } catch { }
                        sourcePresentation = null;
                    }
                }
            }

            // Heavy comparison runs on background thread to keep UI responsive
            progress?.Report(Tuple.Create(60, "Running intelligent comparison engine..."));
            return await Task.Run(() =>
            {
                return ComparisonEngine.ComparePresentations(sourceSnapshot, destSnapshot, settings, progress, cancellationToken);
            }, cancellationToken);
        }

        /// <summary>
        /// Serializes comparison results to JSON string.
        /// </summary>
        public string ExportToJson(PresentationComparisonResult result)
        {
            if (result == null) return "{}";

            using (var ms = new MemoryStream())
            {
                var serializer = new DataContractJsonSerializer(typeof(PresentationComparisonResult));
                serializer.WriteObject(ms, result);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        /// <summary>
        /// Generates a standalone, beautiful HTML comparison report matching the dark UI theme.
        /// </summary>
        public string ExportToHtml(PresentationComparisonResult result)
        {
            if (result == null) return "<html><body><h1>No Results</h1></body></html>";

            int matchedDisplayCount = result.MatchedContentCount > 0 ? result.MatchedContentCount : result.MatchedCount;

            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang='en'>");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset='utf-8'>");
            sb.AppendLine("<meta name='viewport' content='width=device-width, initial-scale=1.0'>");
            sb.AppendLine($"<title>Slide Compare Report &bull; {WebUtility.HtmlEncode(result.DestinationPresentationName)}</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("* { box-sizing: border-box; margin: 0; padding: 0; }");
            sb.AppendLine("body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background: #0B1120; color: #F8FAFC; padding: 24px 16px; font-size: 13px; line-height: 1.5; }");
            sb.AppendLine(".container { max-width: 1050px; margin: 0 auto; }");
            sb.AppendLine(".card { background: #131D31; border-radius: 10px; padding: 20px; margin-bottom: 20px; border: 1px solid #1E293B; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.3); }");
            sb.AppendLine(".header-flex { display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 12px; margin-bottom: 16px; border-bottom: 1px solid #1E293B; padding-bottom: 16px; }");
            sb.AppendLine(".title-area h1 { font-size: 20px; font-weight: 700; color: #F8FAFC; margin-bottom: 4px; display: flex; align-items: center; gap: 8px; }");
            sb.AppendLine(".subtitle { color: #94A3B8; font-size: 12px; }");
            sb.AppendLine(".badge-issues { background: #1E293B; border: 1px solid #F59E0B; color: #F59E0B; padding: 4px 12px; border-radius: 20px; font-weight: 700; font-size: 12px; }");
            
            // Summary Dashboard Grid
            sb.AppendLine(".summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(105px, 1fr)); gap: 8px; margin-bottom: 16px; }");
            sb.AppendLine(".summary-btn { background: #0F172A; border: 1px solid #1E293B; padding: 10px 8px; border-radius: 8px; text-align: center; cursor: pointer; transition: all 0.15s ease; outline: none; }");
            sb.AppendLine(".summary-btn:hover { background: #1E293B; border-color: #334155; transform: translateY(-1px); }");
            sb.AppendLine(".summary-btn.active { border-color: #38BDF8; background: #1E293B; box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.25); }");
            sb.AppendLine(".summary-btn .num { font-size: 17px; font-weight: 800; line-height: 1.2; margin-bottom: 2px; }");
            sb.AppendLine(".summary-btn .label { font-size: 10px; text-transform: uppercase; color: #94A3B8; font-weight: 600; letter-spacing: 0.04em; }");

            // Filter Bar
            sb.AppendLine(".filter-bar { display: flex; gap: 10px; margin-bottom: 16px; flex-wrap: wrap; }");
            sb.AppendLine(".search-box { flex: 1; min-width: 200px; background: #1E293B; border: 1px solid #334155; color: #F8FAFC; padding: 8px 12px; border-radius: 6px; font-size: 12px; outline: none; }");
            sb.AppendLine(".search-box:focus { border-color: #38BDF8; }");
            sb.AppendLine(".btn-action { background: #1E293B; border: 1px solid #334155; color: #CBD5E1; padding: 7px 12px; border-radius: 6px; font-size: 11px; font-weight: 600; cursor: pointer; }");
            sb.AppendLine(".btn-action:hover { background: #334155; color: #FFFFFF; }");

            // Slide Cards
            sb.AppendLine(".slide-card { background: #0F172A; border-radius: 8px; padding: 14px 16px; margin-bottom: 12px; border-left: 4px solid #38BDF8; border-top: 1px solid #1E293B; border-right: 1px solid #1E293B; border-bottom: 1px solid #1E293B; }");
            sb.AppendLine(".slide-card.matched { border-left-color: #22C55E; }");
            sb.AppendLine(".slide-card.modified { border-left-color: #F59E0B; }");
            sb.AppendLine(".slide-card.added { border-left-color: #38BDF8; }");
            sb.AppendLine(".slide-card.deleted { border-left-color: #EF4444; }");
            sb.AppendLine(".slide-card.moved { border-left-color: #FBBF24; }");
            sb.AppendLine(".slide-card.reordered { border-left-color: #FBBF24; }");
            sb.AppendLine(".slide-card.split { border-left-color: #06B6D4; }");
            sb.AppendLine(".slide-card.merged { border-left-color: #EC4899; }");

            // Slide Header
            sb.AppendLine(".slide-header { display: flex; justify-content: space-between; align-items: center; cursor: pointer; user-select: none; }");
            sb.AppendLine(".slide-title-group { display: flex; align-items: center; gap: 8px; font-weight: 600; font-size: 13px; color: #F8FAFC; }");
            sb.AppendLine(".status-pill { font-size: 9px; font-weight: 700; padding: 2px 7px; border-radius: 4px; background: #1E293B; text-transform: uppercase; }");
            sb.AppendLine(".slide-summary { color: #94A3B8; font-size: 11px; margin: 4px 0 6px 24px; }");
            sb.AppendLine(".speaker-notes { background: #1E293B; border-radius: 4px; padding: 6px 10px; margin: 4px 0 8px 24px; font-size: 11px; color: #E2E8F0; }");

            // Shapes Inside Slide
            sb.AppendLine(".shapes-container { margin: 8px 0 0 16px; display: flex; flex-direction: column; gap: 8px; }");
            sb.AppendLine(".shape-item { background: #162032; border-radius: 6px; padding: 10px 12px; border: 1px solid #1E293B; }");
            sb.AppendLine(".shape-item-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }");
            sb.AppendLine(".shape-name { font-weight: 600; font-size: 12px; color: #F8FAFC; display: flex; align-items: center; gap: 6px; }");
            sb.AppendLine(".shape-details { font-size: 10px; color: #94A3B8; margin-bottom: 6px; margin-left: 18px; }");

            // Diff Boxes
            sb.AppendLine(".box-source { background: #2A0B0E; border: 1px solid #7F1D1D; border-radius: 5px; padding: 8px; margin-bottom: 6px; font-size: 11px; color: #CBD5E1; }");
            sb.AppendLine(".box-dest { background: #064E3B; border: 1px solid #065F46; border-radius: 5px; padding: 8px; margin-bottom: 6px; font-size: 11px; color: #F1F5F9; }");
            sb.AppendLine(".box-diff { background: #0B1222; border: 1px solid #334155; border-radius: 5px; padding: 8px; margin-top: 4px; font-size: 12px; line-height: 1.6; }");
            sb.AppendLine(".box-del { background: #2A0B0E; border: 1px solid #7F1D1D; border-radius: 5px; padding: 8px; margin-bottom: 4px; font-size: 11px; color: #EF4444; }");
            sb.AppendLine(".box-add { background: #081E36; border: 1px solid #0369A1; border-radius: 5px; padding: 8px; margin-bottom: 4px; font-size: 11px; color: #38BDF8; }");
            sb.AppendLine(".box-mov { background: #241B08; border: 1px solid #854D0E; border-radius: 4px; padding: 6px 8px; font-size: 11px; color: #FDE68A; margin-top: 4px; }");
            sb.AppendLine(".box-matched { background: #0F172A; border: 1px solid #1E293B; border-radius: 5px; padding: 8px; font-size: 11px; color: #94A3B8; }");

            // Diff Highlights (Red strikethrough = deleted, Blue = inserted, Yellow = moved)
            sb.AppendLine(".diff-del { color: #EF4444; text-decoration: line-through; font-weight: 600; background: rgba(239, 68, 68, 0.18); padding: 1px 4px; border-radius: 3px; }");
            sb.AppendLine(".diff-add { color: #38BDF8; font-weight: bold; background: rgba(56, 189, 248, 0.18); padding: 1px 4px; border-radius: 3px; }");
            sb.AppendLine(".diff-mov { color: #FBBF24; font-weight: 600; background: rgba(251, 191, 36, 0.18); padding: 1px 4px; border-radius: 3px; }");
            sb.AppendLine(".diff-same { color: #CBD5E1; }");
            sb.AppendLine(".badge-tag { font-size: 9px; font-weight: 700; padding: 1px 5px; border-radius: 3px; margin-right: 6px; text-transform: uppercase; }");
            sb.AppendLine(".badge-src { background: #450A0A; color: #F87171; }");
            sb.AppendLine(".badge-dst { background: #064E3B; color: #34D399; }");
            sb.AppendLine(".diff-legend { font-size: 9px; font-weight: 600; color: #94A3B8; margin-bottom: 4px; display: flex; gap: 8px; }");
            sb.AppendLine(".fmt-item { font-size: 10px; color: #F59E0B; margin-top: 2px; }");

            sb.AppendLine("</style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("<div class='container'>");

            // Report Header Card
            sb.AppendLine("<div class='card'>");
            sb.AppendLine("<div class='header-flex'>");
            sb.AppendLine("<div class='title-area'>");
            sb.AppendLine($"<h1><span>⚡</span> Slide Compare Report</h1>");
            sb.AppendLine($"<p class='subtitle'>Destination: <b>{WebUtility.HtmlEncode(result.DestinationPresentationName)}</b> &bull; Source: <b>{WebUtility.HtmlEncode(result.SourcePresentationName)}</b> &bull; Generated: {DateTime.Now:MMM dd, yyyy HH:mm:ss}</p>");
            sb.AppendLine("</div>");
            sb.AppendLine($"<div class='badge-issues'>{result.TotalIssuesCount} Differences Found</div>");
            sb.AppendLine("</div>");

            // Interactive Summary Grid
            sb.AppendLine("<div class='summary-grid'>");
            sb.AppendLine($"<div class='summary-btn' id='btn-matched' onclick=\"filterBy('matched')\"><div class='num' style='color:#22C55E;'>✓ {matchedDisplayCount}</div><div class='label'>Matched</div></div>");
            sb.AppendLine($"<div class='summary-btn' id='btn-modified' onclick=\"filterBy('modified')\"><div class='num' style='color:#F59E0B;'>✎ {result.ModifiedCount}</div><div class='label'>Modified</div></div>");
            sb.AppendLine($"<div class='summary-btn' id='btn-added' onclick=\"filterBy('added')\"><div class='num' style='color:#38BDF8;'>+ {result.AddedCount}</div><div class='label'>Added</div></div>");
            sb.AppendLine($"<div class='summary-btn' id='btn-deleted' onclick=\"filterBy('deleted')\"><div class='num' style='color:#EF4444;'>− {result.DeletedCount}</div><div class='label'>Deleted</div></div>");
            sb.AppendLine($"<div class='summary-btn' id='btn-moved' onclick=\"filterBy('moved')\"><div class='num' style='color:#FBBF24;'>↔ {result.MovedCount}</div><div class='label'>Moved</div></div>");
            sb.AppendLine($"<div class='summary-btn' id='btn-reordered' onclick=\"filterBy('reordered')\"><div class='num' style='color:#FBBF24;'>⇄ {result.ReorderedCount}</div><div class='label'>Reordered</div></div>");
            sb.AppendLine($"<div class='summary-btn' id='btn-split' onclick=\"filterBy('split')\"><div class='num' style='color:#06B6D4;'>⇢ {result.SplitCount}</div><div class='label'>Split</div></div>");
            sb.AppendLine($"<div class='summary-btn' id='btn-merged' onclick=\"filterBy('merged')\"><div class='num' style='color:#EC4899;'>⇠ {result.MergedCount}</div><div class='label'>Merged</div></div>");
            sb.AppendLine("</div>");

            // Search and Controls Bar
            sb.AppendLine("<div class='filter-bar'>");
            sb.AppendLine("<input type='text' id='searchInput' class='search-box' placeholder='Search slide numbers, shape names, or text content...' onkeyup='filterSearch()'/>");
            sb.AppendLine("<button class='btn-action' onclick=\"filterBy('all')\">Show All</button>");
            sb.AppendLine("<button class='btn-action' onclick=\"filterBy('issues')\">Issues Only</button>");
            sb.AppendLine("<button class='btn-action' onclick='toggleAll(true)'>Expand All</button>");
            sb.AppendLine("<button class='btn-action' onclick='toggleAll(false)'>Collapse All</button>");
            sb.AppendLine("</div>");
            sb.AppendLine("</div>"); // end card

            // Slide Results List
            sb.AppendLine("<div id='slidesList'>");
            if (result.SlideResults != null)
            {
                foreach (var s in result.SlideResults)
                {
                    string statusClass = s.MatchType.ToString().ToLowerInvariant();
                    string slideColor = GetSlideColor(s.MatchType);
                    bool hasIssues = s.HasIssues;

                    sb.AppendLine($"<div class='slide-card {statusClass}' data-matchtype='{statusClass}' data-hasissues='{hasIssues.ToString().ToLowerInvariant()}'>");
                    sb.AppendLine("<div class='slide-header' onclick='toggleSlide(this)'>");
                    sb.AppendLine($"<div class='slide-title-group'><span style='color:{slideColor}; font-weight:bold;'>{s.StatusIcon}</span> <span>{WebUtility.HtmlEncode(s.TitleDisplay)}</span></div>");
                    sb.AppendLine($"<span class='status-pill' style='color:{slideColor};'>{s.StatusBadgeText}</span>");
                    sb.AppendLine("</div>");

                    if (!string.IsNullOrEmpty(s.SummaryText))
                    {
                        sb.AppendLine($"<div class='slide-summary'>{WebUtility.HtmlEncode(s.SummaryText)}</div>");
                    }

                    if (!string.IsNullOrEmpty(s.SpeakerNotesDiff))
                    {
                        sb.AppendLine($"<div class='speaker-notes'><b>Speaker Notes:</b> {WebUtility.HtmlEncode(s.SpeakerNotesDiff)}</div>");
                    }

                    // Shapes Container
                    if (s.ShapeResults != null && s.ShapeResults.Count > 0)
                    {
                        sb.AppendLine("<div class='shapes-container'>");
                        foreach (var shp in s.ShapeResults)
                        {
                            string shapeStatusClass = shp.MatchType.ToString().ToLowerInvariant();
                            string shapeColor = GetShapeColor(shp.MatchType);

                            sb.AppendLine($"<div class='shape-item' data-matchtype='{shapeStatusClass}' data-isissue='{shp.IsIssue.ToString().ToLowerInvariant()}'>");
                            sb.AppendLine("<div class='shape-item-header'>");
                            sb.AppendLine($"<div class='shape-name'><span style='color:{shapeColor}; font-weight:bold;'>{shp.StatusIcon}</span> <span>{WebUtility.HtmlEncode(shp.DisplayName)}</span></div>");
                            sb.AppendLine($"<span class='status-pill' style='color:{shapeColor};'>{shp.StatusBadgeText}</span>");
                            sb.AppendLine("</div>");

                            if (!string.IsNullOrEmpty(shp.Details))
                            {
                                sb.AppendLine($"<div class='shape-details'>{WebUtility.HtmlEncode(shp.Details)}</div>");
                            }

                            // Modified Shape: Show SOURCE, DEST, and inline diff tokens
                            if (shp.IsModified)
                            {
                                if (shp.HasSourceText)
                                {
                                    sb.AppendLine($"<div class='box-source'><span class='badge-tag badge-src'>SOURCE</span> {WebUtility.HtmlEncode(shp.SourceText)}</div>");
                                }
                                if (shp.HasDestinationText)
                                {
                                    sb.AppendLine($"<div class='box-dest'><span class='badge-tag badge-dst'>DEST</span> {WebUtility.HtmlEncode(shp.DestinationText)}</div>");
                                }
                                if (shp.HasTextChanges)
                                {
                                    sb.AppendLine("<div class='box-diff'>");
                                    sb.AppendLine("<div class='diff-legend'><span style='color:#F59E0B; font-weight:bold;'>DIFFERENCE:</span> <span class='diff-del' style='background:transparent; padding:0;'>Deleted</span> &bull; <span class='diff-add' style='background:transparent; padding:0;'>Inserted</span> &bull; <span class='diff-mov' style='background:transparent; padding:0;'>Moved</span></div>");
                                    sb.AppendLine("<div style='margin-top:4px;'>");

                                    if (shp.DiffTokens != null && shp.DiffTokens.Count > 0)
                                    {
                                        foreach (var token in shp.DiffTokens)
                                        {
                                            string encodedText = WebUtility.HtmlEncode(token.Text);
                                            switch (token.DiffType)
                                            {
                                                case TextDiffType.Deleted:
                                                    sb.Append($"<span class='diff-del'>{encodedText}</span> ");
                                                    break;
                                                case TextDiffType.Added:
                                                    sb.Append($"<span class='diff-add'>{encodedText}</span> ");
                                                    break;
                                                case TextDiffType.Moved:
                                                    sb.Append($"<span class='diff-mov'>{encodedText}</span> ");
                                                    break;
                                                case TextDiffType.Unchanged:
                                                default:
                                                    sb.Append($"<span class='diff-same'>{encodedText}</span> ");
                                                    break;
                                            }
                                        }
                                    }
                                    else if (!string.IsNullOrEmpty(shp.TextDiff))
                                    {
                                        sb.Append($"<span class='diff-same'>{WebUtility.HtmlEncode(shp.TextDiff)}</span>");
                                    }

                                    sb.AppendLine("</div></div>");
                                }
                            }
                            else if (shp.IsDeleted)
                            {
                                sb.AppendLine("<div class='box-del'>");
                                sb.AppendLine("<div style='font-size:10px; font-weight:bold; color:#EF4444; margin-bottom:3px;'>DELETED FROM DESTINATION:</div>");
                                sb.AppendLine($"<div class='diff-del' style='background:transparent; padding:0;'>{WebUtility.HtmlEncode(shp.SourceText)}</div>");
                                sb.AppendLine("</div>");
                            }
                            else if (shp.IsAdded)
                            {
                                sb.AppendLine("<div class='box-add'>");
                                sb.AppendLine("<div style='font-size:10px; font-weight:bold; color:#38BDF8; margin-bottom:3px;'>INSERTED IN DESTINATION:</div>");
                                sb.AppendLine($"<div class='diff-add' style='background:transparent; padding:0;'>{WebUtility.HtmlEncode(shp.DestinationText)}</div>");
                                sb.AppendLine("</div>");
                            }
                            else if (shp.IsMatched)
                            {
                                sb.AppendLine($"<div class='box-matched'>{WebUtility.HtmlEncode(shp.DestinationText)}</div>");
                            }

                            if (shp.HasPositionDiff)
                            {
                                sb.AppendLine($"<div class='box-mov'><b>↔ MOVED:</b> {WebUtility.HtmlEncode(shp.PositionDiff)}</div>");
                            }

                            if (shp.FormattingDifferences != null && shp.FormattingDifferences.Count > 0)
                            {
                                foreach (var fmt in shp.FormattingDifferences)
                                {
                                    sb.AppendLine($"<div class='fmt-item'>&bull; {WebUtility.HtmlEncode(fmt.ToString())}</div>");
                                }
                            }

                            sb.AppendLine("</div>"); // end shape-item
                        }
                        sb.AppendLine("</div>"); // end shapes-container
                    }

                    sb.AppendLine("</div>"); // end slide-card
                }
            }
            sb.AppendLine("</div>"); // end slidesList

            // Embedded JavaScript for Client-Side Interactivity
            sb.AppendLine("<script>");
            sb.AppendLine(@"
function filterBy(type) {
    document.querySelectorAll('.summary-btn').forEach(b => b.classList.remove('active'));
    var btn = document.getElementById('btn-' + type);
    if (btn) btn.classList.add('active');

    var slides = document.querySelectorAll('.slide-card');
    slides.forEach(s => {
        if (type === 'all') {
            s.style.display = '';
        } else if (type === 'issues') {
            s.style.display = s.getAttribute('data-hasissues') === 'true' ? '' : 'none';
        } else {
            var match = s.getAttribute('data-matchtype') === type;
            var hasShape = s.querySelector('.shape-item[data-matchtype=""' + type + '""]') !== null;
            s.style.display = (match || hasShape) ? '' : 'none';
        }
    });
}

function filterSearch() {
    var q = document.getElementById('searchInput').value.toLowerCase().trim();
    var slides = document.querySelectorAll('.slide-card');
    slides.forEach(s => {
        var text = s.innerText.toLowerCase();
        s.style.display = (q === '' || text.indexOf(q) !== -1) ? '' : 'none';
    });
}

function toggleAll(expand) {
    document.querySelectorAll('.shapes-container').forEach(c => {
        c.style.display = expand ? '' : 'none';
    });
}

function toggleSlide(header) {
    var card = header.closest('.slide-card');
    var container = card.querySelector('.shapes-container');
    if (container) {
        container.style.display = container.style.display === 'none' ? '' : 'none';
    }
}
");
            sb.AppendLine("</script>");

            sb.AppendLine("</div>"); // end container
            sb.AppendLine("</body></html>");
            return sb.ToString();
        }

        private static string GetSlideColor(SlideMatchType type)
        {
            switch (type)
            {
                case SlideMatchType.Matched: return "#22C55E";
                case SlideMatchType.Modified: return "#F59E0B";
                case SlideMatchType.Added: return "#38BDF8";
                case SlideMatchType.Deleted: return "#EF4444";
                case SlideMatchType.Moved: return "#FBBF24";
                case SlideMatchType.Reordered: return "#FBBF24";
                case SlideMatchType.Split: return "#06B6D4";
                case SlideMatchType.Merged: return "#EC4899";
                default: return "#94A3B8";
            }
        }

        private static string GetShapeColor(ShapeMatchType type)
        {
            switch (type)
            {
                case ShapeMatchType.Matched: return "#22C55E";
                case ShapeMatchType.Modified: return "#F59E0B";
                case ShapeMatchType.Added: return "#38BDF8";
                case ShapeMatchType.Deleted: return "#EF4444";
                case ShapeMatchType.Moved: return "#FBBF24";
                default: return "#94A3B8";
            }
        }
    }
}
