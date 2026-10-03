using System;
using System.Collections.Generic;
using OneAssist.Assistant.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Production
{
    public class ShapeSnapshotData
    {
        public int ShapeId { get; set; }
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public string Text { get; set; }
        public string FontName { get; set; }
        public float? FontSize { get; set; }
    }

    public class SlideSnapshot
    {
        public string SnapshotId { get; set; } = Guid.NewGuid().ToString("N");
        public string Description { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public int SlideIndex { get; set; }
        public List<ShapeSnapshotData> Shapes { get; set; } = new List<ShapeSnapshotData>();
    }

    public static class SlideSnapshotService
    {
        private static readonly List<SlideSnapshot> Snapshots = new List<SlideSnapshot>();

        public static SlideSnapshot CreateSnapshot(PowerPoint.Slide slide, string description)
        {
            if (slide == null) return null;

            var snap = new SlideSnapshot
            {
                SlideIndex = slide.SlideIndex,
                Description = description
            };

            try
            {
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    var data = new ShapeSnapshotData
                    {
                        ShapeId = s.Id,
                        Left = s.Left,
                        Top = s.Top,
                        Width = s.Width,
                        Height = s.Height
                    };

                    if (SafeInteropHelper.HasText(s))
                    {
                        data.Text = s.TextFrame.TextRange.Text;
                        data.FontName = SafeInteropHelper.GetFontName(s);
                        data.FontSize = SafeInteropHelper.GetFontSize(s);
                    }

                    snap.Shapes.Add(data);
                }

                Snapshots.Add(snap);
                if (Snapshots.Count > 20) Snapshots.RemoveAt(0);
            }
            catch { }

            return snap;
        }

        public static bool RestoreSnapshot(PowerPoint.Slide slide, SlideSnapshot snapshot)
        {
            if (slide == null || snapshot == null) return false;

            try
            {
                foreach (var data in snapshot.Shapes)
                {
                    foreach (PowerPoint.Shape s in slide.Shapes)
                    {
                        if (s.Id == data.ShapeId)
                        {
                            s.Left = data.Left;
                            s.Top = data.Top;
                            s.Width = data.Width;
                            s.Height = data.Height;

                            if (data.Text != null && SafeInteropHelper.HasText(s))
                            {
                                s.TextFrame.TextRange.Text = data.Text;
                                if (!string.IsNullOrEmpty(data.FontName)) s.TextFrame.TextRange.Font.Name = data.FontName;
                                if (data.FontSize.HasValue) s.TextFrame.TextRange.Font.Size = data.FontSize.Value;
                            }
                            break;
                        }
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
