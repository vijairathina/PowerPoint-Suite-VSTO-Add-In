using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    /// <summary>
    /// Coordinates undo operations and PowerPoint command bar actions.
    /// </summary>
    public static class PowerPointUndoHelper
    {
        public static void Undo(PowerPoint.Application app)
        {
            if (app == null) return;
            try
            {
                app.CommandBars.ExecuteMso("Undo");
            }
            catch { }
        }

        public static void Redo(PowerPoint.Application app)
        {
            if (app == null) return;
            try
            {
                app.CommandBars.ExecuteMso("Redo");
            }
            catch { }
        }
    }
}
