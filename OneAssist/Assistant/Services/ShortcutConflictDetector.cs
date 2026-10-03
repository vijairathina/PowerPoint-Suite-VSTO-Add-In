using System;
using System.Collections.Generic;

namespace OneAssist.Assistant.Services
{
    public static class ShortcutConflictDetector
    {
        private static readonly HashSet<string> ReservedShortcuts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Ctrl+C",
            "Ctrl+V",
            "Ctrl+X",
            "Ctrl+Z",
            "Ctrl+Y",
            "Ctrl+S",
            "Ctrl+A",
            "Ctrl+P",
            "Ctrl+F",
            "Ctrl+N",
            "Ctrl+O",
            "F5",
            "Shift+F5",
            "Escape",
            "Delete"
        };

        public static bool IsConflicting(string shortcut, out string reason)
        {
            reason = string.Empty;
            if (string.IsNullOrWhiteSpace(shortcut)) return false;

            string normalized = shortcut.Trim().Replace(" ", "");
            if (ReservedShortcuts.Contains(normalized))
            {
                reason = $"'{shortcut}' conflicts with a standard Microsoft PowerPoint hotkey.";
                return true;
            }

            return false;
        }
    }
}
