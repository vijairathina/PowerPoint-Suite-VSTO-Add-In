using System;
using System.Collections.Generic;
using System.Linq;

namespace OneAssist.Assistant.CommandPalette
{
    public class PaletteCommand
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Category { get; set; }
        public string Keywords { get; set; }
        public Action Action { get; set; }
    }

    public static class CommandPaletteService
    {
        private static readonly List<PaletteCommand> Commands = new List<PaletteCommand>();

        public static void Register(string id, string title, string category, string keywords, Action action)
        {
            Commands.Add(new PaletteCommand
            {
                Id = id,
                Title = title,
                Category = category,
                Keywords = keywords,
                Action = action
            });
        }

        public static List<PaletteCommand> Search(string query)
        {
            if (string.IsNullOrWhiteSpace(query)) return Commands.Take(12).ToList();

            string q = query.Trim().ToLower();
            return Commands.Where(c => c.Title.ToLower().Contains(q) ||
                                       c.Category.ToLower().Contains(q) ||
                                       (c.Keywords != null && c.Keywords.ToLower().Contains(q)))
                           .Take(12)
                           .ToList();
        }
    }
}
