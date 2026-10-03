using System;
using System.IO;

namespace OneAssist.ImageEditor.Utils
{
    /// <summary>
    /// Manages temporary files securely, ensuring unique names and guaranteed cleanup.
    /// </summary>
    public sealed class TempFileManager : IDisposable
    {
        private static readonly string TempDirectory = Path.Combine(Path.GetTempPath(), "OneAssist_ImageEditor");
        private readonly System.Collections.Concurrent.ConcurrentBag<string> _trackedFiles = 
            new System.Collections.Concurrent.ConcurrentBag<string>();
        private bool _disposed;

        static TempFileManager()
        {
            try
            {
                if (!Directory.Exists(TempDirectory))
                {
                    Directory.CreateDirectory(TempDirectory);
                }
                else
                {
                    // Clean up stale files older than 24 hours
                    var dir = new DirectoryInfo(TempDirectory);
                    foreach (var file in dir.GetFiles())
                    {
                        if (DateTime.UtcNow - file.CreationTimeUtc > TimeSpan.FromHours(24))
                        {
                            try { file.Delete(); } catch { }
                        }
                    }
                }
            }
            catch
            {
                // Silently fallback if temp directory cannot be pre-cleaned
            }
        }

        public string GetTempPngPath()
        {
            EnsureDirectory();
            string path = Path.Combine(TempDirectory, $"img_{Guid.NewGuid():N}.png");
            _trackedFiles.Add(path);
            return path;
        }

        public string GetTempJpgPath()
        {
            EnsureDirectory();
            string path = Path.Combine(TempDirectory, $"img_{Guid.NewGuid():N}.jpg");
            _trackedFiles.Add(path);
            return path;
        }

        private static void EnsureDirectory()
        {
            if (!Directory.Exists(TempDirectory))
            {
                Directory.CreateDirectory(TempDirectory);
            }
        }

        public static void SafeDelete(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch
            {
                // Ignore transient file lock errors
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            foreach (var file in _trackedFiles)
            {
                SafeDelete(file);
            }
        }
    }
}
