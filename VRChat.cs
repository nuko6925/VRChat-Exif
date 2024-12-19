using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace VRChat_Exif;

[SuppressMessage("ReSharper", "InconsistentNaming")]
public partial class VRChat
{
    private static string ExtractWorldNameFromLog(string logFilePath)
    {
        try
        {
            using var fs = new FileStream(logFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(fs);
            var lines = reader.ReadToEnd().Split(Environment.NewLine);
            var ls = lines.Reverse().ToArray();
            foreach (var line in ls)
            {
                if (!line.Contains("Entering Room:")) continue;
                var match = Regex.Match(line, "Entering Room: (.+)");
                if (!match.Success) continue;
                var worldName = match.Groups[1].Value;
                return worldName;
            }
        }
        catch (Exception)
        {
            //
        }
        return "Error World";
    }
    
    public static string? GetWorldName()
    {
        var baseLogFolder =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData).Replace("Local", "LocalLow"),
                "VRChat", "VRChat");
        if (!Directory.Exists(baseLogFolder))
        {
            return null;
        }
        var logFiles = Directory.GetFiles(baseLogFolder, "output_log_*.txt", SearchOption.AllDirectories);
        if (logFiles.Length == 0)
        {
            return null;
        }
        var latestLogFile = logFiles.OrderByDescending(f => new FileInfo(f).LastWriteTime).FirstOrDefault();
        if (latestLogFile == null)
        {
            return null;
        }
        var worldName = ExtractWorldNameFromLog(latestLogFile);
        return !string.IsNullOrEmpty(worldName) ? worldName : null;
    }
}