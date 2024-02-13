using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;

namespace VRChat_Exif;

[SuppressMessage("Interoperability", "CA1416:プラットフォームの互換性を検証")]
internal abstract class Program
{
    private static string? _output;
    private static string? _path;
    public static void Main(string[] args)
    {
        var configPath = Path.Combine(Environment.CurrentDirectory, "config.txt");
        (_path, _output) = ReadConfig(configPath);
        if (_path == null || _output == null)
        {
            Console.WriteLine("config.txtに監視ディレクトリまたは出力先ディレクトリが指定されていません。");
            Console.WriteLine("終了するにはEnterキーを押してください。");
            Console.ReadLine();
        }
        else if (!Directory.Exists(_path) || !Directory.Exists(_output))
        {
            Console.WriteLine("監視ディレクトリまたは出力先ディレクトリに存在しないパスが指定されています。");
            Console.WriteLine("終了するにはEnterキーを押してください。");
            Console.ReadLine();
        }
        else
        {
            if (!Directory.Exists(_output))
            {
                Directory.CreateDirectory(_output);
            }
            var watcher = new FileSystemWatcher();
            watcher.Path = _path;
            watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            watcher.Filter = "*.png";
            watcher.IncludeSubdirectories = true;
            watcher.Created += file_Created;
            watcher.EnableRaisingEvents = true;
            Console.WriteLine("VRChat-スクリーンショットjpg変換&Exif情報追加プログラム");
            Console.WriteLine("監視ディレクトリ: " + _path);
            Console.WriteLine("出力先ディレクトリ: " + _output);
            Console.WriteLine("監視を開始します。");
            Console.WriteLine("終了するにはEnterキーを押してください。");
            Console.ReadLine();
        }
    }
    
    private static (string?, string?) ReadConfig(string configPath)
    {
        if (File.Exists(configPath))
        {
            var strArray = File.ReadAllLines(configPath);
            return strArray.Length >= 2 ? (strArray[0].Trim(), strArray[1].Trim()) : (null, null);
        }

        File.WriteAllText(configPath, $@"C:\Users\{Environment.UserName}\Pictures\VRChat{Environment.NewLine}C:\Users\{Environment.UserName}\Pictures\VRChat\Exif");
        if (!Directory.Exists(@$"C:\Users\{Environment.UserName}\Pictures\VRChat\Exif"))
        {
            Directory.CreateDirectory(@$"C:\Users\{Environment.UserName}\Pictures\VRChat\Exif");
        }
        return ($@"C:\Users\{Environment.UserName}\Pictures\VRChat", $@"C:\Users\{Environment.UserName}\Pictures\VRChat\Exif");
    }
    
    private static void file_Created(object obj, FileSystemEventArgs e)
    {
        Thread.Sleep(1000);
        using var image = Image.FromFile(e.FullPath);
        var pi = image.PropertyItems[0];
        pi.Id = 0x9003;
        pi.Type = 2;
        var filename = e.Name!.Replace(".png", "").Split('\\')[1];
        var datetime = $"{filename.Substring(7, 19)}".Replace("-", ":").Replace("_", "");
        pi.Value = Encoding.ASCII.GetBytes(datetime);
        pi.Len = pi.Value.Length;
        image.SetPropertyItem(pi);
        image.Save(Path.Combine(_output!, $"{filename}.jpg"), ImageFormat.Jpeg);
        if (File.Exists(Path.Combine(_output!, $"{filename}.jpg")))
        {
            Console.WriteLine(Path.GetFileName(Path.Combine(_output!, $"{filename}.jpg")) + " にExif情報を追加しました。");
        }

    }
}

