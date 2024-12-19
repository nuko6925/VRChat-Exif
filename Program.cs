using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

namespace VRChat_Exif;

[SuppressMessage("ReSharper", "InconsistentNaming")]
[SuppressMessage("ReSharper", "ArrangeTypeMemberModifiers")]
[SuppressMessage("Interoperability", "SYSLIB1054:コンパイル時に P/Invoke マーシャリング コードを生成するには、\'DllImportAttribute\' の代わりに \'LibraryImportAttribute\' を使用します")]
[SuppressMessage("Interoperability", "SYSLIB1096:\'GeneratedComInterface\' に変換する")]
[SuppressMessage("ReSharper", "ConvertToPrimaryConstructor")]
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
            watcher.Created += FileCreated;
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
    
    private static void FileCreated(object obj, FileSystemEventArgs e)
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
        var world = VRChat.GetWorldName();
        /*pi = image.PropertyItems[0];
        pi.Id = (int)ExifProperty.Title; //0x9286
        pi.Type = 7;
        var comment = new string('\0', 8) + _worldName;
        pi.Value = Encoding.ASCII.GetBytes(comment);
        image.SetPropertyItem(pi);*/
        image.Save(Path.Combine(_output!, $"{filename}.jpg"), ImageFormat.Jpeg);
        if (File.Exists(Path.Combine(_output!, $"{filename}.jpg")))
        {
            Console.WriteLine(Path.GetFileName(Path.Combine(_output!, $"{filename}.jpg")) + " にExif情報を追加しました。");
            SetSubjectAndComment(Path.Combine(_output!, $"{filename}.jpg"), "VRChat", world!);
        }

    }
    
    [Flags]
    enum GETPROPERTYSTOREFLAGS {
        READWRITE = 0x00000002
    }

    [StructLayout(LayoutKind.Sequential)]
    struct PROPERTYKEY {
        public Guid fmtid;
        public int pid;

        public static PROPERTYKEY FromName(string name) {
            PSGetPropertyKeyFromName(name, out var key);
            return key;
        }
    }

    enum VARTYPE : ushort {
        VT_LPWSTR = 31,
    }

    [StructLayout(LayoutKind.Explicit)]
    struct PROPVARIANT {
        [FieldOffset(0)]
        public VARTYPE vt;
        [FieldOffset(8)]
        public IntPtr bstrVal;

        public PROPVARIANT(string val) {
            vt = VARTYPE.VT_LPWSTR;
            bstrVal = Marshal.StringToCoTaskMemUni(val);
        }
    }

    [ComImport, Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyStore {
        void GetCount(out int cProps);
        void GetAt(int iProp, out PROPERTYKEY pkey);
        void GetValue([In] ref PROPERTYKEY key, out PROPVARIANT pv);
        void SetValue([In] ref PROPERTYKEY key, [In] ref PROPVARIANT propvar);
        void Commit();
    }

    [DllImport("Shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    static extern void SHGetPropertyStoreFromParsingName(string pszPath, IntPtr pbc, GETPROPERTYSTOREFLAGS flags, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IPropertyStore ppv);
    [DllImport("Propsys.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    static extern void PSGetPropertyKeyFromName(string pszString, out PROPERTYKEY pkey);
    [DllImport("ole32.dll", PreserveSig = false)]
    static extern void PropVariantClear(ref PROPVARIANT pvar);
    
    private static void SetSubjectAndComment(string filepath, string subject, string comment) {
        SHGetPropertyStoreFromParsingName(filepath, IntPtr.Zero, GETPROPERTYSTOREFLAGS.READWRITE, typeof(IPropertyStore).GUID, out var prop);

        try
        {
            var keySubject = PROPERTYKEY.FromName("System.Subject");
            var valSubject = new PROPVARIANT(subject);
            prop.SetValue(ref keySubject, ref valSubject);
            PropVariantClear(ref valSubject);

            var keyComment = PROPERTYKEY.FromName("System.Comment");
            var valComment = new PROPVARIANT(comment);
            prop.SetValue(ref keyComment, ref valComment);
            PropVariantClear(ref valComment);
            prop.Commit();
        }
        finally
        {
            if (true)
            {
                Marshal.ReleaseComObject(prop);
            }
        }

    }
}

