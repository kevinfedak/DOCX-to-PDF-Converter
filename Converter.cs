using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyVersion("2.0.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("2.0.0.0")]
internal static class Converter
{
    static void Release(object obj) {
        if (obj != null && Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj);
    }
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 1 && args[0] == "--self-test") return 0;
        dynamic word = null, docs = null, doc = null, options = null;
        string temp = null;
        try {
            if (args.Length != 1 || !File.Exists(args[0]))
                throw new ArgumentException("Wybierz jeden plik DOC lub DOCX.");
            string file = Path.GetFullPath(args[0]);
            string ext = Path.GetExtension(file);
            if (!ext.Equals(".doc", StringComparison.OrdinalIgnoreCase) && !ext.Equals(".docx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Wybierz plik DOC lub DOCX.");
            string target = Path.ChangeExtension(file, ".pdf");
            if (File.Exists(target)) throw new IOException("Plik PDF już istnieje.");
            temp = Path.Combine(Path.GetDirectoryName(file), Guid.NewGuid().ToString() + ".pdf");
            Type type = Type.GetTypeFromProgID("Word.Application");
            if (type == null) throw new InvalidOperationException("Wymagany jest Microsoft Word.");
            word = Activator.CreateInstance(type);
            word.Visible = false;
            word.DisplayAlerts = 0;
            word.AutomationSecurity = 3;
            options = word.Options;
            options.UpdateLinksAtOpen = false;
            docs = word.Documents;
            doc = docs.Open(file, false, true, false);
            doc.ExportAsFixedFormat(temp, 17);
            File.Move(temp, target);
            return 0;
        } catch (Exception ex) {
            MessageBox.Show(ex.Message, "Błąd konwersji PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 2;
        } finally {
            if (doc != null) { try { doc.Close(0); } catch {} }
            if (word != null) { try { word.Quit(0); } catch {} }
            Release((object)doc); Release((object)docs); Release((object)options); Release((object)word);
            if (temp != null) { try { File.Delete(temp); } catch {} }
        }
    }
}
