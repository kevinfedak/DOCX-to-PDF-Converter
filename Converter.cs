using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;

[assembly: System.Reflection.AssemblyVersion("2.1.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("2.1.0.0")]
internal static class Converter
{
    private static void Release(object obj)
    {
        try { if (obj != null && Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); }
        catch { /* Cleanup must not hide the original error. */ }
    }

    private static string GetTarget(string[] args)
    {
        if (args.Length != 1 || !File.Exists(args[0]))
            throw new ArgumentException("Wybierz jeden istniejący plik DOC lub DOCX.");
        string file = Path.GetFullPath(args[0]);
        string ext = Path.GetExtension(file);
        if (!ext.Equals(".doc", StringComparison.OrdinalIgnoreCase) && !ext.Equals(".docx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Wybierz plik DOC lub DOCX.");
        string target = Path.ChangeExtension(file, ".pdf");
        if (File.Exists(target)) throw new IOException("Plik PDF już istnieje: " + target);
        return target;
    }

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 1 && args[0] == "--self-test") return SelfTest();
        dynamic word = null, docs = null, doc = null, options = null;
        string temp = null;
        Mutex conversionLock = null;
        bool lockHeld = false;
        try {
            // Explorer launches a separate process per selected file. Serialize Word automation.
            conversionLock = new Mutex(false, @"Local\KevinWordPdfConversion");
            try { conversionLock.WaitOne(); lockHeld = true; }
            catch (AbandonedMutexException) { lockHeld = true; }
            string target = GetTarget(args);
            string file = Path.GetFullPath(args[0]);
            temp = Path.Combine(Path.GetDirectoryName(file), Guid.NewGuid().ToString() + ".pdf");
            Type type = Type.GetTypeFromProgID("Word.Application");
            if (type == null) throw new InvalidOperationException("Wymagany jest zainstalowany Microsoft Word dla systemu Windows.");
            word = Activator.CreateInstance(type);
            word.Visible = false;
            word.DisplayAlerts = 0;
            word.AutomationSecurity = 3;
            options = word.Options;
            options.UpdateLinksAtOpen = false;
            docs = word.Documents;
            // No repair, conversion or password dialogs should block the invisible Word process.
            doc = docs.Open(FileName: file, ConfirmConversions: false, ReadOnly: true,
                AddToRecentFiles: false, PasswordDocument: Guid.NewGuid().ToString(),
                Revert: false, Visible: false, OpenAndRepair: false, NoEncodingDialog: true);
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
            if (lockHeld) conversionLock.ReleaseMutex();
            if (conversionLock != null) conversionLock.Dispose();
        }
    }

    private static int SelfTest()
    {
        string dir = Path.Combine(Path.GetTempPath(), "WordPdfConverter-test-" + Guid.NewGuid());
        try {
            Directory.CreateDirectory(dir);
            string doc = Path.Combine(dir, "test.DOCX");
            File.WriteAllText(doc, "test");
            if (GetTarget(new[] { doc }) != Path.ChangeExtension(doc, ".pdf")) return 1;
            try { GetTarget(new string[0]); return 1; } catch (ArgumentException) { }
            string other = Path.Combine(dir, "test.txt");
            File.WriteAllText(other, "test");
            try { GetTarget(new[] { other }); return 1; } catch (ArgumentException) { }
            File.WriteAllText(Path.ChangeExtension(doc, ".pdf"), "existing");
            try { GetTarget(new[] { doc }); return 1; } catch (IOException) { }
            return 0;
        } catch { return 1; }
        finally { try { Directory.Delete(dir, true); } catch {} }
    }
}
