using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class Launcher
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length != 1 || !File.Exists(args[0]))
                throw new ArgumentException("Wybierz jeden plik DOC lub DOCX.");
            string file = Path.GetFullPath(args[0]);
            string ext = Path.GetExtension(file);
            if (!ext.Equals(".doc", StringComparison.OrdinalIgnoreCase) &&
                !ext.Equals(".docx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Wybierz plik DOC lub DOCX.");
            string script = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ConvertToPDF.ps1");
            var info = new ProcessStartInfo {
                FileName = Path.Combine(Environment.SystemDirectory, @"WindowsPowerShell\v1.0\powershell.exe"),
                Arguments = "-NoProfile -STA -File \"" + script + "\" \"" + file + "\"",
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using (var process = Process.Start(info)) {
                process.WaitForExit();
                if (process.ExitCode != 0 && process.ExitCode != 2)
                    throw new InvalidOperationException("Konwersja nie powiodła się. Sprawdź konfigurację PowerShell i Word.");
                return process.ExitCode;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Błąd konwersji PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }
}
