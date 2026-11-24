using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyVersion("2.1.1.0")]
[assembly: System.Reflection.AssemblyFileVersion("2.1.1.0")]
internal static class Converter
{
    private static readonly string QueuePath = Path.Combine(Path.GetTempPath(), "KevinWordPdf-queue.txt");
    private static Mutex queueMutex;
    private static Mutex leaderMutex;
    private static bool leaderHeld;

    private static void Lock(Mutex mutex)
    {
        try { mutex.WaitOne(); } catch (AbandonedMutexException) { }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 1 && args[0] == "--self-test") return SelfTest();
        if (args.Length != 1) {
            MessageBox.Show("Wybierz pliki DOC lub DOCX w Eksploratorze Windows.", "Konwersja do PDF");
            return 2;
        }
        try {
            queueMutex = new Mutex(false, @"Local\KevinWordPdfQueue");
            leaderMutex = new Mutex(false, @"Local\KevinWordPdfLeader");
            Lock(queueMutex);
            try {
                try { leaderHeld = leaderMutex.WaitOne(0); }
                catch (AbandonedMutexException) { leaderHeld = true; }
                if (leaderHeld && File.Exists(QueuePath)) File.Delete(QueuePath);
                File.AppendAllText(QueuePath, Convert.ToBase64String(Encoding.UTF8.GetBytes(args[0])) + "\n");
            } finally { queueMutex.ReleaseMutex(); }
            if (!leaderHeld) return 0;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ProgressWindow());
            return 0;
        } catch (Exception ex) {
            MessageBox.Show(ex.Message, "Błąd konwersji PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 2;
        } finally {
            if (leaderHeld) { leaderMutex.ReleaseMutex(); leaderHeld = false; }
            if (leaderMutex != null) leaderMutex.Dispose();
            if (queueMutex != null) queueMutex.Dispose();
        }
    }

    private static List<string> ReadQueue()
    {
        var files = new List<string>();
        Lock(queueMutex);
        try {
            if (File.Exists(QueuePath)) {
                foreach (string line in File.ReadAllLines(QueuePath))
                    if (line.Length > 0) files.Add(Encoding.UTF8.GetString(Convert.FromBase64String(line)));
                File.Delete(QueuePath);
            }
        } finally { queueMutex.ReleaseMutex(); }
        return files;
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

    private static void Release(object obj)
    {
        try { if (obj != null && Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); }
        catch { }
    }

    private sealed class WordSession : IDisposable
    {
        private dynamic word, docs, options;
        private readonly Func<object> factory;
        private bool disposed;

        internal WordSession() : this(CreateWord) { }
        internal WordSession(Func<object> factory) { this.factory = factory; }

        private static object CreateWord()
        {
            Type type = Type.GetTypeFromProgID("Word.Application");
            if (type == null) throw new InvalidOperationException("Wymagany jest zainstalowany Microsoft Word dla systemu Windows.");
            return Activator.CreateInstance(type);
        }

        private void Start()
        {
            if (word != null) return;
            word = factory();
            word.Visible = false;
            word.DisplayAlerts = 0;
            word.AutomationSecurity = 3;
            options = word.Options;
            options.UpdateLinksAtOpen = false;
            docs = word.Documents;
        }

        internal void ConvertFile(string input)
        {
            if (disposed) throw new ObjectDisposedException("WordSession");
            // Validation failures do not affect the Word session.
            string target = GetTarget(new[] { input });
            string file = Path.GetFullPath(input);
            string temp = Path.Combine(Path.GetDirectoryName(file), Guid.NewGuid().ToString() + ".pdf");
            try {
                try {
                    dynamic doc = null;
                    try {
                        Start();
                        doc = docs.Open(FileName: file, ConfirmConversions: false, ReadOnly: true,
                            AddToRecentFiles: false, PasswordDocument: Guid.NewGuid().ToString(),
                            Revert: false, Visible: false, OpenAndRepair: false, NoEncodingDialog: true);
                        doc.ExportAsFixedFormat(temp, 17);
                    } finally {
                        if (doc != null) { try { doc.Close(0); } catch {} }
                        Release((object)doc);
                    }
                } catch {
                    // A failed Word operation may leave COM in an unusable state.
                    Reset();
                    throw;
                }
                File.Move(temp, target);
            } finally {
                try { File.Delete(temp); } catch {}
            }
        }

        private void Reset()
        {
            if (word != null) { try { word.Quit(0); } catch {} }
            Release((object)docs); Release((object)options); Release((object)word);
            docs = null; options = null; word = null;
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            Reset();
        }
    }

    // Lightweight Word substitutes used only by --self-test; they never start Word.
    public sealed class TestWord
    {
        public bool Visible { get; set; }
        public int DisplayAlerts { get; set; }
        public int AutomationSecurity { get; set; }
        public TestOptions Options { get; private set; }
        public TestDocuments Documents { get; private set; }
        public bool WasQuit { get; private set; }
        public TestWord(bool fail) { Options = new TestOptions(); Documents = new TestDocuments(fail); }
        public void Quit(int save) { WasQuit = true; }
    }
    public sealed class TestOptions { public bool UpdateLinksAtOpen { get; set; } }
    public sealed class TestDocuments
    {
        private readonly bool fail;
        public TestDocuments(bool fail) { this.fail = fail; }
        public TestDocument Open(string FileName, bool ConfirmConversions, bool ReadOnly, bool AddToRecentFiles,
            string PasswordDocument, bool Revert, bool Visible, bool OpenAndRepair, bool NoEncodingDialog)
        { return new TestDocument(fail); }
    }
    public sealed class TestDocument
    {
        private readonly bool fail;
        public TestDocument(bool fail) { this.fail = fail; }
        public void ExportAsFixedFormat(string path, int format)
        {
            if (fail) throw new InvalidOperationException("Simulated Word error");
            File.WriteAllText(path, "test pdf");
        }
        public void Close(int save) { }
    }

    private sealed class ProgressWindow : Form
    {
        private readonly Label status = new Label();
        private readonly Label current = new Label();
        private readonly ProgressBar bar = new ProgressBar();
        private readonly TextBox details = new TextBox();
        private readonly Button close = new Button();
        private readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        private readonly Queue<string> pending = new Queue<string>();
        private readonly BlockingCollection<string> work = new BlockingCollection<string>();
        private readonly Thread worker;
        private int total, completed, errors;
        private bool working;
        private DateTime lastArrival = DateTime.UtcNow;

        internal ProgressWindow()
        {
            worker = new Thread(WorkLoop);
            worker.IsBackground = true; worker.SetApartmentState(ApartmentState.STA); worker.Start();
            Text = "Konwersja do PDF";
            Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            ClientSize = new Size(560, 300);
            MinimumSize = new Size(500, 320);
            StartPosition = FormStartPosition.CenterScreen;
            status.SetBounds(20, 20, 520, 24); status.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            current.SetBounds(20, 50, 520, 24); current.AutoEllipsis = true; current.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            bar.SetBounds(20, 80, 520, 23); bar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            details.SetBounds(20, 115, 520, 125); details.Multiline = true; details.ReadOnly = true; details.ScrollBars = ScrollBars.Vertical; details.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            close.Text = "Zamknij"; close.SetBounds(440, 255, 100, 28); close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right; close.Enabled = false; close.Click += delegate { Close(); };
            Controls.AddRange(new Control[] { status, current, bar, details, close });
            status.Text = "Przygotowywanie dokumentów...";
            timer.Interval = 250; timer.Tick += Tick; timer.Start();
            FormClosing += OnClosing;
        }

        private void Tick(object sender, EventArgs e)
        {
            try {
                List<string> added = ReadQueue();
                if (added.Count > 0) {
                    foreach (string file in added) pending.Enqueue(file);
                    total += added.Count; lastArrival = DateTime.UtcNow; close.Enabled = false;
                }
                if (!working && pending.Count > 0 && (DateTime.UtcNow - lastArrival).TotalSeconds >= 1.5) StartNext();
                if (!working && pending.Count == 0 && total > 0) {
                    status.Text = errors == 0 ? "Ukończono: " + completed + " z " + total + " dokumentów." : "Ukończono: " + (completed - errors) + " z " + total + "; błędy: " + errors + ".";
                    current.Text = ""; bar.Maximum = total; bar.Value = completed; close.Enabled = true;
                }
            } catch (Exception ex) { timer.Stop(); MessageBox.Show(ex.Message, "Błąd kolejki", MessageBoxButtons.OK, MessageBoxIcon.Error); close.Enabled = true; }
        }

        private void StartNext()
        {
            string file = pending.Dequeue(); working = true;
            status.Text = "Konwertowanie " + (completed + 1) + " z " + total + " dokumentów...";
            current.Text = Path.GetFileName(file); bar.Maximum = total; bar.Value = completed;
            work.Add(file);
        }

        private void WorkLoop()
        {
            using (WordSession session = new WordSession()) {
                foreach (string file in work.GetConsumingEnumerable()) {
                    string error = null;
                    try { session.ConvertFile(file); }
                    catch (Exception ex) { error = ex.Message; }
                    string result = error;
                    BeginInvoke((Action)delegate {
                        completed++; working = false;
                        if (result != null) { errors++; details.AppendText(Path.GetFileName(file) + ": " + result + Environment.NewLine); }
                        else details.AppendText(Path.GetFileName(file) + ": gotowe" + Environment.NewLine);
                    });
                }
            }
        }

        private void OnClosing(object sender, FormClosingEventArgs e)
        {
            if (working || pending.Count > 0 || !close.Enabled) { e.Cancel = true; return; }
            Lock(queueMutex);
            try {
                // Prevent a new launcher from writing to a queue whose window is closing.
                if (File.Exists(QueuePath)) { e.Cancel = true; return; }
                work.CompleteAdding();
                worker.Join();
                leaderMutex.ReleaseMutex(); leaderHeld = false;
            } finally { queueMutex.ReleaseMutex(); }
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
            string other = Path.Combine(dir, "test.txt"); File.WriteAllText(other, "test");
            try { GetTarget(new[] { other }); return 1; } catch (ArgumentException) { }
            File.WriteAllText(Path.ChangeExtension(doc, ".pdf"), "existing");
            try { GetTarget(new[] { doc }); return 1; } catch (IOException) { }
            // A file validation error must preserve a healthy Word session; a Word error
            // must reset it, and the next document must use a new session.
            string one = Path.Combine(dir, "one.docx");
            string two = Path.Combine(dir, "two.docx");
            string three = Path.Combine(dir, "three.docx");
            File.WriteAllText(one, "test"); File.WriteAllText(two, "test"); File.WriteAllText(three, "test");
            var instances = new List<TestWord>();
            var session = new WordSession(delegate { var w = new TestWord(instances.Count == 0); instances.Add(w); return w; });
            try { session.ConvertFile(one); return 1; } catch (InvalidOperationException) { }
            if (instances.Count != 1 || !instances[0].WasQuit) return 1;
            session.ConvertFile(two);
            if (instances.Count != 2 || instances[1].WasQuit || !File.Exists(Path.ChangeExtension(two, ".pdf"))) return 1;
            try { session.ConvertFile(two); return 1; } catch (IOException) { }
            if (instances.Count != 2 || instances[1].WasQuit) return 1;
            session.ConvertFile(three);
            if (instances.Count != 2 || !File.Exists(Path.ChangeExtension(three, ".pdf"))) return 1;
            session.Dispose(); session.Dispose();
            if (!instances[1].WasQuit) return 1;
            try { session.ConvertFile(one); return 1; } catch (ObjectDisposedException) { }
            return 0;
        } catch { return 1; }
        finally { try { Directory.Delete(dir, true); } catch {} }
    }
}
