using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: System.Reflection.AssemblyTitle("Meeting Capture")]
[assembly: System.Reflection.AssemblyDescription("One fixed region. One-click screenshots.")]
[assembly: System.Reflection.AssemblyVersion("1.1.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.1.0.0")]

namespace MeetingCapture
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Native.ConfigureDpi();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ThreadException += delegate(object sender, System.Threading.ThreadExceptionEventArgs e)
            {
                MessageBox.Show("The operation could not be completed. Please try again.\n\n" + e.Exception.Message, "Meeting Capture", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            Application.Run(new SetupForm(new Preferences()));
        }
    }

    internal static class Native
    {
        [DllImport("user32.dll")] private static extern bool SetProcessDpiAwarenessContext(IntPtr context);
        [DllImport("shcore.dll")] private static extern int SetProcessDpiAwareness(int awareness);
        [DllImport("user32.dll")] private static extern bool SetProcessDPIAware();
        [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] internal static extern bool SetForegroundWindow(IntPtr handle);
        [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr handle);
        [DllImport("user32.dll")] internal static extern bool GetPhysicalCursorPos(out Point point);
        [DllImport("user32.dll")] internal static extern bool GetCursorPos(out Point point);
        [DllImport("user32.dll")] internal static extern bool GetWindowRect(IntPtr handle, out RECT rect);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
        [DllImport("user32.dll")] internal static extern bool ReleaseCapture();
        [DllImport("user32.dll")] internal static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wparam, IntPtr lparam);
        [DllImport("user32.dll")] internal static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int w, int height, uint flags);
        [DllImport("dwmapi.dll")] private static extern int DwmFlush();
        [DllImport("user32.dll")] internal static extern IntPtr GetDC(IntPtr window);
        [DllImport("user32.dll")] internal static extern int ReleaseDC(IntPtr window, IntPtr dc);
        [DllImport("gdi32.dll", SetLastError = true)] internal static extern bool BitBlt(IntPtr destination, int x, int y, int width, int height, IntPtr source, int sx, int sy, uint operation);
        [DllImport("user32.dll")] private static extern bool EnumDisplayMonitors(IntPtr dc, IntPtr clip, MonitorCallback callback, IntPtr data);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool GetMonitorInfo(IntPtr monitor, ref MONITORINFOEX info);
        [DllImport("shcore.dll")] private static extern int GetDpiForMonitor(IntPtr monitor, int type, out uint x, out uint y);
        private delegate bool MonitorCallback(IntPtr monitor, IntPtr dc, ref RECT bounds, IntPtr data);

        [StructLayout(LayoutKind.Sequential)] internal struct RECT
        {
            public int Left, Top, Right, Bottom;
            public Rectangle Rectangle { get { return Rectangle.FromLTRB(Left, Top, Right, Bottom); } }
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct MONITORINFOEX
        {
            public int Size;
            public RECT Monitor, Work;
            public uint Flags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string Device;
        }

        internal static void ConfigureDpi()
        {
            try { if (SetProcessDpiAwarenessContext(new IntPtr(-4))) return; }
            catch (EntryPointNotFoundException) { }
            try { if (SetProcessDpiAwareness(2) == 0) return; }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }
            SetProcessDPIAware();
        }
        internal static void FlushDesktop() { try { DwmFlush(); } catch (DllNotFoundException) { } }
        internal static int WindowDpi(IntPtr window)
        {
            try { uint dpi = GetDpiForWindow(window); if (dpi != 0) return (int)dpi; } catch (EntryPointNotFoundException) { }
            using (var g = Graphics.FromHwnd(window)) return (int)g.DpiX;
        }
        internal static Point CursorPoint()
        {
            Point p;
            if (!GetPhysicalCursorPos(out p)) GetCursorPos(out p);
            return p;
        }
        internal static List<DisplayInfo> Displays()
        {
            var result = new List<DisplayInfo>();
            MonitorCallback callback = delegate(IntPtr monitor, IntPtr dc, ref RECT bounds, IntPtr data)
            {
                var info = new MONITORINFOEX();
                info.Size = Marshal.SizeOf(typeof(MONITORINFOEX));
                if (!GetMonitorInfo(monitor, ref info)) throw new InvalidOperationException("Unable to read monitor information.");
                uint dx = 96, dy = 96;
                try { GetDpiForMonitor(monitor, 0, out dx, out dy); } catch (DllNotFoundException) { } catch (EntryPointNotFoundException) { }
                result.Add(new DisplayInfo(info.Device, info.Monitor.Rectangle, info.Work.Rectangle, dx, dy));
                return true;
            };
            if (!EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, callback, IntPtr.Zero) || result.Count == 0)
                throw new InvalidOperationException("Unable to read the display layout.");
            return result.OrderBy(d => d.Device, StringComparer.Ordinal).ToList();
        }
    }

    internal sealed class DisplayInfo
    {
        internal readonly string Device;
        internal readonly Rectangle Bounds, Work;
        internal readonly uint DpiX, DpiY;
        internal DisplayInfo(string device, Rectangle bounds, Rectangle work, uint dx, uint dy)
        { Device = device; Bounds = bounds; Work = work; DpiX = dx; DpiY = dy; }
        internal string Signature { get { return Device + ":" + Bounds + ":" + DpiX + ":" + DpiY; } }
        internal static string Fingerprint(IEnumerable<DisplayInfo> displays) { return String.Join("|", displays.Select(d => d.Signature)); }
    }

    internal static class FolderInput
    {
        internal static string Normalize(string input)
        {
            string path = (input ?? "").Trim();
            // Strip matching OUTER quotes only; keep spaces and quotes inside names.
            while (path.Length >= 2)
            {
                char first = path[0], last = path[path.Length - 1];
                if (!((first == '"' && last == '"') || (first == '\'' && last == '\'') || (first == '“' && last == '”')))
                    break;
                path = path.Substring(1, path.Length - 2).Trim();
            }
            return path;
        }
    }

    internal sealed class Preferences
    {
        private readonly string path;
        internal string DirectoryPath;
        internal Preferences() : this(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MeetingCapture", "last-folder.txt")) { }
        internal Preferences(string file)
        {
            path = file;
            try { if (File.Exists(path)) DirectoryPath = FolderInput.Normalize(File.ReadAllText(path, Encoding.UTF8)); }
            catch (IOException) { } catch (UnauthorizedAccessException) { }
            if (String.IsNullOrEmpty(DirectoryPath)) DirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
        }
        internal bool Save(out string error)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, DirectoryPath, Encoding.UTF8);
                error = null; return true;
            }
            catch (Exception e) { error = e.Message; return false; }
        }
    }

    internal static class PngStore
    {
        internal static void ValidateDirectory(string directory)
        {
            if (String.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
                throw new DirectoryNotFoundException("The output folder does not exist. Check the path or browse for a folder.");
            string probe = Path.Combine(directory, ".meetingcapture-" + Guid.NewGuid().ToString("N") + ".tmp");
            using (var stream = new FileStream(probe, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1, FileOptions.DeleteOnClose))
                stream.WriteByte(0);
        }
        internal static string Save(Bitmap image, string directory, int sequence, DateTime timestamp)
        {
            if (!Directory.Exists(directory)) throw new DirectoryNotFoundException("The output folder is no longer available. End this session and choose another folder.");
            string stem = timestamp.ToString("yyyyMMdd_HHmmss_fff", CultureInfo.InvariantCulture) + "_" + sequence.ToString("D4", CultureInfo.InvariantCulture);
            string temporary = Path.Combine(directory, ".meetingcapture-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    image.Save(stream, ImageFormat.Png);
                    stream.Flush(true);
                }
                // File.Move without overwrite provides an atomic publish on the same volume.
                for (int collision = 0; collision < 10000; collision++)
                {
                    string final = Path.Combine(directory, stem + (collision == 0 ? "" : "_" + collision.ToString("D2", CultureInfo.InvariantCulture)) + ".png");
                    try { File.Move(temporary, final); return final; }
                    catch (IOException) { if (!File.Exists(final)) throw; }
                }
                throw new IOException("Too many files share this name. Please try again.");
            }
            finally { try { if (File.Exists(temporary)) File.Delete(temporary); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
        }
    }

    internal static class ScreenCapture
    {
        internal static Bitmap Grab(Rectangle region)
        {
            if (region.Width < 2 || region.Height < 2) throw new ArgumentException("The capture region is too small. Select a larger region.");
            var image = new Bitmap(region.Width, region.Height, PixelFormat.Format24bppRgb);
            try
            {
                using (var graphics = Graphics.FromImage(image))
                {
                    IntPtr source = Native.GetDC(IntPtr.Zero);
                    if (source == IntPtr.Zero) throw new System.ComponentModel.Win32Exception("Unable to read the screen.");
                    try
                    {
                        IntPtr destination = graphics.GetHdc();
                        try
                        {
                            if (!Native.BitBlt(destination, 0, 0, region.Width, region.Height, source, region.X, region.Y, 0x40CC0020))
                                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                        }
                        finally { graphics.ReleaseHdc(destination); }
                    }
                    finally { Native.ReleaseDC(IntPtr.Zero, source); }
                }
                return image;
            }
            catch { image.Dispose(); throw; }
        }
    }

    internal static class Style
    {
        internal static readonly Color Ink = Color.FromArgb(30, 44, 61);
        internal static readonly Color Muted = Color.FromArgb(98, 113, 130);
        internal static readonly Color Blue = Color.FromArgb(36, 95, 217);
        internal static readonly Color Canvas = Color.FromArgb(245, 247, 251);
        internal static readonly Color Red = Color.FromArgb(178, 48, 48);
        internal static Button Button(string text, bool primary, bool noFocus = false)
        {
            Button b = noFocus ? new NoFocusButton() : new Button();
            b.Text = text; b.AutoSize = false; b.FlatStyle = FlatStyle.Flat; b.Cursor = Cursors.Hand; b.TabStop = !noFocus;
            b.FlatAppearance.BorderSize = primary ? 0 : 1;
            b.FlatAppearance.BorderColor = Color.FromArgb(212, 220, 232);
            b.BackColor = primary ? Blue : Color.White;
            b.ForeColor = primary ? Color.White : Ink;
            b.Font = new Font("Segoe UI", 10F, primary ? FontStyle.Bold : FontStyle.Regular);
            return b;
        }
        internal static Label Label(string text, float size, bool bold)
        {
            return new Label { Text = text, AutoSize = false, ForeColor = Ink, Font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular), TextAlign = ContentAlignment.MiddleLeft };
        }
        internal static void IconFor(Form form)
        {
            try { form.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
        }
    }

    // Native pushbuttons can call SetFocus even on a WS_EX_NOACTIVATE parent.
    // Consume their mouse messages so a click never activates the presentation's tool.
    internal sealed class NoFocusButton : Button
    {
        private bool pressed;
        internal NoFocusButton() { SetStyle(ControlStyles.Selectable, false); TabStop = false; }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0021) { m.Result = new IntPtr(3); return; }
            if (m.Msg == 0x0201 || m.Msg == 0x0203)
            { if (Enabled) { pressed = true; Capture = true; Invalidate(); } m.Result = IntPtr.Zero; return; }
            if (m.Msg == 0x0202)
            {
                bool fire = pressed && Enabled;
                pressed = false; Capture = false; Invalidate();
                int packed = m.LParam.ToInt32();
                if (fire && ClientRectangle.Contains(new Point((short)(packed & 65535), (short)(packed >> 16)))) OnClick(EventArgs.Empty);
                m.Result = IntPtr.Zero; return;
            }
            if (m.Msg == 0x0215) pressed = false;
            base.WndProc(ref m);
        }
    }

    // While the toolbar is hidden, swallow additional left-click pairs at its old
    // position. Otherwise a user's double-click can advance the slide underneath.
    // The hook exists only for the duration of a capture and records no input.
    internal sealed class CaptureClickGuard : IDisposable
    {
        private delegate IntPtr MouseHook(int code, IntPtr message, IntPtr data);
        [StructLayout(LayoutKind.Sequential)] private struct MouseData
        { public Point Point; public uint Mouse, Flags, Time; public IntPtr Extra; }
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetWindowsHookEx(int type, MouseHook callback, IntPtr module, uint thread);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr message, IntPtr data);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandle(string name);
        private readonly Rectangle bounds;
        private readonly MouseHook callback;
        private IntPtr hook;
        private bool swallowPair;
        internal CaptureClickGuard(Rectangle toolbarBounds)
        {
            bounds = toolbarBounds; callback = Filter;
            hook = SetWindowsHookEx(14, callback, GetModuleHandle(null), 0);
            if (hook == IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Unable to enable protection against repeated clicks.");
        }
        private IntPtr Filter(int code, IntPtr message, IntPtr data)
        {
            if (code >= 0)
            {
                int type = message.ToInt32();
                if (type == 0x0201)
                {
                    var mouse = (MouseData)Marshal.PtrToStructure(data, typeof(MouseData));
                    if (bounds.Contains(mouse.Point)) { swallowPair = true; return new IntPtr(1); }
                }
                if (type == 0x0202 && swallowPair) { swallowPair = false; return new IntPtr(1); }
            }
            return CallNextHookEx(hook, code, message, data);
        }
        public void Dispose()
        {
            if (hook != IntPtr.Zero) { UnhookWindowsHookEx(hook); hook = IntPtr.Zero; }
            GC.KeepAlive(callback);
        }
    }

    // Explicit pixel fonts and logical bounds avoid Framework's optional app.config
    // DPI switches and keep the distributed application a single EXE.
    internal abstract class DpiForm : Form
    {
        private sealed class LayoutItem
        {
            internal Control Control;
            internal Rectangle Bounds;
            internal string Family;
            internal float Points;
            internal FontStyle Style;
        }
        private List<LayoutItem> layout;
        private Size logicalSize;
        private readonly List<Font> scaledFonts = new List<Font>();
        protected override void OnLoad(EventArgs e)
        {
            AutoScaleMode = AutoScaleMode.None;
            logicalSize = ClientSize;
            layout = Controls.Cast<Control>().Select(c => new LayoutItem { Control = c, Bounds = c.Bounds, Family = c.Font.FontFamily.Name, Points = c.Font.SizeInPoints, Style = c.Font.Style }).ToList();
            ApplyDpi(Native.WindowDpi(Handle)); base.OnLoad(e);
        }
        private void ApplyDpi(int dpi)
        {
            float scale = dpi / 96F;
            var old = scaledFonts.ToArray(); scaledFonts.Clear();
            SuspendLayout();
            ClientSize = new Size((int)Math.Round(logicalSize.Width * scale), (int)Math.Round(logicalSize.Height * scale));
            foreach (var item in layout)
            {
                var font = new Font(item.Family, item.Points * 96F / 72F * scale, item.Style, GraphicsUnit.Pixel);
                scaledFonts.Add(font); item.Control.Font = font;
                Rectangle r = item.Bounds;
                item.Control.SetBounds((int)Math.Round(r.X * scale), (int)Math.Round(r.Y * scale), (int)Math.Round(r.Width * scale), (int)Math.Round(r.Height * scale));
            }
            ResumeLayout(false);
            foreach (Font font in old) font.Dispose();
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x02E0 && layout != null)
            {
                var suggested = (Native.RECT)Marshal.PtrToStructure(m.LParam, typeof(Native.RECT));
                ApplyDpi(m.WParam.ToInt32() & 65535);
                Native.SetWindowPos(Handle, IntPtr.Zero, suggested.Left, suggested.Top, Width, Height, 0x0010 | 0x0004);
                m.Result = IntPtr.Zero; return;
            }
            base.WndProc(ref m);
        }
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing) { foreach (var font in scaledFonts) font.Dispose(); scaledFonts.Clear(); }
        }
    }

    internal sealed class SetupForm : DpiForm
    {
        private readonly Preferences preferences;
        internal readonly TextBox FolderBox;
        internal readonly Label RegionLabel, StatusLabel;
        internal readonly Button SelectButton, StartButton, FolderButton;
        internal Rectangle? CaptureArea { get; private set; }
        private string displayFingerprint;
        internal CaptureBar CurrentBar;
        private bool closing;
        internal SetupForm(Preferences settings)
        {
            preferences = settings;
            Text = "Meeting Capture"; BackColor = Style.Canvas; ForeColor = Style.Ink;
            Font = new Font("Segoe UI", 10F);
            AutoScaleMode = AutoScaleMode.None;
            ClientSize = new Size(680, 492); FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false; StartPosition = FormStartPosition.CenterScreen;
            Style.IconFor(this);

            var title = Style.Label("Meeting Capture", 24, true); title.SetBounds(32, 22, 470, 52); Controls.Add(title);
            var subtitle = Style.Label("Choose once. Capture throughout your meeting.", 10, false); subtitle.ForeColor = Style.Muted; subtitle.SetBounds(34, 79, 610, 30); Controls.Add(subtitle);
            var directoryTitle = Style.Label("01   Save screenshots to", 11, true); directoryTitle.SetBounds(34, 134, 400, 30); Controls.Add(directoryTitle);
            FolderBox = new TextBox { Text = preferences.DirectoryPath, ReadOnly = false, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, AccessibleName = "Output folder" };
            FolderBox.SetBounds(34, 177, 470, 30); Controls.Add(FolderBox);
            FolderBox.Leave += delegate { NormalizeFolderInput(); };
            FolderButton = Style.Button("Browse...", false); FolderButton.SetBounds(520, 172, 124, 40); FolderButton.Click += delegate { ChooseFolder(); }; Controls.Add(FolderButton);
            var directoryNote = Style.Label("Paste a folder path (quotes are OK). PNG files go directly here.", 9, false); directoryNote.ForeColor = Style.Muted; directoryNote.SetBounds(34, 215, 610, 27); Controls.Add(directoryNote);
            var regionTitle = Style.Label("02   Capture region", 11, true); regionTitle.SetBounds(34, 265, 400, 30); Controls.Add(regionTitle);
            RegionLabel = Style.Label("No region selected - choose once per session", 10, false); RegionLabel.ForeColor = Style.Muted; RegionLabel.SetBounds(34, 308, 470, 40); Controls.Add(RegionLabel);
            SelectButton = Style.Button("Select region...", false); SelectButton.SetBounds(520, 308, 124, 40); SelectButton.Click += async delegate { await SelectRegionAsync(); }; Controls.Add(SelectButton);
            StartButton = Style.Button("Start session", true); StartButton.SetBounds(34, 387, 610, 48); StartButton.Enabled = false; StartButton.Click += delegate { StartMeeting(); }; Controls.Add(StartButton);
            StatusLabel = Style.Label("A small floating toolbar appears when you start.", 9, false); StatusLabel.ForeColor = Style.Muted; StatusLabel.SetBounds(34, 446, 610, 30); Controls.Add(StatusLabel);
            SystemEvents.DisplaySettingsChanged += DisplayChanged;
            FormClosing += delegate { closing = true; if (CurrentBar != null) CurrentBar.Close(); };
            FormClosed += delegate { SystemEvents.DisplaySettingsChanged -= DisplayChanged; };
            Shown += delegate
            {
                if (!Directory.Exists(FolderBox.Text)) SetStatus("The output folder is unavailable. Paste a valid path or browse.", true);
            };
        }
        internal string NormalizeFolderInput()
        {
            string directory = FolderInput.Normalize(FolderBox.Text);
            if (FolderBox.Text != directory) FolderBox.Text = directory;
            return directory;
        }
        private void ChooseFolder()
        {
            using (var dialog = new FolderBrowserDialog { Description = "Choose an output folder for this session", ShowNewFolderButton = true })
            {
                string typed = NormalizeFolderInput();
                if (Directory.Exists(typed)) dialog.SelectedPath = typed;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    PngStore.ValidateDirectory(dialog.SelectedPath);
                    FolderBox.Text = dialog.SelectedPath;
                    preferences.DirectoryPath = dialog.SelectedPath;
                    string error;
                    if (!preferences.Save(out error)) SetStatus("The folder is usable, but its path could not be remembered: " + error, true);
                    else SetStatus("Folder ready. Screenshots will be saved directly here.", false);
                }
                catch (Exception e) { SetStatus("Cannot write to this folder: " + e.Message, true); }
            }
        }
        internal async Task SelectRegionAsync()
        {
            Hide();
            try
            {
                await Task.Delay(180);
                Native.FlushDesktop();
                using (var selector = new RegionSelector())
                {
                    if (selector.ShowDialog() == DialogResult.OK)
                        SetRegion(selector.Selection, selector.Fingerprint);
                }
            }
            catch (Exception e) { SetStatus("Region selection failed: " + e.Message, true); }
            finally { if (!IsDisposed && !closing) { Show(); Activate(); } }
        }
        internal void SetRegion(Rectangle rect, string fingerprint)
        {
            CaptureArea = rect; displayFingerprint = fingerprint;
            RegionLabel.Text = String.Format("{0} x {1} pixels    |    Top-left ({2}, {3})", rect.Width, rect.Height, rect.X, rect.Y);
            RegionLabel.ForeColor = Style.Ink; StartButton.Enabled = true;
            SetStatus("Region ready. You can start your session.", false);
        }
        internal bool StartMeeting()
        {
            if (!CaptureArea.HasValue) return false;
            try
            {
                string current = DisplayInfo.Fingerprint(Native.Displays());
                if (current != displayFingerprint) { InvalidateRegion(); return false; }
                string directory = NormalizeFolderInput();
                PngStore.ValidateDirectory(directory);
                preferences.DirectoryPath = directory;
                string error; preferences.Save(out error);
                CurrentBar = new CaptureBar(CaptureArea.Value, directory, current);
                CurrentBar.FormClosed += delegate
                {
                    CurrentBar = null;
                    if (closing || IsDisposed) return;
                    CaptureArea = null; displayFingerprint = null; StartButton.Enabled = false;
                    RegionLabel.Text = "No region selected - choose once per session";
                    SetStatus("Session ended. Screenshots are saved. Select a region to start again.", false);
                    Show(); Activate();
                };
                Hide(); CurrentBar.Show(); return true;
            }
            catch (Exception e) { SetStatus("Cannot start: " + e.Message, true); return false; }
        }
        private void DisplayChanged(object sender, EventArgs e)
        {
            if (IsDisposed || !IsHandleCreated) return;
            try { BeginInvoke((Action)delegate { if (CurrentBar == null) InvalidateRegion(); }); } catch (InvalidOperationException) { }
        }
        private void InvalidateRegion()
        {
            CaptureArea = null; displayFingerprint = null; StartButton.Enabled = false;
            RegionLabel.Text = "Display settings changed - select a new region";
            SetStatus("Select the region again after a display layout or resolution change.", true);
        }
        private void SetStatus(string text, bool error) { StatusLabel.Text = text; StatusLabel.ForeColor = error ? Style.Red : Style.Muted; }
    }

    internal sealed class RegionSelector : Form
    {
        private readonly Bitmap desktop;
        private readonly Rectangle virtualBounds;
        private readonly List<DisplayInfo> displays;
        private Point? start;
        private Point finish;
        internal Rectangle Selection;
        internal bool IsSelecting { get { return start.HasValue; } }
        internal readonly string Fingerprint;
        internal RegionSelector()
        {
            displays = Native.Displays(); Fingerprint = DisplayInfo.Fingerprint(displays);
            virtualBounds = displays.Select(d => d.Bounds).Aggregate(Rectangle.Union);
            desktop = ScreenCapture.Grab(virtualBounds);
            Text = "Meeting Capture - Select region";
            FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; TopMost = true;
            StartPosition = FormStartPosition.Manual; AutoScaleMode = AutoScaleMode.None;
            Bounds = virtualBounds; Cursor = Cursors.Cross; KeyPreview = true; DoubleBuffered = true;
            Shown += delegate { Bounds = virtualBounds; Activate(); Focus(); };
        }
        internal static Rectangle Normalize(Point a, Point b)
        { return Rectangle.FromLTRB(Math.Min(a.X, b.X), Math.Min(a.Y, b.Y), Math.Max(a.X, b.X), Math.Max(a.Y, b.Y)); }
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button == MouseButtons.Right) { DialogResult = DialogResult.Cancel; Close(); return; }
            if (e.Button != MouseButtons.Left) return;
            start = e.Location; finish = start.Value; Capture = true; Invalidate();
        }
        protected override void OnMouseMove(MouseEventArgs e)
        { base.OnMouseMove(e); if (start.HasValue) { finish = e.Location; Invalidate(); } }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button != MouseButtons.Left || !start.HasValue) return;
            finish = e.Location; Capture = false;
            Rectangle local = Rectangle.Intersect(new Rectangle(Point.Empty, virtualBounds.Size), Normalize(start.Value, finish));
            if (local.Width < 2 || local.Height < 2) { start = null; Invalidate(); return; }
            Selection = new Rectangle(local.X + virtualBounds.X, local.Y + virtualBounds.Y, local.Width, local.Height);
            DialogResult = DialogResult.OK; Close();
        }
        protected override void OnKeyDown(KeyEventArgs e)
        { if (e.KeyCode == Keys.Escape) { e.Handled = true; DialogResult = DialogResult.Cancel; Close(); } base.OnKeyDown(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.DrawImageUnscaled(desktop, 0, 0);
            using (var dim = new SolidBrush(Color.FromArgb(130, 10, 20, 35))) e.Graphics.FillRectangle(dim, ClientRectangle);
            if (start.HasValue)
            {
                var rect = Rectangle.Intersect(ClientRectangle, Normalize(start.Value, finish));
                if (rect.Width > 0 && rect.Height > 0)
                {
                    e.Graphics.DrawImage(desktop, rect, rect, GraphicsUnit.Pixel);
                    using (var pen = new Pen(Color.FromArgb(77, 162, 255), 2)) e.Graphics.DrawRectangle(pen, rect.X, rect.Y, rect.Width - 1, rect.Height - 1);
                    string size = rect.Width + " × " + rect.Height + " pixels";
                    var label = new Rectangle(Math.Max(0, Math.Min(ClientSize.Width - 220, rect.X)), Math.Max(0, rect.Y - 34), 220, 30);
                    e.Graphics.FillRectangle(Brushes.Black, label);
                    TextRenderer.DrawText(e.Graphics, size, Font, label, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                }
            }
            // Keep the instruction on the monitor containing the pointer, not in a gap between monitors.
            var screen = displays.FirstOrDefault(d => d.Bounds.Contains(Native.CursorPoint()));
            Rectangle monitor = screen == null ? virtualBounds : screen.Bounds;
            var hint = new Rectangle(monitor.X - virtualBounds.X + 24, monitor.Y - virtualBounds.Y + 24, Math.Min(740, monitor.Width - 48), 50);
            e.Graphics.FillRectangle(Brushes.Black, hint);
            TextRenderer.DrawText(e.Graphics, "Drag to select | Release to confirm | Esc / right-click to cancel", Font, hint, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
        protected override void Dispose(bool disposing) { if (disposing) desktop.Dispose(); base.Dispose(disposing); }
    }

    internal sealed class CaptureBar : DpiForm
    {
        internal readonly Button CaptureButton, OpenButton, EndButton;
        internal readonly Label CountLabel, StatusLabel;
        internal readonly Rectangle CaptureArea;
        internal readonly string DirectoryPath;
        internal int SavedCount { get; private set; }
        internal bool Busy { get; private set; }
        internal bool Paused { get; private set; }
        internal string LastFile { get; private set; }
        internal string LastError { get; private set; }
        private readonly string fingerprint;
        private readonly Timer displayTimer;
        private bool closeAfterCapture;
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override CreateParams CreateParams
        {
            get { var cp = base.CreateParams; cp.ExStyle |= 0x08000000 | 0x00000080; return cp; } // NOACTIVATE | TOOLWINDOW
        }
        internal CaptureBar(Rectangle region, string directory, string displayFingerprint)
        {
            CaptureArea = region; DirectoryPath = directory; fingerprint = displayFingerprint;
            Text = "Meeting Capture - Toolbar"; FormBorderStyle = FormBorderStyle.None;
            TopMost = true; ShowInTaskbar = false; StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.None;
            ClientSize = new Size(554, 98); BackColor = Color.FromArgb(235, 240, 249);
            Font = new Font("Segoe UI", 9F); Style.IconFor(this);
            var grip = Style.Label("⠿", 18, true); grip.SetBounds(6, 5, 28, 52); grip.Cursor = Cursors.SizeAll;
            grip.MouseDown += DragBar; Controls.Add(grip);
            CaptureButton = Style.Button("Capture", true, true); CaptureButton.SetBounds(38, 12, 100, 44); CaptureButton.TabStop = false;
            CaptureButton.Click += async delegate { await CaptureAsync(); }; Controls.Add(CaptureButton);
            CountLabel = Style.Label("Saved: 0", 10, true); CountLabel.SetBounds(150, 12, 134, 44); CountLabel.MouseDown += DragBar; Controls.Add(CountLabel);
            OpenButton = Style.Button("Open folder", false, true); OpenButton.SetBounds(290, 12, 116, 44); OpenButton.TabStop = false;
            OpenButton.Click += delegate { OpenDirectory(); }; Controls.Add(OpenButton);
            EndButton = Style.Button("End session", false, true); EndButton.SetBounds(418, 12, 124, 44); EndButton.TabStop = false;
            EndButton.Click += delegate { Close(); }; Controls.Add(EndButton);
            StatusLabel = Style.Label("Region fixed | Drag the left handle to move", 9, false); StatusLabel.ForeColor = Style.Muted; StatusLabel.SetBounds(16, 62, 526, 29);
            StatusLabel.MouseDown += DragBar; Controls.Add(StatusLabel);
            MouseDown += DragBar;
            var work = Native.Displays().OrderByDescending(d => { var r = Rectangle.Intersect(d.Bounds, region); return (long)r.Width * r.Height; }).First().Work;
            Location = new Point(work.Right - Width - 20, work.Top + 24);
            displayTimer = new Timer { Interval = 1000 }; displayTimer.Tick += delegate { CheckDisplay(); };
            SystemEvents.DisplaySettingsChanged += DisplayChanged;
            Shown += delegate { ClampToScreen(); displayTimer.Start(); };
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (Busy) { closeAfterCapture = true; e.Cancel = true; } };
            FormClosed += delegate { displayTimer.Stop(); displayTimer.Dispose(); SystemEvents.DisplaySettingsChanged -= DisplayChanged; };
        }
        protected override void SetVisibleCore(bool value)
        {
            IntPtr foreground = Native.GetForegroundWindow();
            base.SetVisibleCore(value);
            // Framework may select the form itself when reshown on the same GUI
            // thread. Preserve the original foreground if that happens.
            if (value && foreground != IntPtr.Zero && foreground != Handle && Native.IsWindow(foreground) && Native.GetForegroundWindow() == Handle)
                Native.SetForegroundWindow(foreground);
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0021) { m.Result = new IntPtr(3); return; } // MA_NOACTIVATE, still deliver click
            if (m.Msg == 0x02E0 && IsHandleCreated)
            {
                base.WndProc(ref m); ClampToScreen(); return;
            }
            base.WndProc(ref m);
        }
        private void DragBar(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || Busy) return;
            Native.ReleaseCapture(); Native.SendMessage(Handle, 0x00A1, new IntPtr(2), IntPtr.Zero);
            ClampToScreen();
        }
        private void ClampToScreen()
        {
            Rectangle work = Screen.FromRectangle(Bounds).WorkingArea;
            int x = Math.Max(work.Left, Math.Min(Left, work.Right - Width));
            int y = Math.Max(work.Top, Math.Min(Top, work.Bottom - Height));
            Native.SetWindowPos(Handle, new IntPtr(-1), x, y, 0, 0, 0x0010 | 0x0001);
        }
        private void DisplayChanged(object sender, EventArgs e)
        {
            if (IsDisposed || !IsHandleCreated) return;
            try { BeginInvoke((Action)delegate { PauseForDisplay(); }); } catch (InvalidOperationException) { }
        }
        internal bool CheckDisplay()
        {
            if (Paused || IsDisposed) return false;
            try { if (DisplayInfo.Fingerprint(Native.Displays()) == fingerprint) return true; }
            catch { }
            PauseForDisplay(); return false;
        }
        internal void PauseForDisplay()
        {
            if (IsDisposed) return;
            Paused = true; CaptureButton.Enabled = false;
            LastError = "Display changed. End this session and select the region again.";
            StatusLabel.Text = LastError; StatusLabel.ForeColor = Style.Red;
            ClampToScreen();
        }
        internal async Task<bool> CaptureAsync()
        {
            if (Busy || Paused || IsDisposed) return false;
            if (!CheckDisplay()) return false;
            Busy = true; CaptureButton.Enabled = false; OpenButton.Enabled = false; EndButton.Enabled = false;
            LastError = null;
            DateTime timestamp;
            Bitmap image = null;
            CaptureClickGuard clickGuard = null;
            bool saved = false;
            try
            {
                clickGuard = new CaptureClickGuard(Bounds);
                Hide();
                await Task.Delay(180);
                if (!CheckDisplay()) return false;
                Native.FlushDesktop(); timestamp = DateTime.Now;
                image = ScreenCapture.Grab(CaptureArea);
                Bitmap captured = image;
                string file = await Task.Run(() => PngStore.Save(captured, DirectoryPath, SavedCount + 1, timestamp));
                LastFile = file; SavedCount++; saved = true;
                CountLabel.Text = "Saved: " + SavedCount;
                if (!Paused)
                {
                    StatusLabel.Text = "Saved: " + Path.GetFileName(file);
                    StatusLabel.ForeColor = Color.FromArgb(23, 112, 80);
                    errorTip.SetToolTip(StatusLabel, null);
                }
            }
            catch (Exception e)
            {
                LastError = "Save failed: " + e.Message;
                StatusLabel.Text = LastError; StatusLabel.AutoEllipsis = true; StatusLabel.ForeColor = Style.Red;
                errorTip.SetToolTip(StatusLabel, LastError);
            }
            finally
            {
                if (image != null) image.Dispose();
                try
                {
                    Busy = false;
                    if (!IsDisposed)
                    {
                        CaptureButton.Enabled = !Paused; OpenButton.Enabled = true; EndButton.Enabled = true;
                        if (closeAfterCapture) Close(); else Show();
                    }
                }
                finally { if (clickGuard != null) clickGuard.Dispose(); }
            }
            return saved;
        }
        private readonly ToolTip errorTip = new ToolTip { AutoPopDelay = 20000, InitialDelay = 100, ReshowDelay = 100 };
        private void OpenDirectory()
        {
            try
            {
                if (!Directory.Exists(DirectoryPath)) throw new DirectoryNotFoundException("The output folder is no longer available.");
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe"),
                    Arguments = "/n,\"" + DirectoryPath + "\"", UseShellExecute = true
                });
            }
            catch (Exception e) { StatusLabel.Text = "Cannot open the output folder: " + e.Message; StatusLabel.ForeColor = Style.Red; }
        }
        protected override void Dispose(bool disposing) { if (disposing) errorTip.Dispose(); base.Dispose(disposing); }
    }
}
