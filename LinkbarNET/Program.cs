using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json; // alternativ Newtonsoft.Json verwenden
using System.Threading;
using System.Windows.Forms;

namespace LinkbarNET
{
    public enum DockPosition
    {
        Left,
        Right,
        Top,
        Bottom
    }

    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            bool createdNew;
            // Wähle einen eindeutigen Namen für den Mutex (z.B. der Anwendungsname)
            using (Mutex mutex = new Mutex(true, "DeineAnwendungsEindeutigerName", out createdNew))
            {
                if (!createdNew)
                {
                    MessageBox.Show("Das Programm läuft bereits!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                // Standardstartposition: Oben (wird ggf. durch gespeicherte Einstellungen überschrieben)
                Application.Run(new Linkbar(DockPosition.Top));
            }
        }
    }

    public class Linkbar : Form
    {



        #region AppBar-API

        // AppBar-Konstanten
        private const uint ABM_NEW = 0x00000000;
        private const uint ABM_REMOVE = 0x00000001;
        private const uint ABM_QUERYPOS = 0x00000002;
        private const uint ABM_SETPOS = 0x00000003;

        private const int ABE_LEFT = 0;
        private const int ABE_TOP = 1;
        private const int ABE_RIGHT = 2;
        private const int ABE_BOTTOM = 3;

        // Wird beim Registrieren des AppBars gesetzt und in WndProc ausgewertet
        private uint _uCallBackMsg;

        [StructLayout(LayoutKind.Sequential)]
        private struct APPBARDATA
        {
            public int cbSize;
            public IntPtr hWnd;
            public uint uCallbackMessage;
            public uint uEdge;
            public RECT rc;
            public int lParam;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("shell32.dll")]
        private static extern uint SHAppBarMessage(uint dwMessage, ref APPBARDATA pData);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern uint RegisterWindowMessage(string lpString);

        #endregion

        #region WindowPos-API (HWND_TOPMOST-Hack)

        private const int HWND_TOPMOST = -1;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_SHOWWINDOW = 0x0040;

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        private static extern bool SetWindowPos(
            IntPtr hWnd,
            IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy,
            uint uFlags);


        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            // Jetzt ist der Handle vorhanden, also kann der AppBar registriert werden
            RegistriereAppBar();

            // Setze die Position, wenn nötig
            SetzeAppBarPosition();

            // Verknüpfungen laden
            LadeVerknuepfungen(true);

            // Dock-Position umsetzen (ÄndereDockPosition lädt zudem die Icons neu)
            ÄndereDockPosition(this.dockPosition);
        }

        private AppSettings appSettings;

        /// <summary>
        /// Versucht das Fenster nach vorn in die TopMost-Z-Reihenfolge zu bringen.
        /// </summary>
        private void ForceTopMost()
        {
            // Zusätzliche Sicherheit:
            this.TopMost = true;

            SetWindowPos(
                this.Handle,
                (IntPtr)HWND_TOPMOST,
                0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
            );
        }

        #endregion

        #region Felder & Konstruktor

        // Der Shortcut-Pfad wird nun nicht mehr fest codiert, sondern aus den Settings geladen:
        private string shortcutPath;

        private int BarSize = 30;

        private DockPosition dockPosition;
        private Screen aktuellerMonitor;

        // Konstanten für die Icon-Größen:
        private const int ICON_SIZE_NORMAL_WIDTH = 20;
        private const int ICON_SIZE_NORMAL_HEIGHT = 20;
        private const int ICON_SIZE_HOVER_WIDTH = 25;
        private const int ICON_SIZE_HOVER_HEIGHT = 25;
        private const int ICON_SIZE_CLICK_WIDTH = 30;
        private const int ICON_SIZE_CLICK_HEIGHT = 30;

        // UI-Steuerelemente
        private FlowLayoutPanel panel;
        private ContextMenuStrip contextMenu;
        private ToolStripMenuItem monitorMenu;

        // Timer zum regelmäßigen Neuladen der Verknüpfungen
        private System.Windows.Forms.Timer reloadTimer;

        // ToolTip, um den Text beim Überfahren mit der Maus anzuzeigen
        private ToolTip iconToolTip = new ToolTip();

        private void OpenConfigForm()
        {
            using (var configForm = new ConfigForm())
            {
                // Zeige das Formular modal an.
                if (configForm.ShowDialog() == DialogResult.OK)
                {
                    // Übernehme den vom Benutzer ausgewählten Pfad
                    this.shortcutPath = configForm.ShortcutPath;

                    // Optional: Direkt die Einstellungen speichern, damit sie beim nächsten Start verfügbar sind
                    appSettings.DockPosition = this.dockPosition;
                    appSettings.MonitorDeviceName = this.aktuellerMonitor.DeviceName;
                    appSettings.ShortcutPath = this.shortcutPath;
                    SettingsManager.SaveSettings(appSettings);
                }
                else
                {
                    // Falls der Benutzer abbricht, kann die Anwendung beendet werden
                    Application.Exit();
                }
            }
        }
        public Linkbar(DockPosition position)
        {

            // Form-Eigenschaften
            this.FormBorderStyle = FormBorderStyle.None;
            this.TopMost = true;
            this.BackColor = Color.Black;
            this.ShowInTaskbar = false;
            this.ShowIcon = false;

            appSettings = SettingsManager.LoadSettings() ?? new AppSettings();
            if (appSettings.CustomIconMapping == null)
                appSettings.CustomIconMapping = new Dictionary<string, string>();

            // Gespeicherte Einstellungen laden (falls vorhanden)
            AppSettings loadedSettings = SettingsManager.LoadSettings();
            bool setingsfail = false;
            if (loadedSettings != null)
            {
                try
                {
                    this.dockPosition = loadedSettings.DockPosition;
                    // Versuche den Monitor anhand des gespeicherten DeviceNames zu finden:
                    Screen foundMonitor = Screen.AllScreens.FirstOrDefault(s => s.DeviceName == loadedSettings.MonitorDeviceName);
                    if (foundMonitor == null)
                    {
                        // Wenn der gespeicherte Monitor nicht existiert, verwende standardmäßig das erste Display
                        foundMonitor = Screen.AllScreens.FirstOrDefault() ?? Screen.PrimaryScreen;
                    }
                    this.aktuellerMonitor = foundMonitor;

                    // Shortcut-Pfad aus den Einstellungen laden oder Default-Wert nutzen
                    this.shortcutPath = loadedSettings.ShortcutPath;
                }
                catch
                {
                    setingsfail = true;
                };
            }
            else
            {
                // Standardwerte: über Parameter bzw. Primärmonitor und Default-Pfad
                this.dockPosition = position;
                this.aktuellerMonitor = Screen.PrimaryScreen;
                setingsfail = true;
            }

            if (setingsfail)
            {
                OpenConfigForm();
            };


            // Panel für die Icons
            panel = new FlowLayoutPanel
            {
                AutoScroll = true,
                // FlowDirection wird in ÄndereDockPosition gesetzt
                WrapContents = false,
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };
            this.Controls.Add(panel);

            // Kontextmenü erstellen
            ErstelleKontextMenü();

            // Rechtsklick-Ereignis des Panels abfangen
            panel.MouseUp += Panel_MouseUp;



            // Timer zum regelmäßigen Neuladen der Verknüpfungen
            reloadTimer = new System.Windows.Forms.Timer();
            reloadTimer.Interval = 60_000; // 60 Sekunden
            reloadTimer.Tick += (s, e) =>
            {
                LadeVerknuepfungen(false);
            };
            reloadTimer.Start();
        }

        #endregion

        #region MouseUp - Rechtsklick-Handling

        private void Panel_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                // Bestimme die absolute Position des Mausklicks
                Point absolutePos = panel.PointToScreen(e.Location);
                // Ermittle den Monitor, auf dem sich der Klick ereignet hat
                Screen currentScreen = Screen.FromPoint(absolutePos);
                // Nutze den WorkingArea, um Taskleiste etc. auszuschließen (optional)
                Rectangle workingArea = currentScreen.WorkingArea;

                // Ermittle die bevorzugte Größe des Kontextmenüs
                Size menuSize = contextMenu.GetPreferredSize(Size.Empty);

                // Berechne eine angepasste Position, damit das Menü innerhalb des Monitors bleibt
                int x = absolutePos.X;
                int y = absolutePos.Y;

                // Falls das Menü rechts über den Monitor hinausgehen würde, verschiebe es nach links
                if (x + menuSize.Width > workingArea.Right)
                {
                    x = workingArea.Right - menuSize.Width;
                }
                // Falls das Menü unten über den Monitor hinausgehen würde, verschiebe es nach oben
                if (y + menuSize.Height > workingArea.Bottom)
                {
                    y = workingArea.Bottom - menuSize.Height;
                }

                // Konvertiere die berechnete Position in Panel-Koordinaten
                Point adjustedPoint = panel.PointToClient(new Point(x, y));

                // Zeige das Kontextmenü ohne einen fixen DropDownDirection
                contextMenu.Show(panel, adjustedPoint);
            }
        }


        #endregion

        #region AppBar-Registrierung und -Entfernung

        private void RegistriereAppBar()
        {
            var abd = new APPBARDATA();
            abd.cbSize = Marshal.SizeOf(typeof(APPBARDATA));
            abd.hWnd = this.Handle;

            // Eindeutige Callback-Nachricht registrieren
            _uCallBackMsg = RegisterWindowMessage("AppBarMessage");
            abd.uCallbackMessage = _uCallBackMsg;

            // AppBar anmelden
            SHAppBarMessage(ABM_NEW, ref abd);

            // Position festlegen
            SetzeAppBarPosition();
        }

        private void EntferneAppBar()
        {
            var abd = new APPBARDATA();
            abd.cbSize = Marshal.SizeOf(typeof(APPBARDATA));
            abd.hWnd = this.Handle;
            SHAppBarMessage(ABM_REMOVE, ref abd);
        }

        #endregion

        #region AppBar-Positionierung

        private void SetzeAppBarPosition()
        {
            var abd = new APPBARDATA();
            abd.cbSize = Marshal.SizeOf(typeof(APPBARDATA));
            abd.hWnd = this.Handle;

            // Edge ermitteln und Koordinaten setzen
            switch (dockPosition)
            {
                case DockPosition.Left:
                    abd.uEdge = ABE_LEFT;
                    abd.rc.Left = aktuellerMonitor.Bounds.Left;
                    abd.rc.Top = aktuellerMonitor.Bounds.Top;
                    abd.rc.Right = abd.rc.Left + BarSize;
                    abd.rc.Bottom = aktuellerMonitor.Bounds.Bottom;
                    break;

                case DockPosition.Right:
                    abd.uEdge = ABE_RIGHT;
                    abd.rc.Right = aktuellerMonitor.Bounds.Right;
                    abd.rc.Top = aktuellerMonitor.Bounds.Top;
                    abd.rc.Left = abd.rc.Right - BarSize;
                    abd.rc.Bottom = aktuellerMonitor.Bounds.Bottom;
                    break;

                case DockPosition.Top:
                    abd.uEdge = ABE_TOP;
                    abd.rc.Left = aktuellerMonitor.Bounds.Left;
                    abd.rc.Top = aktuellerMonitor.Bounds.Top;
                    abd.rc.Right = aktuellerMonitor.Bounds.Right;
                    abd.rc.Bottom = abd.rc.Top + BarSize;
                    break;

                case DockPosition.Bottom:
                    abd.uEdge = ABE_BOTTOM;
                    abd.rc.Left = aktuellerMonitor.Bounds.Left;
                    abd.rc.Bottom = aktuellerMonitor.Bounds.Bottom;
                    abd.rc.Right = aktuellerMonitor.Bounds.Right;
                    abd.rc.Top = abd.rc.Bottom - BarSize;
                    break;
            }

            // Windows reserviert den Platz (ABM_QUERYPOS)
            SHAppBarMessage(ABM_QUERYPOS, ref abd);

            // Position final setzen (ABM_SETPOS)
            SHAppBarMessage(ABM_SETPOS, ref abd);

            // Fenstergröße anpassen
            this.SetBounds(abd.rc.Left, abd.rc.Top,
                           abd.rc.Right - abd.rc.Left,
                           abd.rc.Bottom - abd.rc.Top);

            // Versuch, das Fenster immer TopMost zu halten
            ForceTopMost();
        }

        #endregion

        #region Kontextmenü

        private void ErstelleKontextMenü()
        {
            // Haupt-Kontextmenü
            contextMenu = new ContextMenuStrip();

            // Untermenü für Docking
            var leftItem = new ToolStripMenuItem("Links andocken", null, (s, e) => ÄndereDockPosition(DockPosition.Left));
            var rightItem = new ToolStripMenuItem("Rechts andocken", null, (s, e) => ÄndereDockPosition(DockPosition.Right));
            var topItem = new ToolStripMenuItem("Oben andocken", null, (s, e) => ÄndereDockPosition(DockPosition.Top));
            var bottomItem = new ToolStripMenuItem("Unten andocken", null, (s, e) => ÄndereDockPosition(DockPosition.Bottom));

            var dockMenu = new ToolStripMenuItem("Andockposition");
            dockMenu.DropDownItems.Add(leftItem);
            dockMenu.DropDownItems.Add(rightItem);
            dockMenu.DropDownItems.Add(topItem);
            dockMenu.DropDownItems.Add(bottomItem);

            // Untermenü für Monitor-Auswahl
            monitorMenu = new ToolStripMenuItem("Monitor auswählen");
            AktualisiereMonitorListe(); // einmal initial befüllen

            // Neuer Menüpunkt "Config öffnen"
            var configOpenItem = new ToolStripMenuItem("Config öffnen", null, (s, e) =>
            {
                // Pfad zur settings.json zusammensetzen
                string settingsFile = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "LinkbarNET",
                    "settings.json"
                );
                // Notepad mit der settings.json starten
                Process.Start("notepad.exe", settingsFile);
            });

            // Menüeintrag: Beenden
            ToolStripMenuItem exitItem = new ToolStripMenuItem("Beenden", null, (s, e) => Beenden());

            // Menüstruktur zusammenbauen
            contextMenu.Items.Add(dockMenu);
            contextMenu.Items.Add(monitorMenu);
            contextMenu.Items.Add(configOpenItem);
            contextMenu.Items.Add(new ToolStripSeparator());
            contextMenu.Items.Add(exitItem);

            // Monitorliste bei jedem Öffnen aktualisieren
            contextMenu.Opening += (sender, e) =>
            {
                AktualisiereMonitorListe();
            };
        }


        private void AktualisiereMonitorListe()
        {
            monitorMenu.DropDownItems.Clear();

            foreach (var screen in Screen.AllScreens)
            {
                var monitorItem = new ToolStripMenuItem(screen.DeviceName, null, (s, e) =>
                {
                    aktuellerMonitor = screen;
                    SetzeAppBarPosition();
                    LadeVerknuepfungen(false);

                    appSettings.DockPosition = this.dockPosition;
                    appSettings.MonitorDeviceName = this.aktuellerMonitor.DeviceName;
                    appSettings.ShortcutPath = this.shortcutPath;
                    SettingsManager.SaveSettings(appSettings);

                })
                {
                    Checked = (screen == aktuellerMonitor)
                };
                monitorMenu.DropDownItems.Add(monitorItem);
            }
        }

        #endregion

        #region Docking-Positionswechsel

        private void ÄndereDockPosition(DockPosition newPosition)
        {
            this.dockPosition = newPosition;

            // FlowLayoutPanel anpassen
            switch (newPosition)
            {
                case DockPosition.Left:
                case DockPosition.Right:
                    BarSize = 45; // schmale vertikale Leiste
                    panel.FlowDirection = FlowDirection.TopDown;
                    panel.WrapContents = false;
                    break;

                case DockPosition.Top:
                case DockPosition.Bottom:
                    BarSize = 40; // höhere Leiste für Icons
                    panel.FlowDirection = FlowDirection.LeftToRight;
                    panel.WrapContents = false;
                    break;
            }

            // Neu positionieren
            SetzeAppBarPosition();

            // Icons neu aufbauen
            LadeVerknuepfungen(false);
        }

        #endregion

        #region Verknüpfungen

        private void LadeVerknuepfungen(bool firststart)
        {
            panel.Controls.Clear();

            if (!Directory.Exists(shortcutPath))
                Directory.CreateDirectory(shortcutPath);

            string[] linkFiles = Directory.GetFiles(shortcutPath, "*.lnk");
            foreach (string lnk in linkFiles)
            {
                string targetPath = GetShortcutTarget(lnk);
                if (!string.IsNullOrEmpty(targetPath))
                {
                    ErstelleIcon(lnk, targetPath);
                }
            }

            // Icons nur zentrieren, wenn nicht der erste Start:
            if (!firststart)
                ZentriereIcons();
        }

        private void ZentriereIcons()
        {
            // Falls die Dock-Position links oder rechts ist,
            // das Panel NICHT zentrieren, sondern Standard-Padding wiederherstellen.
            if (dockPosition == DockPosition.Left || dockPosition == DockPosition.Right)
            {
                panel.Padding = new Padding(5);
                return;
            }

            // Bei TOP/BOTTOM: Icons horizontal zentrieren
            int panelBreite = panel.ClientSize.Width;
            int totalWidth = 0;
            foreach (Control c in panel.Controls)
            {
                totalWidth += c.Width + c.Margin.Horizontal;
            }
            int freieBreite = panelBreite - totalWidth;
            int padding = freieBreite > 0 ? freieBreite / 2 : 0;
            panel.Padding = new Padding(padding, panel.Padding.Top, padding, panel.Padding.Bottom);
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern int ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, int nIcons);

        private void ErstelleIcon(string shortcutFile, string targetPath)
        {
            if (!System.IO.File.Exists(targetPath))
            {
                Debug.WriteLine($"[WARN] Verknüpfungsziel existiert nicht: {targetPath}");
                // Optional: return; falls das Icon nur angezeigt werden soll, wenn das Ziel existiert.
            }

            // PictureBox erzeugen
            var pic = new PictureBox
            {
                Width = ICON_SIZE_NORMAL_WIDTH,
                Height = ICON_SIZE_NORMAL_HEIGHT,
                SizeMode = PictureBoxSizeMode.StretchImage,
                Margin = new Padding(4),
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand
            };

            // ToolTip-Text setzen (z. B. Dateiname ohne Erweiterung)
            string iconText = Path.GetFileNameWithoutExtension(shortcutFile);
            iconToolTip.SetToolTip(pic, iconText);

            // Konsistenter Schlüssel (absoluter Pfad der Shortcut-Datei)
            string key = Path.GetFullPath(shortcutFile);

            // Icon laden: zuerst prüfen, ob in den persistierten Einstellungen ein benutzerdefiniertes Icon hinterlegt ist
            try
            {
                if (appSettings.CustomIconMapping != null &&
                    appSettings.CustomIconMapping.TryGetValue(key, out string customIconPath) &&
                    System.IO.File.Exists(customIconPath))
                {
                    try
                    {
                        // Versuche das benutzerdefinierte Icon zu laden
                        pic.Image = new Icon(customIconPath).ToBitmap();
                    }
                    catch (ArgumentException)
                    {
                        // Falls das Icon ungültig ist, entferne den Eintrag und nutze den Fallback
                        appSettings.CustomIconMapping.Remove(key);
                        SettingsManager.SaveSettings(appSettings);
                        LoadFallbackIcon(shortcutFile, targetPath, pic);
                    }
                }
                else
                {
                    // Falls kein benutzerdefiniertes Icon definiert ist,
                    // wird das in der Shortcut hinterlegte Icon (IconLocation) oder das Standard-Icon verwendet.
                    var shell = new WshShell();
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutFile);
                    string iconLocation = shortcut.IconLocation;

                    if (!string.IsNullOrEmpty(iconLocation))
                    {
                        string iconFile = iconLocation;
                        int iconIndex = 0;
                        // Prüfen, ob ein Index (durch Komma getrennt) vorhanden ist
                        if (iconLocation.Contains(","))
                        {
                            var parts = iconLocation.Split(',');
                            iconFile = parts[0].Trim();
                            int.TryParse(parts[1].Trim(), out iconIndex);
                        }

                        if (System.IO.File.Exists(iconFile))
                        {
                            IntPtr[] largeIconHandles = new IntPtr[1];
                            IntPtr[] smallIconHandles = new IntPtr[1];
                            int extractedCount = ExtractIconEx(iconFile, iconIndex, largeIconHandles, smallIconHandles, 1);
                            if (extractedCount > 0 && largeIconHandles[0] != IntPtr.Zero)
                            {
                                using (Icon extractedIcon = Icon.FromHandle(largeIconHandles[0]))
                                {
                                    pic.Image = extractedIcon.ToBitmap();
                                }
                            }
                            else
                            {
                                LoadFallbackIcon(shortcutFile, targetPath, pic);
                            }
                        }
                        else
                        {
                            LoadFallbackIcon(shortcutFile, targetPath, pic);
                        }
                    }
                    else
                    {
                        LoadFallbackIcon(shortcutFile, targetPath, pic);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERR] Icon laden fehlgeschlagen für {targetPath}: {ex.Message}");
                LoadFallbackIcon(shortcutFile, targetPath, pic);
            }

            // Kontextmenü für das Icon erstellen
            var picContextMenu = new ContextMenuStrip();
            var changeIconItem = new ToolStripMenuItem("Icon ändern");
            changeIconItem.Click += (s, e) =>
            {
                using (var ofd = new OpenFileDialog())
                {
                    // Erlaube neben ICO auch DLL und EXE Dateien, falls gewünscht
                    ofd.Filter = "Icon Dateien (*.ico;*.dll;*.exe)|*.ico;*.dll;*.exe";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        string selectedIcon = ofd.FileName;
                        try
                        {
                            Bitmap newBitmap = null;
                            string ext = Path.GetExtension(selectedIcon).ToLowerInvariant();
                            if (ext == ".ico")
                            {
                                // Direkter Zugriff auf ICO-Dateien
                                newBitmap = new Icon(selectedIcon).ToBitmap();
                            }
                            else if (ext == ".dll" || ext == ".exe")
                            {
                                // Verwende ExtractIconEx, um das erste Icon aus der Datei zu extrahieren
                                IntPtr[] largeIconHandles = new IntPtr[1];
                                IntPtr[] smallIconHandles = new IntPtr[1];
                                int extractedCount = ExtractIconEx(selectedIcon, 0, largeIconHandles, smallIconHandles, 1);
                                if (extractedCount > 0 && largeIconHandles[0] != IntPtr.Zero)
                                {
                                    using (Icon extractedIcon = Icon.FromHandle(largeIconHandles[0]))
                                    {
                                        newBitmap = extractedIcon.ToBitmap();
                                    }
                                }
                            }
                            else
                            {
                                // Standard: versuche, es als ICO zu laden
                                newBitmap = new Icon(selectedIcon).ToBitmap();
                            }

                            if (newBitmap != null)
                            {
                                // Speichere nur in der settings.json
                                appSettings.CustomIconMapping[key] = selectedIcon;
                                SettingsManager.SaveSettings(appSettings);

                                // Aktualisiere das angezeigte Icon
                                pic.Image = newBitmap;
                            }
                            else
                            {
                                throw new Exception("Konnte kein gültiges Icon aus der ausgewählten Datei laden.");
                            }
                        }
                        catch (Exception exIcon)
                        {
                            MessageBox.Show("Fehler beim Ändern des Icons: " + exIcon.Message,
                                            "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            };


            picContextMenu.Items.Add(changeIconItem);


            pic.MouseUp += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    // Ermittelt die absoluten Bildschirmkoordinaten des Klicks
                    Point screenPt = pic.PointToScreen(e.Location);
                    Screen currentScreen = Screen.FromPoint(screenPt);
                    Rectangle workingArea = currentScreen.WorkingArea;

                    // Berechne die bevorzugte Größe des Kontextmenüs
                    Size menuSize = picContextMenu.GetPreferredSize(Size.Empty);
                    int x = screenPt.X;
                    int y = screenPt.Y;

                    // Korrigiere die Position, falls das Menü sonst außerhalb des Monitors liegen würde
                    if (x + menuSize.Width > workingArea.Right)
                        x = workingArea.Right - menuSize.Width;
                    if (y + menuSize.Height > workingArea.Bottom)
                        y = workingArea.Bottom - menuSize.Height;

                    // Zeige das Kontextmenü an der berechneten Position (im Bildschirmkoordinaten-Modus)
                    picContextMenu.Show(new Point(x, y));
                }
            };

            // Falls bereits ein benutzerdefiniertes Icon gesetzt ist, füge "Icon löschen" hinzu.
            if (appSettings.CustomIconMapping.ContainsKey(key))
            {
                var deleteIconItem = new ToolStripMenuItem("Icon löschen");
                deleteIconItem.Click += (s, e) =>
                {
                    try
                    {
                        // Entferne den Eintrag aus der persistierten Instanz
                        appSettings.CustomIconMapping.Remove(key);
                        SettingsManager.SaveSettings(appSettings);

                        // Lade das Icon neu: Versuche, das in der Shortcut hinterlegte Icon zu verwenden,
                        // ansonsten das Standard-Icon
                        var shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutFile);
                        string iconLocation = shortcut.IconLocation;
                        if (!string.IsNullOrEmpty(iconLocation))
                        {
                            string iconFile = iconLocation;
                            int iconIndex = 0;
                            if (iconLocation.Contains(","))
                            {
                                var parts = iconLocation.Split(',');
                                iconFile = parts[0].Trim();
                                int.TryParse(parts[1].Trim(), out iconIndex);
                            }
                            if (System.IO.File.Exists(iconFile))
                            {
                                IntPtr[] largeIconHandles = new IntPtr[1];
                                IntPtr[] smallIconHandles = new IntPtr[1];
                                int extractedCount = ExtractIconEx(iconFile, iconIndex, largeIconHandles, smallIconHandles, 1);
                                if (extractedCount > 0 && largeIconHandles[0] != IntPtr.Zero)
                                {
                                    using (Icon extractedIcon = Icon.FromHandle(largeIconHandles[0]))
                                    {
                                        pic.Image = extractedIcon.ToBitmap();
                                    }
                                }
                                else
                                {
                                    LoadFallbackIcon(shortcutFile, targetPath, pic);
                                }
                            }
                            else
                            {
                                LoadFallbackIcon(shortcutFile, targetPath, pic);
                            }
                        }
                        else
                        {
                            LoadFallbackIcon(shortcutFile, targetPath, pic);
                        }
                    }
                    catch (Exception exDelete)
                    {
                        MessageBox.Show("Fehler beim Löschen des benutzerdefinierten Icons: " + exDelete.Message,
                                        "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };
                picContextMenu.Items.Add(deleteIconItem);
            }

            pic.ContextMenuStrip = picContextMenu;

            // Beim Hovern vergrößern
            pic.MouseEnter += (s, e) =>
            {
                pic.Width = ICON_SIZE_HOVER_WIDTH;
                pic.Height = ICON_SIZE_HOVER_HEIGHT;
            };

            // Beim Verlassen wieder verkleinern
            pic.MouseLeave += (s, e) =>
            {
                pic.Width = ICON_SIZE_NORMAL_WIDTH;
                pic.Height = ICON_SIZE_NORMAL_HEIGHT;
            };

            // Beim linken Mausklick: Icon vergrößern und Ziel starten
            pic.MouseClick += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    pic.Width = ICON_SIZE_CLICK_WIDTH;
                    pic.Height = ICON_SIZE_CLICK_HEIGHT;

                    try
                    {
                        Process.Start(targetPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Fehler beim Starten: {ex.Message}",
                                        "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };

            pic.ContextMenuStrip.Opening += (sender, e) =>
            {
                // Überprüfe, ob für diesen Shortcut (key) ein benutzerdefiniertes Icon gesetzt ist
                bool hasCustomIcon = appSettings.CustomIconMapping.ContainsKey(key);

                // Prüfe, ob "Icon löschen" bereits vorhanden ist
                var deleteItem = pic.ContextMenuStrip.Items
                    .OfType<ToolStripMenuItem>()
                    .FirstOrDefault(item => item.Text == "Icon löschen");

                if (hasCustomIcon && deleteItem == null)
                {
                    // Füge "Icon löschen" hinzu
                    deleteItem = new ToolStripMenuItem("Icon löschen");
                    deleteItem.Click += (s, e2) =>
                    {
                        try
                        {
                            appSettings.CustomIconMapping.Remove(key);
                            SettingsManager.SaveSettings(appSettings);
                            // Nach dem Löschen, das Icon neu laden (Fallback: Shortcut-Icon)
                            LoadFallbackIcon(shortcutFile, targetPath, pic);

                            // Entferne den Menüeintrag sofort
                            pic.ContextMenuStrip.Items.Remove(deleteItem);
                        }
                        catch (Exception exDelete)
                        {
                            MessageBox.Show("Fehler beim Löschen des benutzerdefinierten Icons: " + exDelete.Message,
                                            "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    };
                    pic.ContextMenuStrip.Items.Add(deleteItem);
                }
                else if (!hasCustomIcon && deleteItem != null)
                {
                    // Entferne den Menüeintrag, falls kein benutzerdefiniertes Icon mehr vorhanden ist
                    pic.ContextMenuStrip.Items.Remove(deleteItem);
                }
            };


            panel.Controls.Add(pic);
        }

        /// <summary>
        /// Hilfsmethode zum Laden eines Fallback-Icons (Standard-Icon) aus dem Zielpfad.
        /// </summary>
        private void LoadFallbackIcon(string shortcutFile, string targetPath, PictureBox pic)
        {
            try
            {
                var ico = Icon.ExtractAssociatedIcon(targetPath);
                if (ico != null)
                {
                    pic.Image = ico.ToBitmap();
                }
            }
            catch (Exception exFallback)
            {
                Debug.WriteLine($"[ERR] Fallback-Icon konnte nicht geladen werden: {exFallback.Message}");
            }
        }



        private string GetShortcutTarget(string lnkFile)
        {
            try
            {
                var shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(lnkFile);
                return shortcut.TargetPath;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region Beenden & WndProc

        private void Beenden()
        {
            EntferneAppBar();

            if (reloadTimer != null)
            {
                reloadTimer.Stop();
                reloadTimer.Dispose();
            }

            Application.Exit();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == _uCallBackMsg)
            {
                // m.WParam == 1 signalisiert, dass Windows eine Neupositionierung verlangt
                if (m.WParam.ToInt32() == 1)
                {
                    SetzeAppBarPosition();
                }
            }
            base.WndProc(ref m);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            appSettings.DockPosition = this.dockPosition;
            appSettings.MonitorDeviceName = this.aktuellerMonitor.DeviceName;
            appSettings.ShortcutPath = this.shortcutPath;

            // Speichere die Einstellungen, ohne das CustomIconMapping zu verlieren
            SettingsManager.SaveSettings(appSettings);

            EntferneAppBar();

            if (reloadTimer != null)
            {
                reloadTimer.Stop();
                reloadTimer.Dispose();
            }

            base.OnFormClosing(e);
        }

        #endregion
    }

    // --- Klassen zur Speicherung der Anwendungseinstellungen ---

    public class AppSettings
    {
        public DockPosition DockPosition { get; set; }
        public string MonitorDeviceName { get; set; }
        public string ShortcutPath { get; set; }

        // Neues Mapping: Schlüssel = Pfad der .lnk-Datei, Wert = Pfad des benutzerdefinierten Icons
        public Dictionary<string, string> CustomIconMapping { get; set; } = new Dictionary<string, string>();
    }

    public static class SettingsManager
    {
        private static readonly string settingsFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LinkbarNET");
        private static readonly string settingsFile = Path.Combine(settingsFolder, "settings.json");

        public static AppSettings LoadSettings()
        {
            try
            {
                if (System.IO.File.Exists(settingsFile))
                {
                    string json = System.IO.File.ReadAllText(settingsFile);
                    var settings = JsonSerializer.Deserialize<AppSettings>(json);
                    if (settings != null && settings.CustomIconMapping == null)
                    {
                        settings.CustomIconMapping = new Dictionary<string, string>();
                    }
                    return settings;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error loading settings: " + ex.Message);
            }
            // Rückfall auf Standardwerte
            return new AppSettings { CustomIconMapping = new Dictionary<string, string>() };
        }


        public static void SaveSettings(AppSettings settings)
        {
            try
            {
                Directory.CreateDirectory(settingsFolder);
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(settings, options);
                System.IO.File.WriteAllText(settingsFile, json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error saving settings: " + ex.Message);
            }
        }
    }

}
