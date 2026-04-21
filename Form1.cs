using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using AForge.Video;
using AForge.Video.DirectShow;

namespace tCamView
{
    public partial class Form1 : Form
    {
        // For ellipse/rounded window shapes
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,
            int nTopRect,
            int nRightRect,
            int nBottomRect,
            int nWidthEllipse,
            int nHeightEllipse
        );

        // For WM_SIZING aspect-ratio enforcement
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private const int WM_SIZING = 0x0214;
        private const int WMSZ_LEFT = 1, WMSZ_RIGHT = 2, WMSZ_TOP = 3,
                          WMSZ_TOPLEFT = 4, WMSZ_TOPRIGHT = 5,
                          WMSZ_BOTTOM = 6, WMSZ_BOTTOMLEFT = 7, WMSZ_BOTTOMRIGHT = 8;

        private FilterInfoCollection webcam;
        private VideoCaptureDevice cam;

        List<string> videoCaptureDevicesList = new List<string>();
        List<string> videoCapabilitiesList = new List<string>();

        bool state_flip_vertical = false;
        bool state_flip_horizontal = false;
        int currCamID = -1;
        int currSizeID = -1;
        int cropSize = 0;

        MenuItem Menu_VideoCaptureDevices = new MenuItem("Video Capture Devices");
        MenuItem Menu_VideoCapabilities = new MenuItem("Video Resolutions");

        bool firstimage_captured = false;
        int videoCaptureDevicesListCount = 0;
        int videoCapabilitiesListCount = 0;
        int currMenuItem0VideoCaptureDevices = -1;
        int currMenuItem1VideoCapabilities = -1;
        int currMenuItem3WindowStyles = 0;
        const int currMenuItem3WindowStylesCount = 5;

        bool stretchKeepAspectRatio = true; // Alt.Stretch / UniformToFill

        // --- NEW: Lock Aspect Ratio ---
        bool lockAspectRatio = false;
        double lockedAspect = 0.0; // height / width

        // --- Settings file path & version ---
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "tCamView", "settings.ini");
        private const int SettingsVersion = 2;

        public Form1()
        {
            InitializeComponent();
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.ControlBox = true;
            this.Icon = Properties.Resources.webcam;
            this.Opacity = 1.0;
            this.ShowInTaskbar = true;
            this.TopMost = false; // keep false until Form1_Shown fires
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            stretchKeepAspectRatio = true;
            this.Text = "tCamView (alt.stretch)";
            this.Shown += new EventHandler(Form1_Shown);
        }

        // Changing FormBorderStyle causes Windows to recreate the window handle (HWND),
        // which drops the taskbar entry. Re-assert ShowInTaskbar every time this happens.
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            this.ShowInTaskbar = true;
        }

        // -------------------------------------------------------
        //  SETTINGS  SAVE / LOAD
        // -------------------------------------------------------
        private void SaveSettings()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath));
                var sb = new StringBuilder();
                sb.AppendLine("SettingsVersion=" + SettingsVersion);
                sb.AppendLine("CamID=" + currCamID);
                sb.AppendLine("SizeID=" + currSizeID);
                sb.AppendLine("PictureSizeMode=" + (int)pictureBox1.SizeMode);
                sb.AppendLine("StretchKeepAspectRatio=" + stretchKeepAspectRatio);
                sb.AppendLine("WindowStyle=" + currMenuItem3WindowStyles);
                sb.AppendLine("FlipH=" + state_flip_horizontal);
                sb.AppendLine("FlipV=" + state_flip_vertical);
                sb.AppendLine("Opacity=" + this.Opacity.ToString(System.Globalization.CultureInfo.InvariantCulture));
                sb.AppendLine("AlwaysOnTop=" + this.TopMost);
                sb.AppendLine("LockAspectRatio=" + lockAspectRatio);
                sb.AppendLine("CropSize=" + cropSize);
                // Save window size and position for all styles except maximized full screen
                if (this.WindowState != FormWindowState.Maximized)
                {
                    sb.AppendLine("WinX=" + this.Left);
                    sb.AppendLine("WinY=" + this.Top);
                    sb.AppendLine("WinW=" + this.Width);
                    sb.AppendLine("WinH=" + this.Height);
                }
                File.WriteAllText(SettingsPath, sb.ToString());
            }
            catch { /* silently ignore save errors */ }
        }

        private Dictionary<string, string> LoadSettings()
        {
            var dict = new Dictionary<string, string>();
            try
            {
                if (!File.Exists(SettingsPath)) return dict;
                foreach (var line in File.ReadAllLines(SettingsPath))
                {
                    var idx = line.IndexOf('=');
                    if (idx > 0)
                        dict[line.Substring(0, idx).Trim()] = line.Substring(idx + 1).Trim();
                }
                // If the settings file is from an older version, ignore it entirely
                // and let the app start fresh with defaults
                int savedVersion = 0;
                if (!dict.ContainsKey("SettingsVersion") ||
                    !int.TryParse(dict["SettingsVersion"], out savedVersion) ||
                    savedVersion < SettingsVersion)
                {
                    return new Dictionary<string, string>();
                }
            }
            catch { }
            return dict;
        }

        private T GetSetting<T>(Dictionary<string, string> d, string key, T defaultVal)
        {
            if (!d.ContainsKey(key)) return defaultVal;
            try { return (T)Convert.ChangeType(d[key], typeof(T), System.Globalization.CultureInfo.InvariantCulture); }
            catch { return defaultVal; }
        }

        // -------------------------------------------------------
        //  FORM LOAD
        // -------------------------------------------------------
        private void Form1_Load(object sender, EventArgs e)
        {
            var settings = LoadSettings();

            try
            {
                webcam = new FilterInfoCollection(FilterCategory.VideoInputDevice);
                foreach (FilterInfo VideoCaptureDevice in webcam)
                    videoCaptureDevicesList.Add(VideoCaptureDevice.Name);
            }
            catch (Exception error) { MessageBox.Show(error.Message); }

            Focus();

            for (var i = 0; i < videoCaptureDevicesList.Count; i++)
            {
                MenuItem mi = new MenuItem(videoCaptureDevicesList[i]);
                mi.Click += new EventHandler(MenuItem_VideoCaptureDevices_Click);
                mi.Tag = new object[] { i };
                Menu_VideoCaptureDevices.MenuItems.Add(mi);
            }
            videoCaptureDevicesListCount = videoCaptureDevicesList.Count;

            // Restore saved camera / size, or pick defaults
            int savedCam = GetSetting(settings, "CamID", -1);
            int savedSize = GetSetting(settings, "SizeID", -1);

            bool opened = false;
            if (savedCam >= 0 && savedCam < videoCaptureDevicesListCount)
            {
                if (openVideoCaptureDevice(savedCam, savedSize) == 1)
                {
                    currMenuItem0VideoCaptureDevices = savedCam;
                    currMenuItem1VideoCapabilities = currSizeID;
                    opened = true;
                }
            }
            if (!opened)
            {
                for (int i = 0; i < videoCaptureDevicesList.Count; i++)
                {
                    if (openVideoCaptureDevice(i, -1) == 1)
                    {
                        currMenuItem0VideoCaptureDevices = i;
                        currMenuItem1VideoCapabilities = 0;
                        break;
                    }
                }
            }

            if (videoCaptureDevicesListCount == 0)
                CreateContextMenu();

            // Restore picture mode
            int savedMode = GetSetting(settings, "PictureSizeMode", (int)PictureBoxSizeMode.StretchImage);
            bool savedStretch = GetSetting(settings, "StretchKeepAspectRatio", true);
            stretchKeepAspectRatio = savedStretch;
            pictureBox1.SizeMode = (PictureBoxSizeMode)savedMode;
            switch (pictureBox1.SizeMode)
            {
                case PictureBoxSizeMode.Zoom: this.Text = "tCamView (zoom)"; break;
                case PictureBoxSizeMode.CenterImage: this.Text = "tCamView (center)"; break;
                case PictureBoxSizeMode.StretchImage:
                    this.Text = stretchKeepAspectRatio ? "tCamView (alt.stretch)" : "tCamView (stretch)"; break;
            }

            // Restore flips
            state_flip_horizontal = GetSetting(settings, "FlipH", false);
            state_flip_vertical = GetSetting(settings, "FlipV", false);

            // Restore opacity
            this.Opacity = GetSetting(settings, "Opacity", 1.0);

            // Defer TopMost until Form1_Shown — setting it during Load causes Windows
            // to drop the taskbar entry before the window is fully registered
            _pendingTopMost = GetSetting(settings, "AlwaysOnTop", true);

            // Restore lock aspect ratio
            lockAspectRatio = GetSetting(settings, "LockAspectRatio", false);

            // Restore crop size
            cropSize = GetSetting(settings, "CropSize", 0);

            // Apply window style first so border mode is correct before sizing
            int savedWinStyle = GetSetting(settings, "WindowStyle", 0);
            ApplyWindowStyle(savedWinStyle);

            // Restore size/position AFTER style is set so the region is built
            // with the correct dimensions (critical for borderless modes)
            int wx = GetSetting(settings, "WinX", -1);
            int wy = GetSetting(settings, "WinY", -1);
            int ww = GetSetting(settings, "WinW", -1);
            int wh = GetSetting(settings, "WinH", -1);
            if (wx >= 0 && wy >= 0 && ww > 50 && wh > 50 && savedWinStyle != 4 /* not fullscreen */)
            {
                this.StartPosition = FormStartPosition.Manual;
                this.Left = wx; this.Top = wy;
                this.Width = ww; this.Height = wh;
                // For borderless styles, rebuild the region now that size is final
                RefreshRegion();
            }

            // Update aspect ratio lock state
            if (lockAspectRatio)
                lockedAspect = (double)this.Height / this.Width;

            UpdateMenuItemsChecked();
        }

        // TopMost is applied here rather than in Form1_Load so the window is fully
        // registered with the taskbar first — setting it earlier causes Windows to
        // silently remove the taskbar button.
        private bool _pendingTopMost = true;
        private void Form1_Shown(object sender, EventArgs e)
        {
            this.TopMost = false;
            this.TopMost = _pendingTopMost;
        }

        private void ApplyWindowStyle(int styleIndex)
        {
            currMenuItem3WindowStyles = styleIndex;
            switch (styleIndex)
            {
                case 0: // Normal
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                    if (this.WindowState != FormWindowState.Normal)
                        WindowState = FormWindowState.Normal;
                    break;
                case 1: // Borderless Ellipse
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                    Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, Width, Height));
                    break;
                case 2: // Borderless Rectangle
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                    Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 0, 0));
                    break;
                case 3: // Borderless Rounded Rectangle
                    int m = Math.Min(Width, Height) / 4;
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                    Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, m, m));
                    break;
                case 4: // Full Screen
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                    WindowState = FormWindowState.Normal;
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                    WindowState = FormWindowState.Maximized;
                    break;
            }
            // Windows can silently clear ShowInTaskbar when FormBorderStyle changes,
            // so explicitly re-assert it after every style switch.
            this.ShowInTaskbar = true;
        }

        // -------------------------------------------------------
        //  WM_SIZING  – enforce locked aspect ratio while resizing
        // -------------------------------------------------------
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_SIZING && lockAspectRatio && lockedAspect > 0)
            {
                var rc = (RECT)Marshal.PtrToStructure(m.LParam, typeof(RECT));
                int w = rc.Right - rc.Left;
                int h = rc.Bottom - rc.Top;

                switch (m.WParam.ToInt32())
                {
                    case WMSZ_LEFT:
                    case WMSZ_RIGHT:
                    case WMSZ_BOTTOMLEFT:
                    case WMSZ_BOTTOMRIGHT:
                        rc.Bottom = rc.Top + (int)(w * lockedAspect);
                        break;
                    case WMSZ_TOP:
                    case WMSZ_BOTTOM:
                    case WMSZ_TOPLEFT:
                    case WMSZ_TOPRIGHT:
                        rc.Right = rc.Left + (int)(h / lockedAspect);
                        break;
                }

                Marshal.StructureToPtr(rc, m.LParam, false);
                m.Result = (IntPtr)1;
                return;
            }
            base.WndProc(ref m);
        }

        // -------------------------------------------------------
        //  CONTEXT MENU
        // -------------------------------------------------------
        private void CreateContextMenu()
        {
            ContextMenu cm = new ContextMenu();

            Menu_VideoCapabilities.MenuItems.Clear();
            for (var i = 0; i < videoCapabilitiesList.Count; i++)
            {
                MenuItem mi = new MenuItem(videoCapabilitiesList[i]);
                mi.Click += new EventHandler(MenuItem_VideoCapabilties_Click);
                mi.Tag = new object[] { i };
                Menu_VideoCapabilities.MenuItems.Add(mi);
            }

            MenuItem pictureMode = new MenuItem("PictureSize Mode");
            pictureMode.MenuItems.Add(new MenuItem("Zoom [Z]", MenuItem_Zoom_Click));
            pictureMode.MenuItems.Add(new MenuItem("Stretch [X]", MenuItem_Stretch_Click));
            pictureMode.MenuItems.Add(new MenuItem("Center [C]", MenuItem_Center_Click));
            pictureMode.MenuItems.Add(new MenuItem("Alt.Stretch [A]", MenuItem_AltStretch_Click));

            MenuItem borderOptions = new MenuItem("Window Styles");
            borderOptions.MenuItems.Add(new MenuItem("Normal Border (Resizable) [N or Esc]", MenuItem_Normal_Border_Click));
            borderOptions.MenuItems.Add(new MenuItem("Borderless Ellipse (FixedSize) [E]", MenuItem_Borderless_Ellipse_Click));
            borderOptions.MenuItems.Add(new MenuItem("Borderless Rectangle (FixedSize) [R]", MenuItem_Borderless_Rectangle_Click));
            borderOptions.MenuItems.Add(new MenuItem("Borderless Rounded Rectangle (FixedSize) [W]", MenuItem_Borderless_RoundedRectangle_Click));
            borderOptions.MenuItems.Add(new MenuItem("Full Screen [F]", MenuItem_FullScreen_Click));

            MenuItem imageFlipping = new MenuItem("Image Flipping");
            imageFlipping.MenuItems.Add(new MenuItem("Horizontal Flipping [H]", MenuItem_Horizontal_Flipping_Click));
            imageFlipping.MenuItems.Add(new MenuItem("Vertical Flipping [V]", MenuItem_Vertical_Flipping_Click));

            MenuItem opacityControl = new MenuItem("Opacity Control");
            opacityControl.MenuItems.Add(new MenuItem("Opacity Increase [Up Arrow]", MenuItem_Opacity_Increase_Click));
            opacityControl.MenuItems.Add(new MenuItem("Opacity Decrease [Down Arrow]", MenuItem_Opacity_Decrease_Click));
            opacityControl.MenuItems.Add(new MenuItem("Opacity 100% (max) [Right Arrow]", MenuItem_Opacity_100_Click));
            opacityControl.MenuItems.Add(new MenuItem("Opacity 80%", MenuItem_Opacity_80_Click));
            opacityControl.MenuItems.Add(new MenuItem("Opacity 60%", MenuItem_Opacity_60_Click));
            opacityControl.MenuItems.Add(new MenuItem("Opacity 40%", MenuItem_Opacity_40_Click));
            opacityControl.MenuItems.Add(new MenuItem("Opacity 20% (min) [Left Arrow]", MenuItem_Opacity_20_Click));

            MenuItem addFeatures = new MenuItem("Additional Features");
            addFeatures.MenuItems.Add(new MenuItem("GetImage From Clipboard [G or ^V]", MenuItem_GetImageFromClipboard_Click));
            addFeatures.MenuItems.Add(new MenuItem("SetImage To Clipboard [I or ^C]", MenuItem_SetImageToClipboard_Click));
            addFeatures.MenuItems.Add(new MenuItem("SetImage To Clipboard After 5 Sec. [D]", MenuItem_SetImageToClipboardAfter5Sec_Click));
            addFeatures.MenuItems.Add(new MenuItem("CopyScreen To Clipboard [S]", MenuItem_CopyScreenToClipboard_Click));
            addFeatures.MenuItems.Add(new MenuItem("Resume CameraPreview (From ClipboardView) [Space]", MenuItem_ResumeCameraPreviewFromClipboard_Click));
            addFeatures.MenuItems.Add("-");
            addFeatures.MenuItems.Add(new MenuItem("CropImage(ZoomIn) [Page Up]", MenuItem_CropImageZoomIn_Click));
            addFeatures.MenuItems.Add(new MenuItem("CropImage(ZoomOut) [Page Down]", MenuItem_CropImageZoomOut_Click));
            addFeatures.MenuItems.Add("-");
            addFeatures.MenuItems.Add(new MenuItem("Increase the Window Size [P]", MenuItem_IncreaseWindowSize_Click));
            addFeatures.MenuItems.Add(new MenuItem("Decrease the Window Size [M]", MenuItem_DecreaseWindowSize_Click));

            MenuItem onTop = new MenuItem("Always on Top [T]", MenuItem_AlwaysOnTop_Click);

            // NEW: Lock Aspect Ratio menu item
            MenuItem lockAR = new MenuItem("Lock Aspect Ratio [K]", MenuItem_LockAspectRatio_Click);

            MenuItem minimizeApp = new MenuItem("Minimize [L]", MenuItem_Minimize_Click);
            MenuItem quitApp = new MenuItem("Quit [Q]", MenuItem_Quit_Click);
            MenuItem aboutApp = new MenuItem("About...", MenuItem_About_Click);

            cm.MenuItems.AddRange(new MenuItem[] {
                Menu_VideoCaptureDevices, Menu_VideoCapabilities,
                pictureMode, borderOptions, imageFlipping, opacityControl, addFeatures,
                onTop, lockAR, minimizeApp, quitApp, aboutApp
            });

            pictureBox1.ContextMenu = cm;
        }

        // Menu item indices (after adding Lock Aspect Ratio at index 8):
        // 0 = VideoCaptureDevices
        // 1 = VideoCapabilities
        // 2 = PictureSize Mode
        // 3 = Window Styles
        // 4 = Image Flipping
        // 5 = Opacity Control
        // 6 = Additional Features
        // 7 = Always on Top
        // 8 = Lock Aspect Ratio   ← NEW
        // 9 = Minimize
        // 10 = Quit
        // 11 = About

        private void UpdateMenuItemsChecked()
        {
            for (int i = 0; i < videoCaptureDevicesListCount; i++)
                pictureBox1.ContextMenu.MenuItems[0].MenuItems[i].Checked = false;
            if (currMenuItem0VideoCaptureDevices >= 0 && currMenuItem0VideoCaptureDevices < videoCaptureDevicesListCount)
                pictureBox1.ContextMenu.MenuItems[0].MenuItems[currMenuItem0VideoCaptureDevices].Checked = true;

            for (int i = 0; i < videoCapabilitiesListCount; i++)
                pictureBox1.ContextMenu.MenuItems[1].MenuItems[i].Checked = false;
            if (currMenuItem1VideoCapabilities >= 0 && currMenuItem1VideoCapabilities < videoCapabilitiesListCount)
                pictureBox1.ContextMenu.MenuItems[1].MenuItems[currMenuItem1VideoCapabilities].Checked = true;

            if (pictureBox1.SizeMode == PictureBoxSizeMode.Zoom)
            {
                pictureBox1.ContextMenu.MenuItems[2].MenuItems[0].Checked = true;
                pictureBox1.ContextMenu.MenuItems[2].MenuItems[1].Checked = false;
                pictureBox1.ContextMenu.MenuItems[2].MenuItems[2].Checked = false;
                pictureBox1.ContextMenu.MenuItems[2].MenuItems[3].Checked = false;
            }
            else if (pictureBox1.SizeMode == PictureBoxSizeMode.StretchImage)
            {
                pictureBox1.ContextMenu.MenuItems[2].MenuItems[0].Checked = false;
                pictureBox1.ContextMenu.MenuItems[2].MenuItems[2].Checked = false;
                if (stretchKeepAspectRatio == false)
                {
                    pictureBox1.ContextMenu.MenuItems[2].MenuItems[1].Checked = true;
                    pictureBox1.ContextMenu.MenuItems[2].MenuItems[3].Checked = false;
                }
                else
                {
                    pictureBox1.ContextMenu.MenuItems[2].MenuItems[1].Checked = false;
                    pictureBox1.ContextMenu.MenuItems[2].MenuItems[3].Checked = true;
                }
            }
            else if (pictureBox1.SizeMode == PictureBoxSizeMode.CenterImage)
            {
                pictureBox1.ContextMenu.MenuItems[2].MenuItems[0].Checked = false;
                pictureBox1.ContextMenu.MenuItems[2].MenuItems[1].Checked = false;
                pictureBox1.ContextMenu.MenuItems[2].MenuItems[2].Checked = true;
                pictureBox1.ContextMenu.MenuItems[2].MenuItems[3].Checked = false;
            }

            for (int i = 0; i < currMenuItem3WindowStylesCount; i++)
                pictureBox1.ContextMenu.MenuItems[3].MenuItems[i].Checked = false;
            pictureBox1.ContextMenu.MenuItems[3].MenuItems[currMenuItem3WindowStyles].Checked = true;

            pictureBox1.ContextMenu.MenuItems[4].MenuItems[0].Checked = state_flip_horizontal;
            pictureBox1.ContextMenu.MenuItems[4].MenuItems[1].Checked = state_flip_vertical;

            pictureBox1.ContextMenu.MenuItems[7].Checked = this.TopMost;
            pictureBox1.ContextMenu.MenuItems[8].Checked = lockAspectRatio; // Lock Aspect Ratio
        }

        // -------------------------------------------------------
        //  LOCK ASPECT RATIO
        // -------------------------------------------------------
        private void MenuItem_LockAspectRatio_Click(object sender, EventArgs e)
        {
            ToggleLockAspectRatio();
        }

        private void ToggleLockAspectRatio()
        {
            lockAspectRatio = !lockAspectRatio;
            if (lockAspectRatio)
                lockedAspect = (double)this.Height / this.Width;
            UpdateMenuItemsChecked();
        }

        // -------------------------------------------------------
        //  CAMERA
        // -------------------------------------------------------
        private int openVideoCaptureDevice(int deviceID, int sizeID)
        {
            List<int> numOfPixelsList = new List<int>();

            if (cam != null)
                closeVideoCaptureDevice();

            cropSize = 0;
            cam = new VideoCaptureDevice(webcam[deviceID].MonikerString);
            cam.NewFrame += new NewFrameEventHandler(cam_NewFrame);

            var videoCapabilities = cam.VideoCapabilities;
            videoCapabilitiesList.Clear();
            numOfPixelsList.Clear();

            foreach (var video in videoCapabilities)
            {
                videoCapabilitiesList.Add(video.FrameSize.Width + "x" + video.FrameSize.Height);
                numOfPixelsList.Add(video.FrameSize.Width * video.FrameSize.Height);
            }

            int indexMax = 0;
            if (numOfPixelsList.Count() > 0)
            {
                int max = numOfPixelsList.Max();
                indexMax = Array.IndexOf(numOfPixelsList.ToArray(), max);
            }

            currCamID = deviceID;
            currSizeID = (sizeID < 0 || sizeID >= videoCapabilities.Count()) ? indexMax : sizeID;

            if (videoCapabilities.Count() > 0)
                cam.VideoResolution = cam.VideoCapabilities[currSizeID];

            videoCapabilitiesListCount = videoCapabilities.Count();

            firstimage_captured = false;
            cam.Start();
            CreateContextMenu();

            for (int i = 0; i < 10; i++)
            {
                if (firstimage_captured == true) return 1;
                System.Threading.Thread.Sleep(100);
            }
            return 0;
        }

        private void closeVideoCaptureDevice()
        {
            if (cam != null)
            {
                cam.SignalToStop();
                for (int i = 0; i < 30; i++)
                {
                    if (!cam.IsRunning) break;
                    System.Threading.Thread.Sleep(100);
                }
                if (cam.IsRunning) cam.Stop();
                cam = null;
            }
        }

        private void cam_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            if (this.WindowState == FormWindowState.Minimized) return;
            firstimage_captured = true;

            cropSize = Math.Max(0, Math.Min(cropSize, Math.Min(eventArgs.Frame.Width / 2, eventArgs.Frame.Height / 2)));
            float ratio = (float)eventArgs.Frame.Height / eventArgs.Frame.Width;
            int vcropSize = (int)(cropSize * ratio);

            var image = (Bitmap)eventArgs.Frame.Clone(new System.Drawing.Rectangle(cropSize, vcropSize,
                eventArgs.Frame.Width - 2 * cropSize, eventArgs.Frame.Height - 2 * vcropSize), eventArgs.Frame.PixelFormat);

            if (pictureBox1.SizeMode == PictureBoxSizeMode.StretchImage && stretchKeepAspectRatio == true)
            {
                float ratioImage = (float)eventArgs.Frame.Width / eventArgs.Frame.Height;
                float ratioPictureBox = (float)pictureBox1.ClientSize.Width / pictureBox1.ClientSize.Height;
                Bitmap cropped;
                if (ratioImage >= ratioPictureBox)
                {
                    int newWidth = (int)(image.Height * ratioPictureBox);
                    int wcrop = (int)((image.Width - newWidth) / 2);
                    cropped = (Bitmap)image.Clone(new System.Drawing.Rectangle(wcrop, 0, image.Width - 2 * wcrop, image.Height), image.PixelFormat);
                }
                else
                {
                    int newHeight = (int)(image.Width / ratioPictureBox);
                    int hcrop = (int)((image.Height - newHeight) / 2);
                    cropped = (Bitmap)image.Clone(new System.Drawing.Rectangle(0, hcrop, image.Width, image.Height - 2 * hcrop), image.PixelFormat);
                }
                image.Dispose(); // dispose the intermediate clone
                image = cropped;
            }
            else if ((pictureBox1.SizeMode == PictureBoxSizeMode.CenterImage) && (cropSize != 0))
            {
                var resized = new Bitmap(image, new System.Drawing.Size(eventArgs.Frame.Width, eventArgs.Frame.Height));
                image.Dispose();
                image = resized;
            }

            if (state_flip_horizontal && state_flip_vertical)
                image.RotateFlip(RotateFlipType.RotateNoneFlipXY);
            else if (state_flip_horizontal)
                image.RotateFlip(RotateFlipType.RotateNoneFlipX);
            else if (state_flip_vertical)
                image.RotateFlip(RotateFlipType.RotateNoneFlipY);

            // Dispose the previous frame before replacing it — this is the main memory leak fix.
            // Bitmap holds unmanaged GDI memory that the GC won't release promptly on its own.
            var oldImage = pictureBox1.Image;
            pictureBox1.Image = image;
            oldImage?.Dispose();
        }

        // -------------------------------------------------------
        //  FORM CLOSING  – save settings before exit
        // -------------------------------------------------------
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
            closeVideoCaptureDevice();
        }

        // -------------------------------------------------------
        //  DELAY HELPER
        // -------------------------------------------------------
        private static DateTime Delay(int MS)
        {
            DateTime ThisMoment = DateTime.Now;
            TimeSpan duration = new TimeSpan(0, 0, 0, 0, MS);
            DateTime AfterWards = ThisMoment.Add(duration);
            while (AfterWards >= ThisMoment)
            {
                System.Windows.Forms.Application.DoEvents();
                ThisMoment = DateTime.Now;
            }
            return DateTime.Now;
        }

        // -------------------------------------------------------
        //  MENU HANDLERS
        // -------------------------------------------------------
        private void MenuItem_SetImageToClipboard_Click(object sender, EventArgs e)
            => Clipboard.SetImage(pictureBox1.Image);

        private void MenuItem_SetImageToClipboardAfter5Sec_Click(object sender, EventArgs e)
        {
            if (label1.Visible == true) return;
            label1.Visible = true;
            for (int s = 5; s >= 1; s--)
            {
                label1.Text = s.ToString();
                Delay(1000);
            }
            label1.Visible = false;
            Clipboard.SetImage(pictureBox1.Image);
        }

        private void MenuItem_CopyScreenToClipboard_Click(object sender, EventArgs e)
        {
            Delay(300);
            Image img = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            Graphics g = Graphics.FromImage(img);
            g.CopyFromScreen(PointToScreen(pictureBox1.Location), new System.Drawing.Point(0, 0),
                new System.Drawing.Size(pictureBox1.Width, pictureBox1.Height));
            Clipboard.SetImage(img);
            g.Dispose();
        }

        private void MenuItem_GetImageFromClipboard_Click(object sender, EventArgs e)
        {
            using (var bmp = GetImageFromClipboard())
            {
                if (bmp != null)
                {
                    closeVideoCaptureDevice();
                    pictureBox1.Image = new Bitmap(bmp.Width, bmp.Height);
                    using (Graphics gr = Graphics.FromImage(pictureBox1.Image))
                    {
                        gr.DrawImage(bmp, 0, 0);
                        gr.Dispose();
                    }
                    pictureBox1.Refresh();
                }
            }
        }

        private void MenuItem_About_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "tCamView 1.4.1\n" +
                "Copyright © 2020-2021, Sung Deuk Kim\n" +
                "All rights reserved.\n" +
                "--------------------------------\n" +
                "https://github.com/augamvio/tCamView\n" +
                "Published under the GNU GPLv3 license.\n" +
                "(For details, see license.txt)\n" +
                "--------------------------------\n" +
                "Credits:\nAForge.NET https://github.com/andrewkirillov/AForge.NET\n" +
                "--------------------------------\n" +
                "Added love from Claude AI and NoCodeEdu\n" +
                "https://github.com/NoCodeEdu/tCamView-fork/"

            );
        }

        private void MenuItem_VideoCaptureDevices_Click(object sender, EventArgs e)
        {
            int deviceID = (int)((object[])((MenuItem)sender).Tag)[0];
            openVideoCaptureDevice(deviceID, -1);
            currMenuItem0VideoCaptureDevices = deviceID;
            currMenuItem1VideoCapabilities = currSizeID;
            UpdateMenuItemsChecked();
        }

        private void MenuItem_VideoCapabilties_Click(object sender, EventArgs e)
        {
            int sizeID = (int)((object[])((MenuItem)sender).Tag)[0];
            openVideoCaptureDevice(currCamID, sizeID);
            currMenuItem0VideoCaptureDevices = currCamID;
            currMenuItem1VideoCapabilities = sizeID;
            UpdateMenuItemsChecked();
        }

        private void MenuItem_Minimize_Click(object sender, EventArgs e)
            => this.WindowState = FormWindowState.Minimized;

        private void MenuItem_Quit_Click(object sender, EventArgs e)
            => Application.Exit();

        private void MenuItem_AlwaysOnTop_Click(object sender, EventArgs e)
        {
            this.TopMost = !this.TopMost;
            UpdateMenuItemsChecked();
        }

        private void MenuItem_Opacity_20_Click(object sender, EventArgs e) => this.Opacity = 0.2;
        private void MenuItem_Opacity_40_Click(object sender, EventArgs e) => this.Opacity = 0.4;
        private void MenuItem_Opacity_60_Click(object sender, EventArgs e) => this.Opacity = 0.6;
        private void MenuItem_Opacity_80_Click(object sender, EventArgs e) => this.Opacity = 0.8;
        private void MenuItem_Opacity_100_Click(object sender, EventArgs e) => this.Opacity = 1.0;

        private void MenuItem_Opacity_Decrease_Click(object sender, EventArgs e)
        {
            double step = 0.02;
            this.Opacity = (this.Opacity <= 0.2 + step) ? 0.2 : this.Opacity - step;
        }

        private void MenuItem_Opacity_Increase_Click(object sender, EventArgs e)
        {
            double step = 0.02;
            this.Opacity = (this.Opacity >= 1.0 - step) ? 1.0 : this.Opacity + step;
        }

        private void MenuItem_Vertical_Flipping_Click(object sender, EventArgs e)
        {
            state_flip_vertical = !state_flip_vertical;
            UpdateMenuItemsChecked();
        }

        private void MenuItem_Horizontal_Flipping_Click(object sender, EventArgs e)
        {
            state_flip_horizontal = !state_flip_horizontal;
            UpdateMenuItemsChecked();
        }

        private void MenuItem_Normal_Border_Click(object sender, EventArgs e)
        {
            ApplyWindowStyle(0);
            UpdateMenuItemsChecked();
        }

        private void MenuItem_FullScreen_Click(object sender, EventArgs e)
        {
            ApplyWindowStyle(4);
            UpdateMenuItemsChecked();
        }

        private void MenuItem_Borderless_Rectangle_Click(object sender, EventArgs e)
        {
            ApplyWindowStyle(2);
            UpdateMenuItemsChecked();
        }

        private void MenuItem_Borderless_RoundedRectangle_Click(object sender, EventArgs e)
        {
            ApplyWindowStyle(3);
            UpdateMenuItemsChecked();
        }

        private void MenuItem_Borderless_Ellipse_Click(object sender, EventArgs e)
        {
            ApplyWindowStyle(1);
            UpdateMenuItemsChecked();
        }

        private void MenuItem_Center_Click(object sender, EventArgs e)
        {
            pictureBox1.SizeMode = PictureBoxSizeMode.CenterImage;
            this.Text = "tCamView (center)";
            UpdateMenuItemsChecked();
        }

        private void MenuItem_Stretch_Click(object sender, EventArgs e)
        {
            stretchKeepAspectRatio = false;
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            this.Text = "tCamView (stretch)";
            UpdateMenuItemsChecked();
        }

        private void MenuItem_AltStretch_Click(object sender, EventArgs e)
        {
            stretchKeepAspectRatio = true;
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            this.Text = "tCamView (alt.stretch)";
            UpdateMenuItemsChecked();
        }

        private void MenuItem_Zoom_Click(object sender, EventArgs e)
        {
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            this.Text = "tCamView (zoom)";
            UpdateMenuItemsChecked();
        }

        private void MenuItem_CropImageZoomIn_Click(object sender, EventArgs e) => cropSize += 5;
        private void MenuItem_CropImageZoomOut_Click(object sender, EventArgs e) => cropSize -= 5;

        private void MenuItem_ResumeCameraPreviewFromClipboard_Click(object sender, EventArgs e)
        {
            if (cam == null)
            {
                openVideoCaptureDevice(currCamID, currSizeID);
                currMenuItem0VideoCaptureDevices = currCamID;
                currMenuItem1VideoCapabilities = currSizeID;
                UpdateMenuItemsChecked();
            }
        }

        private void MenuItem_IncreaseWindowSize_Click(object sender, EventArgs e)
        {
            if (Width > 4000 || Height > 2000) return;
            float ratio = (float)Height / Width;
            Size = new Size(Width + 5, (int)(Height + 5 * ratio + 0.5));
            RefreshRegion();
        }

        private void MenuItem_DecreaseWindowSize_Click(object sender, EventArgs e)
        {
            if (Width < 50 || Height < 50) return;
            float ratio = (float)Height / Width;
            Size = new Size(Width - 5, (int)(Height - 5 * ratio + 0.5));
            RefreshRegion();
        }

        private void RefreshRegion()
        {
            if (currMenuItem3WindowStyles == 1)
                Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, Width, Height));
            else if (currMenuItem3WindowStyles == 2)
                Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 0, 0));
            else if (currMenuItem3WindowStyles == 3)
            {
                int m = Math.Min(Width, Height) / 4;
                Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, m, m));
            }
        }

        // -------------------------------------------------------
        //  KEYBOARD
        // -------------------------------------------------------
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            double opacityStep = 0.02;
            if (keyData == Keys.Up)
            {
                this.Opacity = (this.Opacity >= 1.0 - opacityStep) ? 1.0 : this.Opacity + opacityStep;
                return true;
            }
            else if (keyData == Keys.Down)
            {
                this.Opacity = (this.Opacity <= 0.2 + opacityStep) ? 0.2 : this.Opacity - opacityStep;
                return true;
            }
            else if (keyData == Keys.Right) { this.Opacity = 1.0; return true; }
            else if (keyData == Keys.Left)  { this.Opacity = 0.2; return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == (Keys.Control | Keys.C)) { MenuItem_SetImageToClipboard_Click(null, null); return; }
            if (e.KeyData == (Keys.Control | Keys.V)) { MenuItem_GetImageFromClipboard_Click(null, null); return; }

            switch (e.KeyData)
            {
                case Keys.Z:  MenuItem_Zoom_Click(null, null); return;
                case Keys.X:  MenuItem_Stretch_Click(null, null); return;
                case Keys.C:  MenuItem_Center_Click(null, null); return;
                case Keys.A:  MenuItem_AltStretch_Click(null, null); return;
                case Keys.N:
                case Keys.Escape: MenuItem_Normal_Border_Click(null, null); return;
                case Keys.E:  MenuItem_Borderless_Ellipse_Click(null, null); return;
                case Keys.R:  MenuItem_Borderless_Rectangle_Click(null, null); return;
                case Keys.W:  MenuItem_Borderless_RoundedRectangle_Click(null, null); return;
                case Keys.F:  MenuItem_FullScreen_Click(null, null); return;
                case Keys.H:  MenuItem_Horizontal_Flipping_Click(null, null); return;
                case Keys.V:  MenuItem_Vertical_Flipping_Click(null, null); return;
                case Keys.T:  MenuItem_AlwaysOnTop_Click(null, null); return;
                case Keys.K:  ToggleLockAspectRatio(); return;  // NEW
                case Keys.Q:  Application.Exit(); return;
                case Keys.L:  this.WindowState = FormWindowState.Minimized; return;
                case Keys.Space: MenuItem_ResumeCameraPreviewFromClipboard_Click(null, null); return;
                case Keys.P: MenuItem_IncreaseWindowSize_Click(null, null); return;
                case Keys.M: MenuItem_DecreaseWindowSize_Click(null, null); return;
                case Keys.G: MenuItem_GetImageFromClipboard_Click(null, null); return;
                case Keys.I: MenuItem_SetImageToClipboard_Click(null, null); return;
                case Keys.D: MenuItem_SetImageToClipboardAfter5Sec_Click(null, null); return;
                case Keys.S: MenuItem_CopyScreenToClipboard_Click(null, null); return;
                case Keys.PageUp:   MenuItem_CropImageZoomIn_Click(null, null); return;
                case Keys.PageDown: MenuItem_CropImageZoomOut_Click(null, null); return;
            }
        }

        // -------------------------------------------------------
        //  MOUSE DRAG (move borderless window)
        // -------------------------------------------------------
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        // -------------------------------------------------------
        //  CLIPBOARD HELPER
        // -------------------------------------------------------
        private Image GetImageFromClipboard()
        {
            if (Clipboard.GetDataObject() == null) return null;
            if (Clipboard.GetDataObject().GetDataPresent(DataFormats.Dib))
            {
                var dib = ((System.IO.MemoryStream)Clipboard.GetData(DataFormats.Dib)).ToArray();
                var width  = BitConverter.ToInt32(dib, 4);
                var height = BitConverter.ToInt32(dib, 8);
                var bpp    = BitConverter.ToInt16(dib, 14);
                if (bpp == 32)
                {
                    var gch = GCHandle.Alloc(dib, GCHandleType.Pinned);
                    Bitmap bmp = null;
                    try
                    {
                        var ptr = new IntPtr((long)gch.AddrOfPinnedObject() + 52);
                        bmp = new Bitmap(width, height, width * 4, System.Drawing.Imaging.PixelFormat.Format32bppArgb, ptr);
                        bmp.RotateFlip(RotateFlipType.RotateNoneFlipY);
                        return new Bitmap(bmp);
                    }
                    finally
                    {
                        gch.Free();
                        if (bmp != null) bmp.Dispose();
                    }
                }
            }
            return Clipboard.ContainsImage() ? Clipboard.GetImage() : null;
        }
    }
}
