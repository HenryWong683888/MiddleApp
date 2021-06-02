using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using WebSocketSharp;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Drawing.Drawing2D;


namespace MeetNoteApp
{
    public partial class Frm_Main : Form
    {

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool BringWindowToTop(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(String lpClassName, String lpWindowName);

        [DllImport("User32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int cmdShow);
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        private const int WS_SHOWNORMAL = 1;

        public Frm_Main()
        {
            InitializeComponent();
        }

        //记录直线或者曲线的对象
        private System.Drawing.Drawing2D.GraphicsPath mousePath = new System.Drawing.Drawing2D.GraphicsPath();
        //画笔透明度
        private int myAlpha = 100;
        //画笔颜色对象
        private Color myUserColor = new Color();
        //画笔宽度
        private int myPenWidth = 3;
        //签名的图片对象
        public Bitmap SavedBitmap;

        private void Form1_Load(object sender, EventArgs e)
        {
            

            //讓程式在工具列中隱藏
            this.ShowInTaskbar = false;
            //隱藏程式本身的視窗
            this.Hide();

            showOnMonitor(1);

        }

        void showOnMonitor(int showOnMonitor)
        {
            Screen[] sc;
            sc = Screen.AllScreens;
            //get all the screen width and heights 
            
            this.FormBorderStyle = FormBorderStyle.None;
            this.Left = sc[0].Bounds.Width;
            this.Top = 0;
            this.StartPosition = FormStartPosition.Manual;
            this.Show();
            this.WindowState= System.Windows.Forms.FormWindowState.Maximized;
        }
        private Point p1, p2;
        private static bool drawing = false;
        Graphics g;


        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
           
                g = pictureBox1.CreateGraphics();
                
                    if (drawing)
                    {
                        drawing = true;  
                        Point currentPoint = new Point(e.X, e.Y);
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        g.DrawLine(new Pen(Color.Black, 3), p2, currentPoint);

                        p2.X = currentPoint.X;
                        p2.Y = currentPoint.Y;
                    }            
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            Graphics g = ((PictureBox)sender).CreateGraphics();
            g.FillEllipse(Brushes.Black, e.X, e.Y, 4, 4);

                p1 = new Point(e.X, e.Y);
                p2 = new Point(e.X, e.Y);
               
                drawing = true;

            
            
        }


        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            drawing = false;
        }

        #region 参数设置
        public void set()
        {
            //画笔宽度
            myPenWidth = 2;
            //myUserColor = System.Drawing.Color.Blue;
            //myAlpha = 100;
        }
        #endregion

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            try
            {
               
                set();//设置画笔的颜色、宽度、透明度
                Pen CurrentPen = new Pen(Color.Black, myPenWidth);
                e.Graphics.DrawPath(CurrentPen, mousePath);
                
            }
            catch { }
        }

        private void btn_Clear_Click(object sender, EventArgs e)
        {
            pictureBox1.CreateGraphics().Clear(Color.White);
            mousePath.Reset();
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            bool isSave = true;
            SavedBitmap = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            pictureBox1.DrawToBitmap(SavedBitmap, new Rectangle(0, 0, pictureBox1.Width, pictureBox1.Height));
           
            #region 保存图片
            SaveFileDialog saveImageDialog = new SaveFileDialog();
            saveImageDialog.Title = "图片保存";
            saveImageDialog.Filter = @"jpeg|*.jpg|bmp|*.bmp|gif|*.gif";
            if (saveImageDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveImageDialog.FileName.ToString();
                if (fileName != "" && fileName != null)
                {
                    string fileExtName = fileName.Substring(fileName.LastIndexOf(".") + 1).ToString();
                    System.Drawing.Imaging.ImageFormat imgformat = null;
                    //默认保存为JPG格式   
                    if (imgformat == null)
                    {
                        imgformat = System.Drawing.Imaging.ImageFormat.Jpeg;
                    }
                    if (isSave)
                    {
                        try
                        {
                            SavedBitmap.Save(fileName, System.Drawing.Imaging.ImageFormat.Bmp);
                            SavedBitmap.Dispose();
                            MessageBox.Show("图片已经成功保存!");   
                        }
                        catch
                        {
                            MessageBox.Show("保存失败,你还没有截取过图片或已经清空图片!");
                        }
                    }
                }
            }
            #endregion
        }

        public Bitmap SaveImage(PictureBox pictureBox1)
        {
            Bitmap SavedBitmap = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            pictureBox1.DrawToBitmap(SavedBitmap, new Rectangle(0, 0, pictureBox1.Width, pictureBox1.Height));
            return SavedBitmap;
        }


        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
        enum TernaryRasterOperations : uint
        {
            /// <summary>dest = source</summary>
            SRCCOPY = 0x00CC0020,
            /// <summary>dest = source OR dest</summary>
            SRCPAINT = 0x00EE0086,
            /// <summary>dest = source AND dest</summary>
            SRCAND = 0x008800C6,
            /// <summary>dest = source XOR dest</summary>
            SRCINVERT = 0x00660046,
            /// <summary>dest = source AND (NOT dest)</summary>
            SRCERASE = 0x00440328,
            /// <summary>dest = (NOT source)</summary>
            NOTSRCCOPY = 0x00330008,
            /// <summary>dest = (NOT src) AND (NOT dest)</summary>
            NOTSRCERASE = 0x001100A6,
            /// <summary>dest = (source AND pattern)</summary>
            MERGECOPY = 0x00C000CA,
            /// <summary>dest = (NOT source) OR dest</summary>
            MERGEPAINT = 0x00BB0226,
            /// <summary>dest = pattern</summary>
            PATCOPY = 0x00F00021,
            /// <summary>dest = DPSnoo</summary>
            PATPAINT = 0x00FB0A09,
            /// <summary>dest = pattern XOR dest</summary>
            PATINVERT = 0x005A0049,
            /// <summary>dest = (NOT dest)</summary>
            DSTINVERT = 0x00550009,
            /// <summary>dest = BLACK</summary>
            BLACKNESS = 0x00000042,
            /// <summary>dest = WHITE</summary>
            WHITENESS = 0x00FF0062,
            /// <summary>
            /// Capture window as seen on screen.  This includes layered windows 
            /// such as WPF windows with AllowsTransparency="true"
            /// </summary>
            CAPTUREBLT = 0x40000000
        }

        [DllImport("gdi32.dll", EntryPoint = "BitBlt", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool BitBlt([In] IntPtr hdc, int nXDest, int nYDest, int nWidth, int nHeight, [In] IntPtr hdcSrc, int nXSrc, int nYSrc, TernaryRasterOperations dwRop);

        public static Bitmap CopyGraphicsContent(Graphics source, PictureBox pictureBox1)
        {
            Bitmap bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            using (Graphics dest = Graphics.FromImage(bmp))
            {
                IntPtr hdcSource = source.GetHdc();
                IntPtr hdcDest = dest.GetHdc();

                BitBlt(hdcDest, 0, 0, pictureBox1.Width, pictureBox1.Height, hdcSource, 0, 0, TernaryRasterOperations.SRCCOPY);

                source.ReleaseHdc(hdcSource);
                dest.ReleaseHdc(hdcDest);
            }

            return bmp;
        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetImage(CopyGraphicsContent(g, pictureBox1));
        }


        public static void HandleRunningInstance(System.Diagnostics.Process instance)
        {
            // 相同時透過ShowWindowAsync還原，以及SetForegroundWindow將程式至於前景
            ShowWindowAsync(instance.MainWindowHandle, WS_SHOWNORMAL);
            SetForegroundWindow(instance.MainWindowHandle);
        }

        private void button2_Click(object sender, EventArgs e)
        {

           
            if (System.Diagnostics.Process.GetProcessesByName("OUTLOOK").Length > 0)
            {
                HandleRunningInstance(System.Diagnostics.Process.GetProcessesByName("OUTLOOK")[0]);

               
            }
            else
            {
                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\\microsoft\\windows\\currentversion\\app paths\\OUTLOOK.EXE");
                string path = (string)key.GetValue("Path");
                if (path != null)
                {
                    System.Diagnostics.Process.Start("OUTLOOK.EXE");

                  
                }
                else
                    MessageBox.Show("There is no Outlook in this computer!", "SystemError", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

             
            }

           // OpenOutlook();

            using (var ws = new WebSocket("ws://localhost:13812"))
            {
                ws.OnMessage += (s, e1) =>
                {
                    Console.WriteLine("Laputa says: " + e1.Data);
                    
                };
                    
                ws.Connect();
                ws.Send("{\"Status\":\"setting\",\"Mode\":\"3\"," +
                    "\"Handwriting\":\"1\",\"Gamma\":\"4\",\"PenWidth\":" +
                    "\"3\",\"EraserWidth\":\"10\",\"DisplayMode\":\"1\"}");
                //Console.ReadKey(true);


            }


            //讓程式在工具列中隱藏
            this.ShowInTaskbar = false;
            //隱藏程式本身的視窗
            this.Hide();

            System.Environment.Exit(System.Environment.ExitCode);

        }


        static void OpenOutlook()
        {

            System.Diagnostics.Process[] telerecProcs = System.Diagnostics.Process.GetProcessesByName("OUTLOOK.EXE");
            if (telerecProcs.Length > 0)
            {
                bringToFront("OUTLOOK"); //填入視窗的Title Name                      
            }
            else
            {
                System.Diagnostics.Process sample = new System.Diagnostics.Process();
                sample.StartInfo.FileName = "OUTLOOK.EXE";
                sample.Start();
            }

            /*Outlook.Application outlookObj = null;

            if (System.Diagnostics.Process.GetProcessesByName("OUTLOOK").Count().Equals(0))
            {
                System.Diagnostics.Process.Start("OUTLOOK.EXE");
            }
            else
            {
                outlookObj = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
               
            }*/
        }

        public static void bringToFront(string title)
        {
            // Get a handle to the Calculator application.
            IntPtr handle = FindWindow(null, title);
            // Verify that Calculator is a running process.
            if (handle == IntPtr.Zero)
            {
                return;
            }
            BringWindowToTop(handle); // 將視窗浮在最上層
            ShowWindow(handle, 3); // 將視窗最大化
        }

    }
}
