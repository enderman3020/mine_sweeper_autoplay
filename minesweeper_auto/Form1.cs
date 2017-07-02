using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace minesweeper_auto
{
    public partial class Form1 : Form
    {
        public bool is_running = true;
        public bool restart_ = false;
        
        public int cell_width = 16, cell_height = 16; // マスの大きさ

        public int process_span = 150; // 実行周期(ms)

        public int map_width = 8, map_height = 8; // 横方向と縦方向のマスの数
                                                  // メモ
                                                  // 初級: 8 * 8
                                                  // 中級: 16 * 16
                                                  // 上級: 30 * 16
                                                  // 超上級: 48 * 24

        public int[,] map_field;
        public bool[,] map_bomb;
        public bool[,] map_flag;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            is_running = false;
            KeyChecker();
            Initializing();
            processing();
        }

        public void Initializing()
        {
            SendKeys.Send("{F2}");
            map_field = new int[map_width, map_height];
            map_bomb = new bool[map_width, map_height];
            map_flag = new bool[map_width, map_height];
            label1.Text = "waiting...";
        }

        // ウインドウを一番前にする
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        // クライアントの座標を取得
        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hwnd, out Point lpPoint);

        // クライアントの大きさを取得
        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hwnd, out Rectangle lpRect);

        // 最小化からの復元
        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        // 最小化のチェック
        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        /// <summary>
        /// mine2000が起動しているか確認
        /// </summary>
        private bool mine2000_find(ref Point point, ref Rectangle rect)
        {
            const int SW_RESTORE = 9;
            foreach (Process p in Process.GetProcesses())
            {
                //メインウィンドウのタイトルがある時だけ列挙する
                if (p.MainWindowTitle.Length != 0)
                {
                    Trace.WriteLine("プロセス名:" + p.ProcessName);
                    Trace.WriteLine("タイトル名:" + p.MainWindowTitle);
                }
                if (p.ProcessName == "mine2000")
                {
                    if (IsIconic(p.MainWindowHandle))
                    {
                        ShowWindowAsync(p.MainWindowHandle, SW_RESTORE);
                    }
                    SetForegroundWindow(p.MainWindowHandle);

                    ClientToScreen(p.MainWindowHandle, out point);
                    GetClientRect(p.MainWindowHandle, out rect);
                    Trace.WriteLine("座標: (" + point.X.ToString() + ", " + point.Y.ToString() + ")");
                    Trace.WriteLine("大きさ: (" + rect.Width.ToString() + ", " + rect.Height.ToString() + ")");

                    return true;
                }
            }
            return false;
        }

        // マウスのイベント
        [DllImport("USER32.dll", CallingConvention = CallingConvention.StdCall)]
        static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        private const int MOUSEEVENTF_LEFTDOWN = 0x2;
        private const int MOUSEEVENTF_LEFTUP = 0x4;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x8;
        private const int MOUSEEVENTF_RIGHTUP = 0x16;

        private async void processing()
        {
            CancellationTokenSource cancelsource = new CancellationTokenSource();

            Point mine2000_pos = new Point();
            Rectangle mine2000_rect = new Rectangle();
            Point mine2000_offset = new Point(2, 34);

            mine2000_find(ref mine2000_pos, ref mine2000_rect);

            var progress1 = new Progress<int>((hoge) =>
            {
                if (is_running)
                {
                    Bitmap img_bit = copy_from_screen(new Point(mine2000_pos.X + mine2000_offset.X, mine2000_pos.Y + mine2000_offset.Y), cell_width * map_width, cell_height * map_height);
                    BitmapPlus img_bit_p = new BitmapPlus(img_bit);
                    img_bit_p.BeginAccess();
                    for (int x = 0; x < map_width; x++)
                    {
                        for (int y = 0; y < map_height; y++)
                        {
                            map_field[x, y] = RGBsum_to_number(img_bit_p, x, y, cell_width, cell_height);
                        }
                    }
                    img_bit_p.EndAccess();
                    img_bit.Dispose();

                    // 操作をしたか
                    bool is_exist = false;

                    // 旗を立てる
                    for (int x = 0; x < map_width; x++)
                    {
                        for (int y = 0; y < map_height; y++)
                        {
                            if (map_field[x, y] > 0)
                            {
                                int count = 0;
                                for (int dx = -1; dx <= 1; dx++)
                                {
                                    for (int dy = -1; dy <= 1; dy++)
                                    {
                                        if (x + dx >= 0 && x + dx < map_width && y + dy >= 0 && y + dy < map_height)
                                        {
                                            if (map_field[x + dx, y + dy] < 0)
                                            {
                                                count++;
                                            }
                                        }
                                    }
                                }
                                if (count == map_field[x, y])
                                {
                                    for (int dx = -1; dx <= 1; dx++)
                                    {
                                        for (int dy = -1; dy <= 1; dy++)
                                        {
                                            if (x + dx >= 0 && x + dx < map_width && y + dy >= 0 && y + dy < map_height)
                                            {
                                                if (map_field[x + dx, y + dy] < 0)
                                                {
                                                    map_bomb[x + dx, y + dy] = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                /*
                                else if (count < map_field[x, y]) // 操作ミス用(?)
                                {
                                    for (int dx = -1; dx <= 1; dx++)
                                    {
                                        for (int dy = -1; dy <= 1; dy++)
                                        {
                                            if (x + dx >= 0 && x + dx < map_width && y + dy >= 0 && y + dy < map_height)
                                            {
                                                if (map_field[x + dx, y + dy] < 0)
                                                {
                                                    map_bomb[x + dx, y + dy] = false;
                                                }
                                            }
                                        }
                                    }
                                }
                                */
                            }
                        }
                    }
                    if (MouseButtons == MouseButtons.None)
                    {
                        Point mouse_pos = Cursor.Position;
                        for (int x = 0; x < map_width; x++)
                        {
                            for (int y = 0; y < map_height; y++)
                            {
                                if (map_bomb[x, y] && map_field[x, y] != -10 && !map_flag[x, y])
                                {
                                    Cursor.Position = new Point(mine2000_pos.X + mine2000_offset.X + cell_width * x, mine2000_pos.Y + mine2000_offset.Y + cell_height * y);
                                    mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                                    mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                                    map_flag[x, y] = true;

                                    is_exist = true;
                                }
                            }
                        }
                        Cursor.Position = mouse_pos;
                    }
                    // マスを拓く

                    if (MouseButtons == MouseButtons.None)
                    {
                        Point mouse_pos = Cursor.Position;

                        for (int x = 0; x < map_width; x++)
                        {
                            for (int y = 0; y < map_height; y++)
                            {
                                if (map_field[x, y] > 0)
                                {
                                    int count = 0;
                                    bool valid = false;
                                    for (int dx = -1; dx <= 1; dx++)
                                    {
                                        for (int dy = -1; dy <= 1; dy++)
                                        {
                                            if (x + dx >= 0 && x + dx < map_width && y + dy >= 0 && y + dy < map_height)
                                            {
                                                if (map_field[x + dx, y + dy] == -10)
                                                {
                                                    count++;
                                                }
                                                else if (map_field[x + dx, y + dy] == -1)
                                                {
                                                    valid = true;
                                                }
                                            }
                                        }
                                    }
                                    if (count == map_field[x, y] && valid)
                                    {
                                        Cursor.Position = new Point(mine2000_pos.X + mine2000_offset.X + cell_width * x, mine2000_pos.Y + mine2000_offset.Y + cell_height * y);
                                        mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                                        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                                        mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                                        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                                        is_exist = true;
                                    }
                                }
                            }
                        }

                        Cursor.Position = mouse_pos;
                    }
                    
                    if (!is_exist)
                    {
                        label1.Text = "Nothing";
                        
                        Point mouse_pos = Cursor.Position;
                        
                        for (int x = 0; x < map_width; x++)
                        {
                            for (int y = 0; y < map_height; y++)
                            {
                                if (map_field[x, y] == -1)
                                {
                                    Cursor.Position = new Point(mine2000_pos.X + mine2000_offset.X + cell_width * x, mine2000_pos.Y + mine2000_offset.Y + cell_height * y);
                                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                                    is_exist = true;
                                    break;
                                }
                            }
                            if (is_exist)
                            {
                                break;
                            }
                        }
                        
                        is_exist = false;
                        Cursor.Position = mouse_pos;
                        
                    }
                    
                }
                if (pictureBox1.Image != null)
                {
                    pictureBox1.Image.Dispose();
                }
                pictureBox1.Image = copy_from_screen(new Point(mine2000_pos.X + mine2000_offset.X, mine2000_pos.Y + mine2000_offset.Y), cell_width * map_width, cell_height * map_height);

                if (!is_running)
                {
                    label1.Text = "stop";
                }
                if (restart_)
                {
                    Initializing();
                    restart_ = false;
                }
            });
            bool ret = await Span1(progress1, cancelsource.Token, process_span);
        }

        private Task<bool> Span1(IProgress<int> progress, CancellationToken cancelltoken, int wait_time)
        {
            return Task.Run(() =>
            {
                for (int i = 0; i < 1000000; i++)
                {
                    if (cancelltoken.IsCancellationRequested)
                    {
                        return false;
                    }
                    progress.Report(i);
                    Thread.Sleep(wait_time);
                }
                return true;
            });
        }

        /// <summary>
        /// 画面から画像を取得
        /// </summary>
        public Bitmap copy_from_screen(Point p1, int width, int height)
        {
            //Bitmapの作成
            Bitmap bmp = new Bitmap(width, height);
            //Graphicsの作成
            Graphics g = Graphics.FromImage(bmp);
            //画面をコピーする
            g.CopyFromScreen(p1, new Point(0, 0), bmp.Size);
            //解放
            g.Dispose();

            return bmp;
        }
        
        /// <summary>
        /// 画像の中心線上のRGBの合計値を計算し、該当する数字を返す
        /// </summary>
        /// <param name="img">画像</param>
        /// <returns></returns>
        public int RGBsum_to_number(BitmapPlus img_bit_p, int pos_x, int pos_y, int cell_width, int cell_height)
        {
            int R_sum = 0, G_sum = 0, B_sum = 0;
            for (int x = 0; x < cell_width; x++)
            {
                Color pixeldata = img_bit_p.GetPixel(cell_width * pos_x + x, cell_height * pos_y + cell_height / 2);
                R_sum += pixeldata.R;
                G_sum += pixeldata.G;
                B_sum += pixeldata.B;
            }
            
            if(pos_x == 0 && pos_y == 0)
            {
                //label1.Text = R_sum.ToString() + ", " + G_sum.ToString() + ", " + B_sum.ToString();
            }
            
            if (R_sum == 3008 && G_sum == 3008 && B_sum == 3008)
            {
                return 0;
            }
            else if(R_sum == 2432 && G_sum == 2432 && B_sum == 3197)
            {
                return 1;
            }
            else if (R_sum == 2048 && G_sum == 2688 && B_sum == 2048)
            {
                return 2;
            }
            else if(R_sum == 3386 && G_sum == 1856 && B_sum == 1856)
            {
                return 3;
            }
            else if (R_sum == 1088 && G_sum == 1088 && B_sum == 2368)
            {
                return 4;
            }
            else if (R_sum == 2368 && G_sum == 1088 && B_sum == 1088)
            {
                return 5;
            }
            else if (R_sum == 1088 && G_sum == 2368 && B_sum == 2368)
            {
                return 6;
            }
            else if (R_sum == 2432 && G_sum == 2432 && B_sum == 2432)
            {
                return 7;
            }
            else if (R_sum == 3070 && G_sum == 3070 && B_sum == 3070)
            {
                return -1;
            }
            else if (R_sum == 2878 && G_sum == 2878 && B_sum == 2878)
            {
                return -10;
            }
            else if (R_sum == 512 && G_sum == 512 && B_sum == 512)
            {
                restart_ = true;
            }
            else if (R_sum == 638 && G_sum == 128 && B_sum == 128)
            {
                restart_ = true;
            }
            else
            {
                label1.Text = "error";
                //is_running = false;
            }
            return -1;
            // メモ
            // 埋: 3070, 3070, 3070
            // 旗: 2878, 2878, 2878
            // 0 : 3008, 3008, 3008
            // 1 : 2432, 2432, 3197
            // 2 : 2048, 2688, 2048
            // 3 : 3386, 1856, 1856
            // 4 : 1088, 1088, 2368
            // 5 : 2368, 1088, 1088
            // 6 : 1088, 2368, 2368
            // 7 : 2432, 2432, 2432
            // × :  512,  512,  512 
            // ×2:  638,  128,  128
        }

        // キーが押下されたかどうか確認する
        [DllImport("user32")]
        static extern short GetAsyncKeyState(Keys vKey);

        /// <summary>
        /// 指定されたキーが押下されたか確認
        /// </summary>
        private async void KeyChecker()
        {
            var progress = new Progress<int>((hoge) =>
            {
                if (GetAsyncKeyState(Keys.Escape) < 0)
                {
                    this.Close();
                }
                else if (GetAsyncKeyState(Keys.Space) < 0)
                {
                    is_running = true;
                }
                else if (GetAsyncKeyState(Keys.B) < 0)
                {
                    is_running = false;
                }
                else if (GetAsyncKeyState(Keys.F2) < 0)
                {
                    Initializing();
                    is_running = true;
                }
            });
            CancellationTokenSource cancelsource = new CancellationTokenSource();
            var ret = await Span1(progress, cancelsource.Token, 1000);
        }
        
        /// <summary>
        /// Bitmap処理を高速化するためのクラス
        /// </summary>
        public class BitmapPlus
        {
            /// <summary>
            /// オリジナルのBitmapオブジェクト
            /// </summary>
            private Bitmap _bmp = null;

            /// <summary>
            /// Bitmapに直接アクセスするためのオブジェクト
            /// </summary>
            private BitmapData _img = null;

            /// <summary>
            /// コンストラクタ
            /// </summary>
            /// <param name="original"></param>
            public BitmapPlus(Bitmap original)
            {
                // オリジナルのBitmapオブジェクトを保存
                _bmp = original;
            }

            /// <summary>
            /// Bitmap処理の高速化開始
            /// </summary>
            public void BeginAccess()
            {
                // Bitmapに直接アクセスするためのオブジェクト取得(LockBits)
                _img = _bmp.LockBits(new Rectangle(0, 0, _bmp.Width, _bmp.Height),
                    ImageLockMode.ReadWrite,
                    PixelFormat.Format24bppRgb);
            }

            /// <summary>
            /// Bitmap処理の高速化終了
            /// </summary>
            public void EndAccess()
            {
                if (_img != null)
                {
                    // Bitmapに直接アクセスするためのオブジェクト開放(UnlockBits)
                    _bmp.UnlockBits(_img);
                    _img = null;
                }
            }

            /// <summary>
            /// BitmapのGetPixel同等
            /// </summary>
            /// <param name="x">Ｘ座標</param>
            /// <param name="y">Ｙ座標</param>
            /// <returns>Colorオブジェクト</returns>
            public Color GetPixel(int x, int y)
            {
                if (_img == null)
                {
                    // Bitmap処理の高速化を開始していない場合はBitmap標準のGetPixel
                    return _bmp.GetPixel(x, y);
                }

                // Bitmap処理の高速化を開始している場合はBitmapメモリへの直接アクセス
                IntPtr adr = _img.Scan0;
                int pos = x * 3 + _img.Stride * y;
                byte b = System.Runtime.InteropServices.Marshal.ReadByte(adr, pos + 0);
                byte g = System.Runtime.InteropServices.Marshal.ReadByte(adr, pos + 1);
                byte r = System.Runtime.InteropServices.Marshal.ReadByte(adr, pos + 2);
                return Color.FromArgb(r, g, b);
            }

            /// <summary>
            /// BitmapのSetPixel同等
            /// </summary>
            /// <param name="x">Ｘ座標</param>
            /// <param name="y">Ｙ座標</param>
            /// <param name="col">Colorオブジェクト</param>
            public void SetPixel(int x, int y, Color col)
            {
                if (_img == null)
                {
                    // Bitmap処理の高速化を開始していない場合はBitmap標準のSetPixel
                    _bmp.SetPixel(x, y, col);
                    return;
                }

                // Bitmap処理の高速化を開始している場合はBitmapメモリへの直接アクセス
                IntPtr adr = _img.Scan0;
                int pos = x * 3 + _img.Stride * y;
                System.Runtime.InteropServices.Marshal.WriteByte(adr, pos + 0, col.B);
                System.Runtime.InteropServices.Marshal.WriteByte(adr, pos + 1, col.G);
                System.Runtime.InteropServices.Marshal.WriteByte(adr, pos + 2, col.R);
            }
        }
    }
}
