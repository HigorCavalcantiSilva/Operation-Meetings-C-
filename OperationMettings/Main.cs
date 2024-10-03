using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using dotenv.net;
using DotNetEnv;
using Microsoft.Win32;

namespace OperationMettings
{
    public partial class Main : Form
    {
        const int GWL_STYLE = -16; // Índice para pegar o estilo da janela
        const int WS_CAPTION = 0xC00000; // Borda da janela com barra de título e botões
        const int WS_THICKFRAME = 0x00040000; // Borda redimensionável

        [DllImport("user32.dll")]
        static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private List<IntPtr> handleList;

        public Main()
        {
            InitializeComponent();
            handleList = new List<IntPtr>();
            DotEnv.Load();
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (var handle in handleList)
            {
                try
                {
                    if (handle != null)
                        SetParent(handle, IntPtr.Zero);
                        SendMessage(handle, 0x0010, IntPtr.Zero, IntPtr.Zero);
                } catch
                {
                    // Se cair aqui é porque o processo foi encerrado manualmente
                }
            }
        }

        private void InitializeZoom()
        {
            string meetingId = Environment.GetEnvironmentVariable("MEETING_ID");
            string password = Environment.GetEnvironmentVariable("PASSWORD");

            Process process = new Process();
            process.StartInfo.FileName = Environment.GetEnvironmentVariable("PATH_EXECUTABLE_ZOOM");
            process.StartInfo.Arguments = $"--url=zoommtg://zoom.us/join?confno={meetingId}&pwd={password}";
            process.Start();

            IntPtr handle = FindWindow(null, "Zoom Reunião");
            while (handle == IntPtr.Zero)
            {
                Thread.Sleep(1000);
                handle = FindWindow(null, "Zoom Reunião");
            }

            Invoke((MethodInvoker)(() =>
            {
                int style = GetWindowLong(handle, GWL_STYLE);
                SetWindowLong(handle, GWL_STYLE, style & ~WS_CAPTION & ~WS_THICKFRAME);

                SetParent(handle, panel3.Handle);

                MoveWindow(handle, 0, 0, panel3.Width, panel3.Height, true);

                ShowWindow(handle, 5);

                handleList.Add(handle);
            }));
        }

        private void OpenProcess(Panel panel, ProcessStartInfo startInfo)
        {
            Process process = Process.Start(startInfo);
            IntPtr handle;

            handle = process.MainWindowHandle;

            while (handle == IntPtr.Zero)
            {
                Thread.Sleep(1000);    
                handle = process.MainWindowHandle;
            }

            Invoke((MethodInvoker)(() =>
            {
                int style = GetWindowLong(handle, GWL_STYLE);
                SetWindowLong(handle, GWL_STYLE, style & ~WS_CAPTION & ~WS_THICKFRAME);

                SetParent(handle, panel.Handle);

                MoveWindow(handle, 0, 0, panel.Width, panel.Height, true);

                ShowWindow(handle, 5);

                handleList.Add(handle);
            }));
        }

        private void Main_Shown(object sender, EventArgs e)
        {
            // Obter o nome do executável do aplicativo atual
            string appName = System.IO.Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            // Configurar as chaves de emulação do navegador para usar IE11
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey($@"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"))
            {
                key.SetValue(appName, (uint)11000, RegistryValueKind.DWord);
            }

            webBrowser1.Navigate("http://192.168.100.102:8096/Index");
            InitializeZoom();
        }
    }
}
