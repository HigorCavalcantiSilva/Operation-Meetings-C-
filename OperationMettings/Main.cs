using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using dotenv.net;

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

        private void InitializeObsStudio()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WorkingDirectory = Environment.GetEnvironmentVariable("PATH_OBS");
            startInfo.FileName = $"{startInfo.WorkingDirectory}{Environment.GetEnvironmentVariable("EXECUTABLE_OBS")}";

            OpenProcess(panel1, startInfo);
        }

        private void InitializeJWLibrary()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = $"shell:appsFolder\\{Environment.GetEnvironmentVariable("PATH_JW")}!App";

            OpenProcessJW(startInfo);
        }

        private void InitializeOnlyT()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WorkingDirectory = Environment.GetEnvironmentVariable("PATH_ONLYT");
            startInfo.FileName = $"{startInfo.WorkingDirectory}{Environment.GetEnvironmentVariable("EXECUTABLE_ONLYT")}";

            OpenProcess(panel4, startInfo);
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

        private void OpenProcessJW(ProcessStartInfo startInfo)
        {
            IntPtr handle;
            Process process = Process.Start(startInfo);

            Thread.Sleep(4000);
            handle = FindWindow(null, "JW Library");

            Invoke((MethodInvoker)(() =>
            {
                MoveWindow(handle, Screen.PrimaryScreen.Bounds.Width - 375, 0, 300, 768, true);

                handleList.Add(handle);
            }));
        }

        private async void Main_Shown(object sender, EventArgs e)
        {
            // Executar os métodos simultaneamente usando Task.Run para cada método.
            Task task1 = Task.Run(() => InitializeOnlyT());
            Task task2 = Task.Run(() => InitializeObsStudio());
            Task task3 = Task.Run(() => InitializeZoom());
            Task task4 = Task.Run(() => InitializeJWLibrary());

            // Aguarda todos os métodos terminarem.
            await Task.WhenAll(task1, task2, task3, task4);
        }
    }
}
