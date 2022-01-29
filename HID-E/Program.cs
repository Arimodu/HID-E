using System.Runtime.InteropServices;
using System.Windows.Input;
using System.Windows.Forms;
using System;
using System.Threading;

namespace HID_E
{
    internal class Program
    {
        [DllImport("user32.dll")]
        static extern bool GetCursorPos(out POINT point);

        public struct POINT
        {
            public int x;
            public int y;
        }

        public static POINT GetMousePosition()
        {
            GetCursorPos(out POINT pos);
            return pos;
        }

        [DllImport("user32.dll")]
        static extern bool GetAsyncKeyState(int button);
        public static bool IsMouseButtonPressed(MouseButton button)
        {
            int k;
            switch (button)
            {
                case MouseButton.Left:
                    k = 0x01;
                    break;
                case MouseButton.Middle:
                    return false;
                case MouseButton.Right:
                    k = 0x02;
                    break;
                case MouseButton.XButton1:
                    return false;
                case MouseButton.XButton2:
                    return false;
                default:
                    return false;
            }
            var o = GetAsyncKeyState(k);
            //Console.WriteLine(o);
            return o;
        }

        //This is a replacement for Cursor.Position
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        //This simulates a left mouse click
        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr zeroOnly, string lpWindowName);

        static void Main(string[] args)
        {
            Console.SetWindowSize(55, 30);
            Console.WriteLine("------------------------------------------------------");
            Console.WriteLine("HID - E");
            Console.WriteLine("Made by Arimodu");
            Console.WriteLine("License: MIT");
            Console.WriteLine("------------------------------------------------------");

            Console.WriteLine("Press any key to enter setup...");
            Console.ReadLine();

            Console.WriteLine("Performing setup...");

            Thread.Sleep(1000);

            Console.WriteLine("Please click SET input field in Konsole...");

            while (true)
            {
                if (IsMouseButtonPressed(MouseButton.Left))
                {
                    Thread.Sleep(50);
                    if (IsMouseButtonPressed(MouseButton.Left)) break;
                }
            }
            
            POINT SETInputPos = GetMousePosition();
            WriteColored("SET Field position saved", ConsoleColor.Green);

            Thread.Sleep(500);

            Console.WriteLine("Please click this application title bar...");

            while (true)
            {
                if (IsMouseButtonPressed(MouseButton.Left))
                {
                    Thread.Sleep(50);
                    if (IsMouseButtonPressed(MouseButton.Left)) break;
                }
            }

            POINT AppTitleBar = GetMousePosition();
            WriteColored("Application title bar position saved", ConsoleColor.Green);

            //var oTitle = Console.Title;
            //var wGuid = Guid.NewGuid().ToString();
            //Console.Title = wGuid;
            //IntPtr handle = FindWindowByCaption(IntPtr.Zero, wGuid);
            //Console.Title = oTitle;

            Thread.Sleep(2000);
            ClearAndReplaceHeader();

            while (true)
            {
                Console.Write("Read / Input SET: ");
                string? nSET = Console.ReadLine();
                Console.WriteLine("Checking SET...");
                if (nSET == null || !nSET.StartsWith("SET"))
                {
                    Console.WriteLine("Invalid SET, please try again");
                    continue;
                }

                WriteColored("Valid SET confirmed, executing...", ConsoleColor.Green);
                string SET = nSET;

                LeftMouseClick(SETInputPos.x, SETInputPos.y);

                Console.WriteLine($"Sending {SET} to window...");
                SendKeys.SendWait(SET);
                SendKeys.SendWait("{ENTER}");

                LeftMouseClick(AppTitleBar.x, AppTitleBar.y);

                //SetForegroundWindow(handle);

                Console.WriteLine("Print documents? (Y/N/O)");
                ConsoleKey key;
                do
                {
                    key = Console.ReadKey(true).Key;
                    if (key == ConsoleKey.Y || key == ConsoleKey.N || key == ConsoleKey.O)
                        break;
                } while (true);

                if (key == ConsoleKey.N)
                {
                    WriteColored("Done", ConsoleColor.Green);
                    Thread.Sleep(3000);
                    ClearAndReplaceHeader();
                    LeftMouseClick(AppTitleBar.x, AppTitleBar.y);
                    continue;
                }

                Console.WriteLine("Opening SET details....");

                LeftMouseClick(SETInputPos.x, SETInputPos.y);
                SendKeys.SendWait("%{F1}");
                SendKeys.SendWait(SET);
                SendKeys.SendWait("{ENTER}");

                Thread.Sleep(5000);

                if (key == ConsoleKey.O)
                {
                    WriteColored("Done", ConsoleColor.Green);
                    Thread.Sleep(3000);
                    ClearAndReplaceHeader();
                    LeftMouseClick(AppTitleBar.x, AppTitleBar.y);
                    continue;
                }

                Console.WriteLine("Printing documents....");

                Thread.Sleep(100);

#if !DEBUG      
                SendKeys.SendWait("^p");
#endif
                for (int i = 0; i < 4; i++)
                    SendKeys.SendWait("{DOWN}");
                SendKeys.SendWait("{ENTER}");

                Thread.Sleep(100);

#if !DEBUG      
                SendKeys.SendWait("^p");
#endif
                for (int i = 0; i < 5; i++)
                    SendKeys.SendWait("{DOWN}");
                SendKeys.SendWait("{ENTER}");

                Thread.Sleep(100);

                SendKeys.SendWait("{ESC}");

                LeftMouseClick(AppTitleBar.x, AppTitleBar.y);

                WriteColored("Done", ConsoleColor.Green);
                //SetForegroundWindow(handle);

                Thread.Sleep(5000);

                ClearAndReplaceHeader();
            }
        }

        public static void ClearAndReplaceHeader()
        {
            Console.Clear();
            Console.WriteLine("------------------------------------------------------");
            Console.WriteLine("HID - E");
            Console.WriteLine("Made by Arimodu");
            Console.WriteLine("License: MIT");
            Console.WriteLine("------------------------------------------------------");
        }

        public static void WriteColored(object? value, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(value);
            Console.ForegroundColor = ConsoleColor.White;
        }
    }
}