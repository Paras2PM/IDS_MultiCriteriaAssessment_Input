using System;
using System.IO;
using System.Threading;
using WinApi.User32;


namespace TestIDS
{
    class Program
    {
        static string appName = "IDS - Intelligent Decision System for Multiple Criteria Assessment - ";
        static int time = 600;
        static void Main(string[] args)
        {
            string idsPath = args[0];
            string inputPath = args[1];
            string outputPath = args[2];


            Console.WriteLine("Start!");
            var main = FindIDS("IDS1");

            if (main == IntPtr.Zero)
            {
                Console.WriteLine("can't connect to the Ids");
                return;
            }
            Console.WriteLine("contected to Ids");

            Console.WriteLine("open his file");
            OpenHistorical(main, idsPath);


            Console.WriteLine("add text files");
            var files = Directory.GetFiles(inputPath);

            foreach (var file in files)
            {
                AddTextFile(main, file);
                // analysis -> alternative
                User32Methods.SendMessage(main, (uint)WM.COMMAND, (IntPtr)23813, IntPtr.Zero);
                SaveFile(main, Path.Combine(outputPath, Path.GetFileName(file)));
                Thread.Sleep(time);
            }
            Console.WriteLine("End!");
        }


        static void TypeInInput(IntPtr hWnd, string text)
        {
            foreach (var chr in text)
            {
                User32Methods.SendMessage(hWnd, (uint)WM.CHAR, (IntPtr)chr, IntPtr.Zero);
            }
        }

        static void ButtonClick(IntPtr hWnd)
        {
            User32Methods.SendMessage(hWnd, (uint)WM.LBUTTONDOWN, (IntPtr)MouseInputKeyStateFlags.MK_LBUTTON, IntPtr.Zero);
            User32Methods.SendMessage(hWnd, (uint)WM.LBUTTONUP, (IntPtr)MouseInputKeyStateFlags.MK_LBUTTON, IntPtr.Zero);
        }

        static IntPtr FindIDS(string name)
        {
            var windowHandle = User32Methods.FindWindow(null, appName + name);
            if (windowHandle == IntPtr.Zero)
            {
                return IntPtr.Zero;
            }

            return windowHandle;
        }

        static void OpenHistorical(IntPtr windowHandle, string idsPath)
        {
            //open ids file           
            User32Methods.PostMessage(windowHandle, (uint)WM.COMMAND, (IntPtr)57601, IntPtr.Zero);

            Thread.Sleep(time);
            var inputForm = User32Methods.FindWindow(null, "Open");

            var comboboxEx = User32Methods.FindWindowEx(inputForm, IntPtr.Zero, "ComboBoxEx32", null);
            var combobox = User32Methods.FindWindowEx(comboboxEx, IntPtr.Zero, "ComboBox", null);
            var edit = User32Methods.FindWindowEx(combobox, IntPtr.Zero, "Edit", null);
            Thread.Sleep(time);
            TypeInInput(edit, idsPath);

            var button = User32Methods.FindWindowEx(inputForm, IntPtr.Zero, "Button", "&Open");
            ButtonClick(button);
            Thread.Sleep(time);
        }

        private static void SaveFile(IntPtr main, string path)
        {
            // report -> ext -> attribute
            User32Methods.PostMessage(main, (uint)WM.COMMAND, (IntPtr)32790, IntPtr.Zero);
            Thread.Sleep(time);

            var inputForm = User32Methods.FindWindow(null, "Save As");

            var comboboxEx = User32Methods.FindWindowEx(inputForm, IntPtr.Zero, "ComboBoxEx32", null);
            var combobox = User32Methods.FindWindowEx(comboboxEx, IntPtr.Zero, "ComboBox", null);
            var edit = User32Methods.FindWindowEx(combobox, IntPtr.Zero, "Edit", null);
            Thread.Sleep(time);
            TypeInInput(edit, path);

            var button = User32Methods.FindWindowEx(inputForm, IntPtr.Zero, "Button", "&Save");
            ButtonClick(button);
            Thread.Sleep(time);

            var dialog = User32Methods.FindWindow(null, "IDS");
            button = User32Methods.FindWindowEx(dialog, IntPtr.Zero, "Button", "OK");
            ButtonClick(button);
            ButtonClick(button);

            Thread.Sleep(time);
        }

        private static void AddTextFile(IntPtr main, string path)
        {
            //open text file           
            User32Methods.PostMessage(main, (uint)WM.COMMAND, (IntPtr)32859, IntPtr.Zero);

            Thread.Sleep(time);
            var inputForm = User32Methods.FindWindow(null, "IDS Dialog: Read Data from a Text File");
            var edit = User32Methods.FindWindowEx(inputForm, IntPtr.Zero, "Edit", null);
            TypeInInput(edit, path);

            var button = User32Methods.FindWindowEx(inputForm, IntPtr.Zero, "Button", "&Read");
            ButtonClick(button);

            Thread.Sleep(time);

            var dialog = User32Methods.FindWindow(null, "IDS");
            button = User32Methods.FindWindowEx(dialog, IntPtr.Zero, "Button", "OK");
            ButtonClick(button);
            ButtonClick(button);

            Thread.Sleep(time);
        }
    }
}
