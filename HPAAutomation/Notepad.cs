using System;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.IO;

namespace HPAAutomation
{
    /// <summary>
    /// This class handles the operations for interacting with 
    /// an instance of Notepad using win32 API calls
    /// </summary>
    static class Notepad
    {
        #region Imports

        public delegate bool EnumWindowProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, int hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, IntPtr lParam);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private static Guid FolderDownloads = new Guid("374DE290-123F-4565-9164-39C4925E467B");
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern int SHGetKnownFolderPath(ref Guid id, int flags, IntPtr token, out IntPtr path);

        #endregion

        #region Constants
        const int WM_SETTEXT = 0x000C;
        const int BM_CLICK = 0x00F5;
        const int WM_ACTIVATE = 0x0006;
        const int WA_ACTIVE = 0x0001;
        const int WM_COMMAND = 0x0111;
        #endregion

        static Process mNotepad;

        /// <summary>
        /// This method will search for an open Notepad instance.
        /// If one is not found, it will open a new Notepad instance.
        /// </summary>
        public static void OpenNotepad()
        {
            // check if notepad is open
            Console.WriteLine("1. Checking if Notepad is open...");
            var openedNotepad = Process.GetProcessesByName("notepad").FirstOrDefault();
            if (openedNotepad == null)
            {
                Console.WriteLine("2. No Notepad instance found. Opening a new instance of Notepad...");
                mNotepad = Process.Start(@"notepad.exe");
            }
            else
            {
                Console.WriteLine("2. Notepad instance found. Continuing...");
                mNotepad = openedNotepad;
            }
        }

        /// <summary>
        /// This method adds the text to the Notepad textbox.
        /// </summary>
        /// <param name="message">The message to add to the Notepad textbox</param>
        public static void AddText(string message)
        {
            // wait for the windows to open
            System.Threading.Thread.Sleep(300);

            Console.WriteLine("4. Writing \"" + message + "\" to Notepad...");
            var notepadTextBox = FindWindowEx(mNotepad.MainWindowHandle, 0, "Edit", null);
            SendMessage(notepadTextBox, WM_SETTEXT, 0, message);
        }

        /// <summary>
        /// This method click the 'New' link in the File Menu of Notepad
        /// </summary>
        public static void New()
        {
            // wait for UI
            System.Threading.Thread.Sleep(300);

            Console.WriteLine("3. Clicking \"New\" in File Menu...");
            // 0x0000 = 1st item in the file menu 'New'
            PostMessage(mNotepad.MainWindowHandle, WM_COMMAND, 0x0001, IntPtr.Zero);


            System.Threading.Thread.Sleep(300);
            // if there are unsaved changes to the document
            // when we try to click the new button, click don't save
            // and continue.
            var saveChangesWindow = FindWindow(null, "Notepad");
            if((int)saveChangesWindow != 0)
            {
                Console.WriteLine("Unsaved Changes Dialog detected...");
                var directUIWNDwindow = FindWindowEx(saveChangesWindow, 0, "DirectUIHWND", string.Empty);
                var directUIWNDHandleAllChildren = GetAllChildHandles(directUIWNDwindow);
                IntPtr doNotSaveButton = IntPtr.Zero;

                foreach (var handle in directUIWNDHandleAllChildren)
                {
                    doNotSaveButton = FindWindowEx(handle, 0, "Button", "Do&n\'t Save");
                    if ((int)doNotSaveButton != 0)
                        break;
                }

                Console.WriteLine("Clicking Don't Save...");
                // click the Don't Save button
                PostMessage(doNotSaveButton, WM_ACTIVATE, new IntPtr(WA_ACTIVE), IntPtr.Zero);
                PostMessage(doNotSaveButton, BM_CLICK, 0, new IntPtr(0));
            }
        }

        /// <summary>
        /// This method handles the saving of the Notepad.
        /// It detects if there is a Save As dialog and acts accordingly.
        /// </summary>
        public static void Save()
        {
            // wait for UI
            System.Threading.Thread.Sleep(300);

            Console.WriteLine("5. Clicking \"Save As\" in File Menu...");

            //0x0003 = the 4th item in the file menu 'Save As'
            PostMessage(mNotepad.MainWindowHandle, WM_COMMAND, 0x0003, new IntPtr(0));

            // wait on UI
            System.Threading.Thread.Sleep(300);
            Console.WriteLine("6. Changing name of file in Save As Dialog...");

            // go through and find the Edit handle
            var saveAsWindowHandle = FindWindow(null, "Save As");
            var DUIViewWndClassNameHandle = FindWindowEx(saveAsWindowHandle, 0, "DUIViewWndClassName", string.Empty);
            var DirectUIHWNDHandle = FindWindowEx(DUIViewWndClassNameHandle, 0, "DirectUIHWND", string.Empty);
            var floatNotifySinkHandle = FindWindowEx(DirectUIHWNDHandle, 0, "FloatNotifySink", string.Empty);
            var comboBoxHandle = FindWindowEx(floatNotifySinkHandle, 0, "ComboBox", string.Empty);
            var EditHandle = FindWindowEx(comboBoxHandle, 0, "Edit", string.Empty);
            SendMessage(EditHandle, WM_SETTEXT, 0, "PathAndFileNameToHelloWorld.txt");

            Console.WriteLine("Clicking the save button...");
            var saveButtonHandle = FindWindowEx(saveAsWindowHandle, 0, "Button", "&Save");
            PostMessage(saveButtonHandle, WM_ACTIVATE, new IntPtr(WA_ACTIVE), IntPtr.Zero);
            PostMessage(saveButtonHandle, BM_CLICK, 0, new IntPtr(0));

            Console.WriteLine("7. Checking if Confirm Save Dialog is present...");
            CheckConfirmDialog();
        }

        /// <summary>
        /// This method will verify whether the Confirm Save As window has popped up.
        /// If it has popped up, it will act accordingly.
        /// </summary>
        public static void CheckConfirmDialog()
        {
            // wait for UI
            System.Threading.Thread.Sleep(300);

            var confirmSaveAsWindow = FindWindow(null, "Confirm Save As");
            if ((int)confirmSaveAsWindow == 0x0)
            {
                Console.WriteLine("No Confirm Save As Dialog present. Continuing...");
                return;
            }

            Console.WriteLine("Confirm Save As Dialog present. Clicking Yes button to confirm...");
            var directUIHWNDHandle = FindWindowEx(confirmSaveAsWindow, 0, "DirectUIHWND", string.Empty);

            // DirectUIHWND has like 8 different children all with the same name.
            // only 2 of those 8 have buttons, and 1 of those is the yes button.
            // the hex values of the children change every time the program is ran,
            // so no hardcoding the yes button's hex value to click it.
            // i have to iterate through those 8 to find the yes button.
            var directUIWNDHandleAllChildren = GetAllChildHandles(directUIHWNDHandle);
            IntPtr yesButton = IntPtr.Zero;

            foreach (var handle in directUIWNDHandleAllChildren)
            {
                yesButton = FindWindowEx(handle, 0, "Button", "&Yes");
                if ((int)yesButton != 0)
                    break;
            }

            // click the Yes button
            PostMessage(yesButton, WM_ACTIVATE, new IntPtr(WA_ACTIVE), IntPtr.Zero);
            PostMessage(yesButton, BM_CLICK, 0, new IntPtr(0));
        }

        /// <summary>
        /// This method will verify that the operations taken place by this program were fruitful.
        /// It will look in the Documents folder for the saved Notepad file that this program has done.
        /// </summary>
        /// <param name="fileName">The name of the file to search for.</param>
        public static void VerifyFileHasBeenStored(string fileName)
        {
            List<string> paths = new()
            {
                System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                GetDownloadsPath(),
                System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            foreach (string path in paths)
            {
                Console.WriteLine("Searching " + path + " for the file...");
                if (File.Exists(path + "\\" + fileName))
                {
                    Console.WriteLine("File found at " + path);
                    return;
                }
            }

            Console.Write("File not found in Desktop, Document, or Downloads folder...");
        }

        /// <summary>
        /// Since the downloads folder is not located in the environment.specialfolder data object,
        /// we have to do some extra stuff to get the generalized path for the downloads folder.
        /// </summary>
        /// <returns>The path to the Downloads folder.</returns>
        public static string GetDownloadsPath()
        {
            if (Environment.OSVersion.Version.Major < 6) throw new NotSupportedException();
            IntPtr pathPtr = IntPtr.Zero;
            try
            {
                SHGetKnownFolderPath(ref FolderDownloads, 0, IntPtr.Zero, out pathPtr);
                return Marshal.PtrToStringUni(pathPtr);
            }
            finally
            {
                Marshal.FreeCoTaskMem(pathPtr);
            }
        }

        /// <summary>
        /// This method gets all the child handles for a given handle.
        /// This is useful when a parent handle has many children handles
        /// with the same name.
        /// </summary>
        /// <param name="handle">The window handle</param>
        /// <returns></returns>
        public static List<IntPtr> GetAllChildHandles(IntPtr handle)
        {
            List<IntPtr> childHandles = new List<IntPtr>();

            GCHandle gcChildhandlesList = GCHandle.Alloc(childHandles);
            IntPtr pointerChildHandlesList = GCHandle.ToIntPtr(gcChildhandlesList);

            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(handle, childProc, pointerChildHandlesList);
            }
            finally
            {
                gcChildhandlesList.Free();
            }

            return childHandles;
        }

        /// <summary>
        /// This enumerates the specified window and uses a 
        /// callback function in order to add all the children handles
        /// </summary>
        /// <param name="hWnd">The child window handle</param>
        /// <param name="lParam">Size</param>
        /// <returns></returns>
        private static bool EnumWindow(IntPtr hWnd, IntPtr lParam)
        {
            GCHandle gcChildhandlesList = GCHandle.FromIntPtr(lParam);

            if (gcChildhandlesList == null || gcChildhandlesList.Target == null)
            {
                return false;
            }

            List<IntPtr> childHandles = gcChildhandlesList.Target as List<IntPtr>;
            childHandles.Add(hWnd);

            return true;
        }
    }
}
