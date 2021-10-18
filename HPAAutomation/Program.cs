using System;

namespace HPAAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            Notepad.OpenNotepad();
            Notepad.New();
            Notepad.AddText("Hello world");
            Notepad.Save();
            Notepad.VerifyFileHasBeenStored("PathAndFileNameToHelloWorld.txt");
            Console.WriteLine("Task to automate editing and saving Notepad has been completed.");
            Console.WriteLine("Press any key to end the program.");
            Console.Read();
        }
    }
}
