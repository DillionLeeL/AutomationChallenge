/// Written by Dillion Lowry

using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Threading;
using System.IO;

namespace AutomationChallenge
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                /// Open Notepad
                Process np;
                int sleepTime = 1000;

                var notepad = Process.GetProcessesByName("notepad");
                if (notepad.Length == 0)
                {
                    Console.WriteLine("Opening notepad");
                    np = Process.Start("notepad.exe");
                    Thread.Sleep(sleepTime);
                }
                else
                {
                    Console.WriteLine("Notepad is already open");
                    np = notepad.FirstOrDefault();
                }
                AutomationElement window = AutomationElement.FromHandle(np.MainWindowHandle);
                window.SetFocus();
                Thread.Sleep(1000);


                //Open a new document
                Console.WriteLine("Opening File menu");
                AutomationElement fileMenu = OpenMenuByName(window, "File");
                Thread.Sleep(sleepTime);

                Console.WriteLine("Creating new file");
                PressButtonByName(fileMenu, "New");
                Thread.Sleep(sleepTime);


                //If there was an open notepad already, exit it
                AutomationElement exitExisting = window.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Notepad"));
                if (exitExisting != null)
                {
                    Console.WriteLine("Closing your old file. Hopefully you didn't need that");
                    PressButtonByName(exitExisting, "Don't Save");

                }
                Thread.Sleep(sleepTime);

                Console.WriteLine("Writing to file");
                SendText(window, "Hello World");
                Thread.Sleep(sleepTime);

                Console.WriteLine("Opening File menu");
                fileMenu = OpenMenuByName(window, "File");
                Thread.Sleep(sleepTime);

                Console.WriteLine("Selecting 'Save As'");
                PressButtonByName(fileMenu, "Save As...");
                Thread.Sleep(2000);


                /// Detect save as window, write filename, confirm
                AutomationElement saveAsWindow = window.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Save As"));

                if (saveAsWindow == null)
                {
                    Console.WriteLine("There was a problem saving the file. Exiting.");
                    Environment.Exit(1);
                }
                Console.WriteLine("Changing Filename");
                SendText(saveAsWindow,"PathAndFileNameToHelloWorld.txt");

                Console.WriteLine("Changing location to Downloads folder");
                MakeSelectionFromList(saveAsWindow, "Downloads");
                Thread.Sleep(sleepTime);

                Console.WriteLine("Saving");
                PressButtonByName(saveAsWindow, "Save");
                Thread.Sleep(sleepTime);


                /// Confirmation window popup
                AutomationElement saveAsConfirm = saveAsWindow.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Confirm Save As"));
                if (saveAsConfirm != null)
                {
                    PressButtonByName(saveAsConfirm, "Yes");
                    Console.WriteLine("Confirming save");
                }


                /// Confirm that file was saved
                /// Downloads folder is not in the enumeration Windows.Storage.UserDataPaths so you can't use that
                /// The below solution only works on the default download path
                /// Syroot.Windows.IO.KnownFolders nuget package can be used to find the actual path using the Win32 KNOWNFOLDERID
                string path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads";
                Console.WriteLine("Looking for Downloads folder");
                if (!Directory.Exists(path))
                    Console.WriteLine("Couldn't find Downloads Folder");
                else
                {
                    Console.WriteLine("Checking Downloads Folder for file");
                    Console.WriteLine(path + "\\PathAndFileNameToHelloWorld.txt");
                    Console.WriteLine(File.Exists(path+ "\\PathAndFileNameToHelloWorld.txt") ? "File exists.\nTask Completed Sucessfully" : "File does not exist.");
                }
                
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            Console.ReadLine();
        }

        public static AutomationElement OpenMenuByName(AutomationElement window, string name)
        {
            try
            {
                AutomationElement menuItem = window.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name));
                ExpandCollapsePattern menuToggle = (ExpandCollapsePattern)menuItem.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                menuToggle.Expand();
                return menuItem;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        
        public static void PressButtonByName(AutomationElement menu, string name)
        {
            try
            {
                AutomationElement selection = menu.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name));
                InvokePattern buttonPress = (InvokePattern)selection.GetCurrentPattern(InvokePattern.Pattern);
                // Run invoke on new thread because of blocking dialog box causing poor performance
                new Thread(() => buttonPress.Invoke()).Start();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }


        public static void SendText(AutomationElement element, string text)
        {
            try
            { 
                element.SetFocus();

                if (element == null || text == null)
                    throw new ArgumentNullException("Null input or window");
                if (!element.Current.IsEnabled || !element.Current.IsKeyboardFocusable)
                    throw new InvalidOperationException("Couldn't find or write to window");

                SendKeys.SendWait(text);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        public static void MakeSelectionFromList(AutomationElement window, string name)
        {
            try
            {
                AutomationElement selectionMade = window.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name));
                SelectionItemPattern selDown = (SelectionItemPattern)selectionMade.GetCurrentPattern(SelectionItemPattern.Pattern);
                selDown.Select();
                Thread.Sleep(50);
                SendKeys.SendWait("{ENTER}");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

    }
}