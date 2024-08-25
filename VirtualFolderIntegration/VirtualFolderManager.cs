using Microsoft.Win32;
using Shell32;

namespace VirtualFolderIntegration
{
    public class VirtualFolderManager
    {
        static void Main(string[] args)
        {
            CreateShortcut();
            CreateRegistryEntry();
        }

        public void CreateShortcut(string shortcutPath, string targetPath, string description)
        {
            Shell shell = new Shell();
            Folder folder = shell.NameSpace(Path.GetDirectoryName(shortcutPath));
            FolderItem folderItem = folder.ParseName(Path.GetFileName(shortcutPath));

            ShellLinkObject shortcut = (ShellLinkObject)folderItem.GetLink;
            shortcut.Path = targetPath;
            shortcut.Description = description;
            shortcut.Save();
        }

        public void AddToThisPC(string folderName, string executablePath)
        {
            string clsid = Guid.NewGuid().ToString("B"); // Generate a new GUID
            string keyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\" + clsid;

            using (RegistryKey key = Registry.LocalMachine.CreateSubKey(keyName))
            {
                if (key != null)
                {
                    key.SetValue(null, folderName);
                    key.SetValue("Target", executablePath);
                }
            }
        }

        public void CreateShortcutInQuickAccess(string targetPath)
        {
            string quickAccessPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                                  @"Microsoft\Windows\Recent\CustomDestinations");
            string shortcutPath = Path.Combine(quickAccessPath, "MyVirtualFolder.lnk");

            Directory.CreateDirectory(quickAccessPath);

            Shell shell = new Shell();
            Folder folder = shell.NameSpace(quickAccessPath);
            FolderItem folderItem = folder.ParseName("MyVirtualFolder");

            ShellLinkObject shortcut = (ShellLinkObject)folderItem.GetLink;
            shortcut.Path = targetPath;
            shortcut.Save(shortcutPath);
        }

        public void AddToThisPC(string folderName)
        {
            string clsid = Guid.NewGuid().ToString("B"); // Generate a new GUID
            string keyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\" + clsid;

            using (RegistryKey key = Registry.LocalMachine.CreateSubKey(keyName))
            {
                if (key != null)
                {
                    key.SetValue(null, folderName);
                }
            }
        }

        static void CreateShortcut()
        {
            string quickAccessPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                                  @"Microsoft\Windows\Recent\CustomDestinations");
            string shortcutPath = Path.Combine(quickAccessPath, "MyVirtualFolder.lnk");

            string targetPath = @"C:\Path\To\Your\VirtualFolder";

            // Ensure the directory exists
            Directory.CreateDirectory(quickAccessPath);

            // Create a Shell object
            Shell shell = new Shell();
            Folder folder = shell.NameSpace(quickAccessPath);

            // Create the shortcut
            FolderItem folderItem = folder.ParseName("ElectroShare");
            ShellLinkObject shortcut = (ShellLinkObject)folderItem.GetLink;

            shortcut.Path = targetPath;
            shortcut.Save(shortcutPath);

            Console.WriteLine("Shortcut created successfully!");
        }

        static void CreateRegistryEntry()
        {
            string clsid = Guid.NewGuid().ToString("B"); // Generate a new GUID for your virtual folder
            string keyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\" + clsid;

            using (RegistryKey key = Registry.LocalMachine.CreateSubKey(keyName))
            {
                if (key != null)
                {
                    key.SetValue(null, "Electro Share");
                    Console.WriteLine("Registry entry created successfully!");
                }
            }
        }
    }
}
