using IWshRuntimeLibrary;
using Setup.Properties;
using System;
using System.IO.Compression;
using File = System.IO.File;
using Directory = System.IO.Directory;
using System.Diagnostics;
using System.Threading;
using Microsoft.Win32;
namespace Setupp
{
    internal class Program
    {
        /* Variables */
        static string Path = $"{Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}/DiscordProtector";

        /* Print header */
        static void PrintHeader()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Discord Protector Setup v1.0.3\n");
            Console.ResetColor();
        }

        /* Main entry point */
        static void Main(string[]args)
        {
            Console.Title = "Discord Protector Setup";
            Console.Clear();
            PrintHeader();
            while (true)
            {
                Console.CursorVisible = false;
                HomePage();
            };
        }

        /* Show home page */
        static void HomePage()
        {
            var CurrentChoice = 1;
            Console.Clear();
            PrintHeader();
            Console.WriteLine("Use the arrow keys to change your choice and press enter to confirm your choice\n");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("1) Install");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("2) Uninstall");

            /* Read user input */
            while (true)
            {
                /* Read key */
                var Key = Console.ReadKey(true).Key;

                /* Change current choice */
                if ((Key == ConsoleKey.UpArrow) && (CurrentChoice > 1))
                {
                    CurrentChoice--;
                };
                if ((Key == ConsoleKey.DownArrow) && (CurrentChoice < 2))
                {
                    CurrentChoice++;
                };

                /* Check if cofirm */
                if (Key == ConsoleKey.Enter)
                {
                    switch (CurrentChoice)
                    {
                        case 1:
                            Install();
                            break;
                        case 2:
                            Uninstall();
                            break;
                    };
                    break;
                };

                /* Re print */
                Console.CursorLeft = 0;
                Console.CursorTop = 4;
                switch (CurrentChoice)
                {
                    case 1:
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("1) Install");
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("2) Uninstall");
                        break;
                    case 2:
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("1) Install");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("2) Uninstall");
                        break;
                };
            };
        }

        /* Install */
        static void Install()
        {
            Console.Clear();
            PrintHeader();
            Console.WriteLine("Installing Discord Protector\n");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("Killing Discord");
            foreach(var p in Process.GetProcessesByName("discord"))
            {
                p.Kill();
            };
            foreach(var p in Process.GetProcessesByName("discordptb"))
            {
                p.Kill();
            };
            foreach(var p in Process.GetProcessesByName("discordcanary"))
            {
                p.Kill();
            };
            foreach(var p in Process.GetProcessesByName("discorddevelopment"))
            {
                p.Kill();
            };
            Thread.Sleep(5000);
            Console.WriteLine("Creating Directory");
            try
            {
                if (Directory.Exists(Path))
                {
                    Directory.Delete(Path, true);
                };
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Failed to install Discord Protector, Try closing any open instances of Discord Protector or try restarting your pc.", "Discord Protector Setup", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(1);
                return;
            };
            Directory.CreateDirectory(Path);
            Console.WriteLine("Writing DiscordProtector.exe");
            File.WriteAllBytes($"{Path}/DiscordProtector.exe",Resources.DiscordProtector);
            Console.WriteLine("Writing Discord Protector API.exe");
            File.WriteAllBytes($"{Path}/api.exe",Resources.api);
            Console.WriteLine("Writing Resources");
            File.WriteAllBytes($"{Path}/Resources.tmp",Resources.DPResources);
            Console.WriteLine("Extracting Resources");
            try
            {
                ZipFile.ExtractToDirectory($"{Path}/Resources.tmp",$"{Path}");
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Failed to Extract Resources, Try closing any open instances of Discord Protector or try restarting your pc.", "Discord Protector Setup",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(1);
                return;
            };
            Console.WriteLine("Cleaning up");
            File.Delete($"{Path}/Resources.tmp");
            Console.WriteLine("Creating shortcuts");
            CreateShortcut("Discord Protector",Environment.GetFolderPath(Environment.SpecialFolder.Desktop),$"{Path}/DiscordProtector.exe");
            CreateShortcut("Discord Protector",Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),$"{Path}/DiscordProtector.exe");
            Console.WriteLine("Removing registry keys");
            try
            {
                var key = Registry.CurrentUser.CreateSubKey("DiscordProtector",true);
                key.SetValue("Installed",1);
            }catch{};
            Console.WriteLine("Launching Discord Protector");
            Process.Start($"{Path}/DiscordProtector.exe");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Successfully installed Discord Protector");
            Thread.Sleep(5000);
            Environment.Exit(0);
        }

        /* Uninstall */
        static void Uninstall()
        {
            Console.Clear();
            PrintHeader();
            Console.WriteLine("Uninstalling Discord Protector\n");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("WARNING: You SHOULD uninstall Discord Protector from all your Discord clients BEFORE uninstalling Discord Protector. If you have done so, type \"yes\" and press enter to continue");
            if (Console.ReadLine().ToLower().Contains("yes"))
            {
                try
                {
                    Console.WriteLine("Killing Discord");
                    foreach (var p in Process.GetProcessesByName("discord"))
                    {
                        p.Kill();
                    };
                    foreach (var p in Process.GetProcessesByName("discordptb"))
                    {
                        p.Kill();
                    };
                    foreach (var p in Process.GetProcessesByName("discordcanary"))
                    {
                        p.Kill();
                    };
                    foreach (var p in Process.GetProcessesByName("discorddevelopment"))
                    {
                        p.Kill();
                    };
                    Thread.Sleep(5000);
                    Console.WriteLine("Removing directory");
                    if (Directory.Exists(Path))
                    {
                        Directory.Delete(Path, true);
                    };
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("Failed to uninstall Discord Protector, Try closing any open instances of Discord Protector or try restarting your pc.", "Discord Protector Setup", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    Environment.Exit(1);
                    return;
                };
                try
                {
                    Console.WriteLine("Removing shortcuts");
                    if (File.Exists($"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}/Discord Protector.lnk"))
                    {
                        File.Delete($"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}/Discord Protector.lnk");
                    };
                    if (File.Exists($"{Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)}/Discord Protector.lnk"))
                    {
                        File.Delete($"{Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)}/Discord Protector.lnk");
                    };
                }
                catch
                {
                };
                Console.WriteLine("Removing registry keys");
                try
                {
                    Registry.CurrentUser.DeleteSubKey("DiscordProtector",false);
                }catch{};
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Successfully uninstalled Discord Protector");
            }
            else
            {
                Console.WriteLine("Cancelled");
            };
            Thread.Sleep(5000);
            Environment.Exit(0);
        }

        /* Create shorcut */
        static void CreateShortcut(string Name,string SPath,string FPath)
        {
            string LPath = System.IO.Path.Combine(SPath,$"{Name}.lnk");
            WshShell shell = new WshShell();
            IWshShortcut shortcut = shell.CreateShortcut(LPath);
            shortcut.Description = "Discord Protector";
            shortcut.TargetPath = FPath;
            shortcut.Save();
        }
    };
};