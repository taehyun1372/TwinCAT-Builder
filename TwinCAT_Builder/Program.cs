using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE80;
using System.IO;
using TCatSysManagerLib;
using System.Threading;
using TwinCAT.Ads;
using NDesk.Options;

namespace ActivatePreviousConfiguration
{
    
    class Program
    {
        private static string solutionPath = null;
        private static string solutionName = null;
        private static string libraryName = null;
        private static string projectTreeName = "TIPC^Main^Main Project";
        private static string ADSpath = @"192.168.3.210.1.1";
        private static bool suppressUI = false;
        private static bool showHelp = false;

        private const int DEFAULT_WAITING_TIME = 10000;

        [STAThread]
        static void Main(string[] args)
        {
            OptionSet options = new OptionSet()
            .Add("v=|SolutionFilePath=", "path to the TwinCAT solution", v => solutionPath = v)
            .Add("n=|SolutionName=", "TwinCAT solution name", n => solutionName = n)
            .Add("l=|LibraryName=", "Library name to export", l => libraryName = l)
            .Add("p=|ProjectTreeName=", "Project tree name", p => projectTreeName = p)
            .Add("a=|ADSPath=", "ADSPath", a => ADSpath = a)
            .Add("s=|SuppressUI=", "Suppress UI", s => suppressUI = (s != null))
            .Add("?|h|help", "Help message", h => showHelp = (h != null));

            try
            {
                options.Parse(args);
            }
            catch (OptionException e)
            {
                Console.WriteLine(e.Message);
                options.WriteOptionDescriptions(Console.Out);
                Environment.Exit(0);
            }

            if (showHelp)
            {
                options.WriteOptionDescriptions(Console.Out);
                Environment.Exit(0);
            }

            if (solutionName == null || solutionName == null)
            {
                Console.WriteLine("Solution path is incorrect");
                options.WriteOptionDescriptions(Console.Out);
                Environment.Exit(0);
            }

            if (libraryName == null)
            {
                var index = solutionName.IndexOf('.');
                if(index == -1 )
                {
                    Console.WriteLine("We cannot identify extension from the given solution name");
                    options.WriteOptionDescriptions(Console.Out);
                    Environment.Exit(0);
                }
                libraryName = solutionName.Remove(index) + ".Library";
            }

            System.Type t = System.Type.GetTypeFromProgID(
                "TcXaeShell.DTE.15.0", true);
            // Create a new instance of the IDE.
            object obj = System.Activator.CreateInstance(t, true);
            // Cast the instance to DTE2 and assign to variable dte.
            EnvDTE80.DTE2 dte = (EnvDTE80.DTE2)obj;
            // Show IDE Main Window
            dte.SuppressUI = suppressUI;
            dte.MainWindow.Visible = !suppressUI;

            Console.WriteLine("Opening TwinCAT solution..");
            Console.WriteLine($"Solution path : {Path.Combine(solutionPath, solutionName)}");
            EnvDTE.Solution sol = dte.Solution;

            try
            {
                sol.Open(Path.Combine(solutionPath, solutionName));
            }
            catch (Exception e)
            {
                CleanSolution(dte, sol);
                throw new Exception(e.Message);
            }

            Console.WriteLine("TwinCAT solution successfully opened");
            Console.WriteLine($"Waiting for delay time {DEFAULT_WAITING_TIME/1000} seconds..");
            Thread.Sleep(DEFAULT_WAITING_TIME);

            try
            {
                EnvDTE.Project pro = sol.Projects.Item(1);
                Console.WriteLine("Project detected..");
                Console.WriteLine($"Project name : {pro.FullName}");

                ITcSysManager10 sysMan = (ITcSysManager10)pro.Object;

                TcConfigManager configManager =  sysMan.ConfigurationManager;

                bool foundTargetConfiguration = false;
                foreach (EnvDTE80.SolutionConfiguration2 config in dte.Solution.SolutionBuild.SolutionConfigurations)
                {
                    if (config.Name == "Release" && config.PlatformName == "TwinCAT CE7 (ARMV7)")
                    {
                        foundTargetConfiguration = true;
                        Console.WriteLine("Found the target configuration..");
                        config.Activate();
                    }
                }

                if (!foundTargetConfiguration)
                {
                    Console.Write("Could not find the target configuration");
                    CleanSolution(dte, sol);
                }

                Console.WriteLine("Setting ADS Id..");
                Console.WriteLine($"ADS Path : {ADSpath}");
                sysMan.SetTargetNetId(ADSpath);
                
                Console.WriteLine("Building the project");
                sysMan.ActivateConfiguration();

                //Library Operation
                ITcSmTreeItem plcProject = sysMan.LookupTreeItem(projectTreeName);
                
                ITcPlcIECProject importExport = (ITcPlcIECProject)plcProject;
                importExport.SaveAsLibrary(Path.Combine(solutionPath, libraryName));

                Console.WriteLine("Building was success");
                Console.WriteLine("Downloading application..");
                sysMan.StartRestartTwinCAT();
            }
            catch (Exception e)
            {
                CleanSolution(dte, sol);
                throw new Exception(e.Message);
            }

            Console.WriteLine("Downloaded successfully..");

            CleanSolution(dte, sol);

            Console.WriteLine("Exiting the programme..");
        }
        static void CleanSolution(EnvDTE80.DTE2 dte, EnvDTE.Solution sol)
        {
            Console.WriteLine("Clean the solution..");
            sol.Close();

            // Quit the TwinCAT IDE
            dte.Quit();

            // Release COM object
            System.Runtime.InteropServices.Marshal.ReleaseComObject(dte);
            dte = null;

            // Force garbage collection (optional but good for COM interop)
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}