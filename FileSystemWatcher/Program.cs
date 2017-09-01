using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.IO;
using Spire.Xls;
using System.Threading;

namespace FileSystemWatcher
{
    class Program
    {
        void FileWatcherprogram()
        {
            System.IO.FileSystemWatcher fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            // fileSystemWatcher1.EnableRaisingEvents = true;
            // Control.CheckForIllegalCrossThreadCalls = false;
            fileSystemWatcher1.Path = @"D:\ExcelFileWatcherInput\";
            fileSystemWatcher1.IncludeSubdirectories = true;
            fileSystemWatcher1.NotifyFilter |= NotifyFilters.Attributes | NotifyFilters.CreationTime | NotifyFilters.LastAccess | NotifyFilters.Size;

            while (true)
            {
                WaitForChangedResult w = fileSystemWatcher1.WaitForChanged(WatcherChangeTypes.All);
                // label1.Text = w.Name + " " + w.ChangeType;
                // go in action whenever new file is created
                if (w.ChangeType == WatcherChangeTypes.Created)
                {

                    // I used D drive for demo purpose. You can change it or enahnce code as per your need
                    string filePath = @"D\ExcelFileWatcherInput\InputActivity.txt";
                    StreamWriter sw = new StreamWriter(filePath, true);
                    sw.WriteLine("\n New file created at " + DateTime.Now.ToString());
                    sw.Close();

                    // check file is excel file or not
                    if (Path.GetExtension(fileSystemWatcher1.Path + w.Name).ToString().ToLower() == ".xls" || Path.GetExtension(fileSystemWatcher1.Path + w.Name).ToString().ToLower() == ".xlsx")
                    {
                        // workbook is class of Spire.xls nuget package. free and open source tools
                        Workbook workbook = new Workbook();
                        workbook.LoadFromFile(fileSystemWatcher1.Path + "\\" + w.Name);

                        //convert Excel to HTML 
                        Worksheet sheet = workbook.Worksheets[0];
                        string outputpath = @"D:\ExcelFileWatcherOutput\";
                        try
                        {
                            if (File.Exists(outputpath + Path.GetFileNameWithoutExtension(fileSystemWatcher1.Path + "\\" + w.Name) + ".html"))
                            {
                                File.Delete(outputpath + Path.GetFileNameWithoutExtension(fileSystemWatcher1.Path + "\\" + w.Name) + ".html");
                            }
                            sheet.SaveToHtml(outputpath + Path.GetFileNameWithoutExtension(fileSystemWatcher1.Path + "\\" + w.Name) + ".html");
                        }
                        catch (Exception)
                        {

                        }

                       // System.Diagnostics.Process.Start(outputpath + Path.GetFileNameWithoutExtension(fileSystemWatcher1.Path + "\\" + w.Name) + ".html");
                    }
                }
            }
        }
        static void Main()
        {
            Program obj = new FileSystemWatcher.Program();
            Thread t = new Thread(obj.FileWatcherprogram);
            t.IsBackground = true;
            t.Start();
            Console.ReadLine();

        }
    }
}
