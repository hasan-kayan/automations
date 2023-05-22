using System;
using System.Diagnostics;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        Console.Write("Please enter the process name you would like to check: ");
        string processName = Console.ReadLine();

        Process[] processes = Process.GetProcessesByName(processName);

        if (processes.Length > 0)
        {
            Process targetProcess = processes.First();

            Console.WriteLine($"{processName} program is running.");

            if (!targetProcess.Responding)
            {
                Console.WriteLine($"{processName} is not responding.");
            }
            else
            {
                Console.WriteLine($"{processName} is responding.");
            }

            if (targetProcess.Threads.Cast<System.Diagnostics.ProcessThread>().Any(thread => thread.ThreadState == System.Diagnostics.ThreadState.Wait || thread.ThreadState == System.Diagnostics.ThreadState.WaitSleepJoin))
            {
                Console.WriteLine($"{processName} is in a waiting state.");
            }
            else
            {
                Console.WriteLine($"{processName} is not in a waiting state.");
            }

            if (targetProcess.Threads.Cast<System.Diagnostics.ProcessThread>().Any(thread => thread.ThreadState == System.Diagnostics.ThreadState.Suspended))
            {
                Console.WriteLine($"{processName} is suspended.");
            }
            else
            {
                Console.WriteLine($"{processName} is not suspended.");
            }
        }
        else
        {
            Console.WriteLine($"{processName} is not running.");
        }
    }
}
