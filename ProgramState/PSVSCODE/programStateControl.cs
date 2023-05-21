using System;
using System.Diagnostics;
using System.Threading;

class Program
{
    static void Main(string[] args)
    {
        Console.Write("Please enter the process name you would like to check : ");
        string processName = Console.ReadLine();
        bool isRunning = IsProcessRunning(processName);

        if (isRunning)
        {
            Console.WriteLine($"{processName} program works");

            Process[] processes = Process.GetProcessesByName(processName);
            bool isNotResponding = !processes[0].Responding;
            bool isSuspended = false;

            foreach (ProcessThread thread in processes[0].Threads)
            {
                if (thread.ThreadState == ThreadState.Wait || thread.ThreadState == ThreadState.Suspended)
                {
                    isSuspended = true;
                    break;
                }
            }

            if (isNotResponding)
            {
                Console.WriteLine($"{processName} not responding.");
            }
            else
            {
                Console.WriteLine($"{processName} responding.");
            }

            if (isSuspended)
            {
                Console.WriteLine($"{processName} suspended.");
            }
            else
            {
                Console.WriteLine($"{processName} not suspended.");
            }
        }
        else
        {
            Console.WriteLine($"{processName} not working.");
        }
    }

    static bool IsProcessRunning(string processName)
    {
        Process[] processes = Process.GetProcessesByName(processName);
        return processes.Length > 0;
    }
}
