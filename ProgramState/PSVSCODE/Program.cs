using System;
using System.Diagnostics;
using System.Threading;

class Program
{
    static void Main(string[] args)
    {
        Console.Write("Please enter the process name you would like to check: ");
        string processName = Console.ReadLine();
        bool isRunning = IsProcessRunning(processName);

        if (isRunning)
        {
            Console.WriteLine($"{processName} program is running.");

            Process[] processes = Process.GetProcessesByName(processName);
            bool isNotResponding = !processes[0].Responding;
            bool isSuspended = false;
            
           if( processes[0].Threads[0].WaitReason == ThreadWaitReason.Suspended)  // is waiting kaldır
            {
                isSuspended = true;
            }

           
            if (isNotResponding)
            {
                Console.WriteLine($"{processName} is not responding.");
            }
            else
            {
                Console.WriteLine($"{processName} is responding.");
            }

            if (isSuspended)
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

    static bool IsProcessRunning(string processName) // Maine taşı 
    {
        Process[] processes = Process.GetProcessesByName(processName);
        return processes.Length > 0;
    }
}



