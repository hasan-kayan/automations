using System;
using System.Diagnostics;
using System.Threading;

class Program
{
    static void Main(string[] args)
    {
        Console.Write("Please enter the process name you would like to check: ");
        string processName = Console.ReadLine();

        Process[] processes = Process.GetProcessesByName(processName);
        bool isRunning = processes.Length > 0;

        if (isRunning)
        {
            Console.WriteLine($"{processName} program is running.");

            bool isNotResponding = !processes[0].Responding;
            bool isSuspended = processes[0].Threads[0].WaitReason == ThreadWaitReason.Suspended;

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
}
