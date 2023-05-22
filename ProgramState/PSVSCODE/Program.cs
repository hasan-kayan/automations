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
            bool isWaiting = false;

           if( processes[0].Threads[0].ThreadState = System.Threading.ThreadState.Suspended)
            {
                isSuspended = true;
            }

            foreach (ProcessThread thread in processes[0].Threads)
            {
                if (thread.ThreadState == System.Threading.ThreadState.WaitSleepJoin)
                {
                    isWaiting = true;
                }

                if (thread.ThreadState == System.Threading.ThreadState.Suspended)
                {
                    isSuspended = true;
                    break;
                }
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

            if (isWaiting)
            {
                Console.WriteLine($"{processName} is in a waiting state.");
            }
            else
            {
                Console.WriteLine($"{processName} is not in a waiting state.");
            }
        }
        else
        {
            Console.WriteLine($"{processName} is not running.");
        }
    }

    static bool IsProcessRunning(string processName)
    {
        Process[] processes = Process.GetProcessesByName(processName);
        return processes.Length > 0;
    }
}



