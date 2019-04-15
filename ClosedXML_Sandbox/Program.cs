using System;

namespace ClosedXML_Sandbox
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            var i = 1;
            var count = i;
            do
            {
                Console.WriteLine("Running {0}({1})", nameof(PerformanceRunner.RunMerge), count);
                PerformanceRunner.TimeAction(() => PerformanceRunner.RunMerge(count));
                count *= 2;
                i++;
            } while (i < 12);
            //PerformanceRunner.TimeAction(() => PerformanceRunner.RunMerge(256));

            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}
