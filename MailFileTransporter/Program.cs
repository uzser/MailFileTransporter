using System;

namespace MailFileTransporter
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            var mailProcessor =new MailProcessor();

            if (args.Length == 1)
                mailProcessor.SendFilesFromFolder(args[0]);
            else if (args.Length > 1)
            {
                Console.WriteLine("Program takes only one argument.");
            }
            else
            {
                mailProcessor.SetupToReceive();
            }
            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();
        }
    }
}