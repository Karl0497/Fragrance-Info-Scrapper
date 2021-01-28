using System;
using System.Threading;
using System.Threading.Tasks;

namespace FragranceInfo
{
    class Program
    {
        static void Main(string[] args)
        {
            Task.Run(() => FragranticaCrawler.ProcessURLs()).Wait();     
        }
    }
}
