using Newtonsoft.Json;
using System;

namespace SoftDogStandingBookNetCore
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                Console.WriteLine("Arguments error: must be 3 args");
                return ;
            }
            string option = args[0];
            if (option != "word" && option != "excel")
            {
                Console.WriteLine("Option error: must be one of 'word' or 'excel'");
                return ;
            }
            string input = args[1]; //json
            string output = AppDomain.CurrentDomain.BaseDirectory + "output\\" + args[2];    //outpudata
            if (option == "word")
            {
                var data = JsonConvert.DeserializeObject<Datas>(input);
                Functions.ExportWord(output, data);
            }
            else
            {
                //var data = JsonConvert.DeserializeObject<IEnumerable<Datas>>(input);
                //Functions.ExportExcel(output, args[1], args[2]);
            }
            return ;
        }
    }
}
