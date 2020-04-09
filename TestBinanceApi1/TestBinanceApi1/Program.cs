using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace TestBinanceApi1
{
    class Program
    {
        static string api = "https://api.binance.com";

        static void Main(string[] args)
        {
            Console.WriteLine(outputToDisplay());
            Console.ReadKey();
        }

        private static string outputToDisplay()
        {
            return getWebResponse(api + "/api/v1/exchangeInfo");
        }

        private static string getWebResponse(string url)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            // create request..
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);

            // use GET method
            webRequest.Method = "GET";

            // POST!
            HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

            // read response into StreamReader
            Stream responseStream = webResponse.GetResponseStream();
            StreamReader _responseStream = new StreamReader(responseStream);

            // get raw result
            return _responseStream.ReadToEnd();
        }
    }
}