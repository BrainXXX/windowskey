using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Net;
using System.IO;
using System.Globalization;

namespace TestBTC
{
    class Program
    {
        static void Main(string[] args)
        {
            string Coin1 = "BTC-NEO";
            string Coin2 = "USDT-NEO";
            string Coin3 = "USDT-BTC";
            string Coin4 = "BTC-OMG";

            // display
            outputToDisplayPrice(Coin1);
            outputToDisplayPrice(Coin2);
            outputToDisplayPrice(Coin3);
            outputToDisplayPrice(Coin4);
            outputToDisplayVolume(Coin1);
            outputToDisplayVolume(Coin4);
            Console.WriteLine("40 NEO =  {0:F2} $ " + 40 * getCoinInfo(Coin1), (40 * getCoinInfo(Coin2)));
            Console.WriteLine("190 OMG = {0:F2} $ " + 190 * getCoinInfo(Coin4), (190 * getCoinInfo(Coin4) * getCoinInfo(Coin3)));
            //Console.WriteLine("");
            //Console.WriteLine(getWebResponse());
            //File.WriteAllText("output.txt", getWebResponse());
            File.WriteAllLines("output.txt", getWebResponse());
            Console.ReadKey();
        }

        private static void outputToDisplayPrice(string coin)
        {
            Console.WriteLine(String.Format(coin + "\t{0} price", getCoinInfo(coin)));
        }

        private static void outputToDisplayVolume(string coin)
        {
            Console.WriteLine(String.Format(coin + "\t{0:F2} volume", getCoinVolume(coin)));
        }

        private static decimal getCoinInfo(string coin)
        {
            string byc = getWebResponse("https://bittrex.com/api/v1.1/public/getticker?market=" + coin);

            int pos1 = byc.IndexOf("Last", 0);
            int pos2 = byc.IndexOf(":", pos1);
            int pos3 = byc.IndexOf("}", pos2);

            return Convert.ToDecimal(byc.Substring(pos2 + 1, pos3 - pos2 - 1), CultureInfo.InvariantCulture);
        }

        private static decimal getCoinVolume(string coin)
        {
            string byc = getWebResponse("https://bittrex.com/api/v1.1/public/getmarketsummary?market=" + coin);

            int pos1 = byc.IndexOf("BaseVolume", 0);
            int pos2 = byc.IndexOf(":", pos1);
            int pos3 = byc.IndexOf(",", pos2);

            return Convert.ToDecimal(byc.Substring(pos2 + 1, pos3 - pos2 - 1), CultureInfo.InvariantCulture);
        }

        private static string[] getWebResponse()
        {
            string byc = getWebResponse("https://bittrex.com/api/v1.1/public/getmarketsummaries");
            //string byc = getWebResponse("https://bittrex.com/api/v1.1/public/getcurrencies");
            //string byc = getWebResponse("https://bittrex.com/api/v1.1/public/getmarkets");
            //string byc = getWebResponse("https://bittrex.com/api/v1.1/public/getorderbook?market=BTC-LTC&type=both");
            //string byc = getWebResponse("https://bittrex.com/api/v1.1/public/getmarkethistory?market=BTC-DOGE");

            int pos1 = byc.IndexOf(":[", 0);
            int pos2 = byc.IndexOf("[", pos1);
            int pos3 = byc.IndexOf("]", pos2);

            string[] str = byc.Substring(pos2 + 1, pos3 - pos2 - 1).Split('}');
            string[] str2 = new string[str.Length];

            for (int i = 0, b = 0; i < str.Length; i++)
            {
                if (str[i].Contains("BTC-"))
                {
                    str2[b] = str[i].Replace(",{", "").Replace("{", "");
                    b++;
                }
            }

            return str2 = str2.Where(x => x != null).ToArray();
        }

        public static string getWebResponse(string url)
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