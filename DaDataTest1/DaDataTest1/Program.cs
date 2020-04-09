using System;
using Dadata;

namespace DaDataTest1
{
    class Program
    {
        static void Main(string[] args)
        {
            var token = "5fd4343f6200a3a67dbe7ec194673f883c5f7645";
            var api = new SuggestClient(token);

            var response = api.SuggestAddress("Владивосток", 1);
            var address = response.suggestions[0].data;

            Console.WriteLine("Геокоординаты: " + address.geo_lat + ", " + address.geo_lon + ".");
            Console.ReadKey();
        }
    }
}