using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace test
{
    public class SportsCar : Car
    {
        public string GetPetName()
        {
            petName = "Fred";
            return petName;
        }
    }
}