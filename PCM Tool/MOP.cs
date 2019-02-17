using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PMC_Tool
{
    public class MOP
    {
        public Item[] item;
        public MOP(int number)
        {
            item = new Item[number];
        }

        public struct Item
        {
            public string name;
            public double mopqty;
        }
    }
}
