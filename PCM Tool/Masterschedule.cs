using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PCM_Tool
{
    public class Masterschedule
    {
        public Item[] item;
        public Masterschedule(int number)
        {
            item = new Item[number];
        }

        public struct Item
        {
            public string product;
            public int quantity;
            public string status;
            public bool found;
        }
    }
}
