using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PCM_Tool
{
    public class Assembly
    {
        public string name;
        public List<Item> lstItems;
        public double wip_quantity;
        public double inventory_quantity;
        public double ms_quantity;
        public int row;
        public int level;


        public Assembly(string name)
        {
            this.name = name;
            lstItems = new List<Item>();
        }
         
        public void AddItem(string name, double qty)
        {
            Item newItem = new Item();
            newItem.material_name = name;
            newItem.material_qty = Math.Round(qty);
            this.lstItems.Add(newItem);
        }

        private static double stock;
        public double pcba_on_stock
        {
            get
            {
                return stock = this.wip_quantity + this.inventory_quantity;
            }
        }

        public class Item
        {
            public string material_name;
            public double material_qty;

            public double total_quantity;
            public double total_qty
            {
                get
                {
                    total_quantity = stock * material_qty;
                    //total_quantity = sumStock() * material_qty;
                    return total_quantity;
                }
                set
                { }
            }
        }
    }
}

