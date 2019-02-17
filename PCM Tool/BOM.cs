using PMC_Tool;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PCM_Tool
{
    public class BOM
    {
        public Item[] item;
        public int numberOfHDTFlags;
        public BOM(int number)
        {
            item = new Item[number];
            numberOfHDTFlags = 0;
        }

        public struct Item
        {
            public  string name;//ime koponente
            public  double bomqty;//koliko treba za bomove
            public double difference;//razlika stock-bomovi
            public string product;//koji je bom

            public string description;//opis komponente
            public double stockqty;//koliko ima na stoku
            public double mpq;//
            public double moq;//
            public double unit_price;//cena jedne komponente
            public double total_price;// ukupna cena negativne razlike stock-bomreq
            public double po;// kolicina ordered komponenti
            public double op_po;//includes further materials, which have been ordered
            public double wip;// kolicina wip komponenti
            public double scrap;//scrap
            public double safetystock;// kolicina safety stock komponenti
            public int delivery_time;//??
            public string supplier;//
            public bool HDTflag;//flag za hdt komponente koje se ne kupuju
            public double assembly_qty;// kolicina koja se nalazi na sastavljenim komponentama

            public double calculated_order_qty;//ovo bi trebalo izracunati nakon analize difference-a i moq-a(mpq-a), a kasnije i saftystock-a u funkciji calcOrderQty
        }    
        public void calcTotalPrice(int i)
        {
            //if(this.item[i].difference<0)
            this.item[i].total_price = this.item[i].unit_price * this.item[i].calculated_order_qty;
        }
        public void calcDifference(int i)
        {
            this.item[i].difference = this.item[i].stockqty - this.item[i].bomqty - this.item[i].safetystock + this.item[i].wip + this.item[i].po + this.item[i].op_po + this.item[i].assembly_qty;// - this.item[i].scrap;
        }
        public void calcOrderQty(int i)
        {
            if (searchIfMPQTypeExist(i))
            {
                if (this.item[i].mpq <= this.item[i].moq)
                {
                    if (this.item[i].difference * (-1) < this.item[i].moq)
                    {
                        this.item[i].calculated_order_qty = this.item[i].moq;
                    }
                    else
                    {
                        int ratio = (int)Math.Ceiling((this.item[i].difference * (-1) / this.item[i].mpq));
                        this.item[i].calculated_order_qty = this.item[i].mpq * ratio;
                    }
                }
                else
                {
                    if (this.item[i].difference * (-1) < this.item[i].mpq)
                    {
                        this.item[i].calculated_order_qty = this.item[i].mpq;
                    }
                    else
                    {
                        int ratio = (int)Math.Ceiling(this.item[i].difference * (-1) / this.item[i].mpq);
                        this.item[i].calculated_order_qty = this.item[i].mpq * ratio;
                    }
                }
            }
            else
            {
                if (this.item[i].mpq <= this.item[i].moq)
                {
                    if (this.item[i].difference * (-1) < this.item[i].moq)
                    {
                        this.item[i].calculated_order_qty = this.item[i].moq;
                    }
                    else
                    {
                        this.item[i].calculated_order_qty = this.item[i].difference * (-1);
                    }
                }
                else
                {
                    if (this.item[i].difference * (-1) < this.item[i].mpq)
                    {
                        this.item[i].calculated_order_qty = this.item[i].mpq;
                    }
                    else
                    {
                        this.item[i].calculated_order_qty = this.item[i].difference * (-1);
                    }
                } 
            }
        }

        private bool searchIfMPQTypeExist(int i)
        {
            foreach (string fileName in Settings.lines)
            {
                if (this.item[i].name.Substring(0, fileName.Length) == fileName)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
