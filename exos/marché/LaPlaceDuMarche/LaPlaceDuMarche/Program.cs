using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;

namespace LaPlaceDuMarche
{
    internal class Program
    {
        static void Main(string[] args)
        {
            int sellerNb = 0;
            string sellerName = "";
            int standNb = 0;
            int maxValue = 0;
            int fruitQty = 0;
            WorkBook workbook = WorkBook.Load("C:\\Users\\pt04ihk\\Documents\\GitHub\\323-Programmation_fonctionnelle\\exos\\marché\\Place du marché.xlsx");
            WorkSheet sheet = workbook.WorkSheets[1];

            foreach (var cell in sheet["C2:C75"])
            {
                string cellValue = cell.StringValue;
                if (cellValue == "Pêches")
                {
                    sellerNb += 1;
                }
            }

                
            foreach (var cell in sheet["C2:C75"])
            {
                string cellValue = cell.StringValue;
                if(cellValue == "Pastèques")
                {
                   fruitQty = sheet["D" + (cell.RowIndex + 1)].Int32Value;
                   if( fruitQty > maxValue)
                   {
                        maxValue = fruitQty;
                        standNb = sheet["A" + (cell.RowIndex +1)].Int32Value;
                        sellerName = sheet["B" + (cell.RowIndex + 1)].StringValue;
                   }
                }
            }
            Console.WriteLine("Il y a " + sellerNb + " vendeurs de pêches");
            Console.WriteLine("C'est " + sellerName + " qui a le plus de pastèques (stand "+ standNb + ", " + maxValue + " pièces)" );
            Console.ReadLine();


        }
    }
}
