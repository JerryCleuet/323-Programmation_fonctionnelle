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
                    // Prendre la valeur de la case d'à côté et la comparer avec maxValue => si supérieure, valeur de la case = maxValue
                }
                // Prendre le nom du vendeur pour la ligne où se situe maxValue et l'assigner à sellerName
                // Prendre le numéro de stand pour la même ligne et l'assigner à standNb
            }


            Console.WriteLine("Il y a " + sellerNb + " vendeurs de pêches");
            Console.WriteLine("C'est " + sellerName + "qui a le plus de pastèques (stand "+ standNb + ", " + maxValue + " pièces)" );
            Console.ReadLine();


        }
    }
}
