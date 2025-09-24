using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.CodeDom.Compiler;
using System.Diagnostics;

namespace MibMap
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var i18n = new Dictionary<string, string>()
            {
                { "Pommes","Apples"},
                { "Poires","Pears"},
                { "Pastèques","Watermelons"},
                { "Melons","Melons"},
                { "Noix","Nuts"},
                { "Raisin","Grapes"},
                { "Pruneaux","Plums"},
                { "Myrtilles","Blueberries"},
                { "Groseilles","Berries"},
                { "Tomates","Tomatoes"},
                { "Courges","Pumpkins"},
                { "Pêches","Peaches"},
                { "Haricots","Beans"}
            };


            //Ex 1
            /* 
              var changedProducer = products.Select(p => $"{p.Producer.Substring(0, 3)}...{p.Producer.Substring(p.Producer.Length -1)} {p.ProductName = i18n[p.ProductName]} {p.Quantity * p.PricePerUnit}");
            Console.WriteLine("Seller | Product | CA");
            //changedProducer.ToList().ForEach(name => { Console.WriteLine(name); });

            string filePath = "MyCSV.csv";

            using (StreamWriter sw = new StreamWriter(filePath))
            {
                Console.WriteLine("Seller | Product | CA");
                changedProducer.ToList().ForEach(name => sw.WriteLine(name));
            }*/

            // Dashboard 
            /*
            var result = products.Select(p => (
                 Nom: p.Producer.Substring(0, 1)+(p.Producer.Length - 1)+p.Producer.Last(), 
                 NomProduit: p.ProductName, 
                 Stock: (p.Quantity < 10 ? "Stock faible"  : p.Quantity >=10 && p.Quantity <= 15 ? "Stock normal" : "stock élevé"),
                 Prix: p.Quantity < 10 ? (15 * p.PricePerUnit / 100) + p.PricePerUnit : p.Quantity >= 10 && p.Quantity <= 15 ? (5 * p.PricePerUnit / 100) + p.PricePerUnit : p.PricePerUnit,
                 CA: p.PricePerUnit * p.Quantity > 100 ? "Premium" : "Standard"));

            // result.ToList().ForEach(res => Console.WriteLine(res));
             var options = new JsonSerializerOptions { IncludeFields = true };

             string json = JsonSerializer.Serialize(result, options);

             File.WriteAllText("file.json", json); */

            List<Product> products = new List<Product>
            {
                new Product { Location = 1, Producer = "Bornand", ProductName = "Pommes", Quantity = 20,Unit = "kg", PricePerUnit = 5.50 },
                new Product { Location = 1, Producer = "Bornand", ProductName = "Poires", Quantity = 16,Unit = "kg", PricePerUnit = 5.50 },
                new Product { Location = 1, Producer = "Bornand", ProductName = "Pastèques", Quantity = 14,Unit = "pièce", PricePerUnit = 5.50 },
                new Product { Location = 1, Producer = "Bornand", ProductName = "Melons", Quantity = 5,Unit = "kg", PricePerUnit = 5.50 },
                new Product { Location = 2, Producer = "Dumont", ProductName = "Pommes", Quantity = 20,Unit = "sac", PricePerUnit = 5.50 },
                new Product { Location = 2, Producer = "Dumont", ProductName = "Raisin", Quantity = 6,Unit = "kg", PricePerUnit = 5.50 },
                new Product { Location = 2, Producer = "Dumont", ProductName = "Pommes", Quantity = 13,Unit = "kg", PricePerUnit = 5.50 },
                new Product { Location = 2, Producer = "Dumont", ProductName = "Myrtilles", Quantity = 12,Unit = "kg", PricePerUnit = 5.50 },

                //...
            };

            // Le retour du marché
            // 1
            /*
            var availableGroseilles = products.Where(p => p.ProductName == "Pommes").Select(p => p.Quantity).Sum();
            Console.WriteLine(availableGroseilles);*/

            // 2
            var possibleCA = products
                .GroupBy(p => p.Producer).Distinct()
                .Select(group => new { Prix = group.Quantity * group.PricePerUnit });
            possibleCA.ToList().ForEach(name => Console.WriteLine(name));


        } 

        class Product
        {
            public int Location  { get; set; }
            public string Producer { get; set; }
            public string ProductName { get; set; }
            public int Quantity { get; set; }
            public string Unit { get; set; }
            public double PricePerUnit { get; set; }
        }
        // Mesure de perfs
        /*
        static (long time, long memory) MesurePerf(Action action, int iterations = 1000)
        {
            // Forcer le garbage collection avant mesure
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            var memoryBefore = GC.GetTotalMemory(false);
            var stopwatch = Stopwatch.StartNew();

            for (int i = 0; i < iterations; i++)
            {
                action();
            }

            stopwatch.Stop();
            var memoryAfter = GC.GetTotalMemory(false);

            return (stopwatch.ElapsedMilliseconds, memoryAfter - memoryBefore);
        }*/

    }
}
