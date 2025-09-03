using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Words
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Partie 1
            // Exercice A
            /* string[] words = { "bonjour", "hello", "monde", "vert", "rouge", "bleu", "jaune" };
             Func<string, bool> noXAndMoreThanFour = w => !w.Contains('x') && w.Count() >= 4 && w.Count() == Math.Round(words.Average(word => word.Length), 0);
             words = words.Where(noXAndMoreThanFour).ToArray();
             words.OrderByDescending(w => w).ToList().ForEach(w => {Console.WriteLine(w); });*/

            // Exercice B
            /* string[] words = { "whatThe!!!", "bonjour", "hello", "monde", "vert", "rouge", "bleu", "jaune", "My kingdom for a horse !", "Ooops I did it again" };
             words = words.Skip(1).Reverse().Skip(2).Reverse().ToArray();
             words.ToList().ForEach(w => { Console.WriteLine(w); });*/

            // Exercice C
            /*  string[] words = { "+++++", "<<<<<", ">>>>>", "bonjour", "hello", "@@@@", "vert", "rouge", "bleu", "jaune", "#####", "%%%%%%%" };
              Func<string, bool> startsWithALetter = word=> !Regex.IsMatch(word, "^[a-zA-Z]");
              words = words.SkipWhile(startsWithALetter).ToArray();
              words.ToList().ForEach(word => {Console.WriteLine(word);});*/
            // Exercice D
            /*
            string[] words = { "i am the winner", "hello", "monde", "vert", "rouge", "bleu", "i am the looser" };
            var winner = words.First();
            var looser = words.Last();
            Console.WriteLine($"The winner is {winner}");
            Console.WriteLine($"The looser is {looser}");
            Console.ReadLine(); */

            // Partie 2

            Func<string, double> Epsilon = word => Math.Sqrt(word.Length);


            // Partie 3
            List<string> frenchWords = new List<string>() {
                "Merci",
                "Hotdog",
                "Oui",
                "Non",
                "Désolé",
                "Réunion",
                "Manger",
                "Boire",
                "Téléphone",
                "Ordinateur",
                "Internet",
                "Email",
                "Sandwich",
                "Hello",
                "Taxi",
                "Hotel",
                "Gare",
                "Train",
                "Bus",
                "Métro",
                "Tramway",
                "Vélo",
                "Voiture",
                "Piéton",
                "Feu rouge",
                "Cédez",
                "Ralentir",
                "gauche",
                "droite",
                "Continuer",
                "Sandwich",
                "Retourner",
                "Arrêter",
                "Stationnement",
                "Parking",
                "Interdit",
                "Péage",
                "Trafic",
                "Route",
                "Rond-point",
                "Football",
                "Carrefour",
                "Feu",
                "Panneau",
                "Vitesse",
                "Tramway",
                "Aéroport",
                "Héliport",
                "Port",
                "Ferry",
                "Bateau",
                "Canot",
                "Kayak",
                "Paddle",
                "Surf",
                "Plage",
                "Mer",
                "Océan",
                "Rivière",
                "Lac",
                "Étang",
                "Marais",
                "Forêt",
                "Hello",
                "Montagne",
                "Vallée",
                "Plaine",
                "Désert",
                "Jungle",
                "Savane",
                "Volleyball",
                "Tundra",
                "Glacier",
                "Neige",
                "Pluie",
                "Soleil",
                "Nuage",
                "Vent",
                "Tempête",
                "Ouragan",
                "Tornade",
                "Séisme",
                "Tsunami",
                "Volcan",
                "Éruption",
                "Ciel"
            };

            List<string> englishWords = new List<string>()
                {"Thanks", "Hotdog", "Yes", "No", "Sorry","Reunion","Eat","Internet","Sandwich", "Hello","Bus"
            };

            var filteredWords = frenchWords.Where(w => englishWords.Contains(w));

            foreach (var item in filteredWords)
            {
                Console.WriteLine(item);
            }
        }
    }
        
} 
