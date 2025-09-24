using System.Text.Json;

async Task<string> Call(string query)
{
    var client = new HttpClient();
    var response = await client.GetAsync("https://swapi.dev/api/" + query);
    var json = await response.Content.ReadAsStringAsync();

    return json;
}
var moviesJson = await Call("films");
var moviesResult = JsonSerializer.Deserialize<FilmResult>(moviesJson);
var movies = moviesResult.results;

/*Console.WriteLine("Le film avec le nom le plus long:" +
          movies.Where(
          m => m.title.Length == movies.Max(m2 => m2.title.Length))
          .Select(r => r.title + $" [{r.title.Length} lettres]")
          .First());
Console.WriteLine($"Nombre de films: {moviesResult.count}");*/

var peoplesJson = await Call("people");
var peopleResults = JsonSerializer.Deserialize<PeopleResult>(peoplesJson);
var peoples = peopleResults.results;

Console.WriteLine("Personnage apparaîssant dans le plus de films : " + peoples
    .Where(p=>p.films==peoples.Max(p2=>p2.films))
    .Select(p3=>p3.name));


class PeopleResult
{
    public int count {  get; set; }
    public List<People> results {  get; set; }
}
class People
{
    public string name { get; set; }
    public string[] films { get; set; }
}
class FilmResult
{
    public int count { get; set; }
    public List<Film> results { get; set; }
}

class Film
{
    public string title { get; set; }
}