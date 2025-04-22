using HtmlAgilityPack;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.IO;
using System.ComponentModel;
using System.Xml;

class CarScraper
{
    private static readonly string BASE_URL = "https://www.polovniautomobili.com";
    private static readonly string SEARCH_URL = "https://www.polovniautomobili.com/auto-oglasi/pretraga";
    private static readonly HttpClient client = new HttpClient();

    static async Task Main()
    {
        var carList = new List<Dictionary<string, string>>();
        var seenLinks = new HashSet<string>();

        var parameters = new List<KeyValuePair<string, string>>
        {
            new KeyValuePair<string, string>("price_to", "7000"),
            new KeyValuePair<string, string>("year_from", "2010"),
            new KeyValuePair<string, string>("power_to", "80"),
            new KeyValuePair<string, string>("door_num", "3013"),
            new KeyValuePair<string, string>("damaged[]", "3799"),
            new KeyValuePair<string, string>("sort", "basic")
        };

        var chassisValues = new List<string> { "277", "2631", "2633", "2634", "2632" };
        foreach (var chassis in chassisValues)
        {
            parameters.Add(new KeyValuePair<string, string>("chassis[]", chassis));
        }

        for (int page = 1; page <= 1; page++)
        {
            Console.WriteLine($"Obradjujem stranicu {page}...");
            parameters.Add(new KeyValuePair<string, string>("page", page.ToString()));
            var url = AddQueryString(SEARCH_URL, parameters);
            var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();

            var doc = new HtmlDocument();
            doc.LoadHtml(content);

            var ads = doc.DocumentNode.SelectNodes("//div[contains(@class,'top-featured-wrapper')]//h2[@class='brand-and-model']/a");
            if (ads != null)
            {
                foreach (var ad in ads)
                {
                    var relativeUrl = ad.GetAttributeValue("href", "");
                    if (!string.IsNullOrEmpty(relativeUrl))
                    {
                        var fullUrl = BASE_URL + relativeUrl;
                        if (!seenLinks.Contains(fullUrl))
                        {
                            var carData = await GetFullCarData(fullUrl);
                            if (carData != null && IsCriteriaFulfilled(carData))
                            {
                                carData["Ocjena"] = CalculateRating(carData);
                                carList.Add(carData);
                                seenLinks.Add(fullUrl);
                            }
                        }
                        await Task.Delay(RandomTime(1000, 2500));
                    }
                }
            }

            var ads2 = doc.DocumentNode.SelectNodes("//article[contains(@class,'classified')]//div[@class='textContent']/h2/a");
            if (ads2 != null)
            {
                foreach (var ad2 in ads2)
                {
                    var relativeUrl = ad2.GetAttributeValue("href", "");
                    if (!string.IsNullOrEmpty(relativeUrl))
                    {
                        var fullUrl = BASE_URL + relativeUrl;
                        if (!seenLinks.Contains(fullUrl))
                        {
                            var carData = await GetFullCarData(fullUrl);
                            if (carData != null && IsCriteriaFulfilled(carData))
                            {
                                carData["Ocjena"] = CalculateRating(carData);
                                carList.Add(carData);
                                seenLinks.Add(fullUrl);
                            }
                        }
                        await Task.Delay(RandomTime(1000, 2500));
                    }
                }
            }

            await Task.Delay(RandomTime(2000, 4000));
        }

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Automobili");
            if (carList.Count > 0)
            {
                var headers = new List<string>(carList[0].Keys);
                for (int i = 0; i < headers.Count; i++)
                    worksheet.Cell(1, i + 1).Value = headers[i];

                for (int i = 0; i < carList.Count; i++)
                    for (int j = 0; j < headers.Count; j++)
                        worksheet.Cell(i + 2, j + 1).Value = carList[i][headers[j]];
            }
            workbook.SaveAs("detaljni_polovni_automobili.xlsx");
        }

        Console.WriteLine("✅ Gotovo! Sačuvano u detaljni_polovni_automobili.xlsx");
    }

    static async Task<Dictionary<string, string>> GetFullCarData(string url)
    {
        try
        {
            var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();

            var doc = new HtmlDocument();
            doc.LoadHtml(content);

            var titleNode = doc.DocumentNode.SelectSingleNode("//span[contains(@class,'generationTitle')]");
            var priceNode = doc.DocumentNode.SelectSingleNode("//span[contains(@class,'priceClassified')]");
            var locationNode = doc.DocumentNode.SelectSingleNode("//div[contains(@class,'infoBox')]/div[@class = 'uk-grid']");

            var specs = new Dictionary<string, string>();
            var rows = doc.DocumentNode.SelectNodes("//div[contains(@class,'divider')]/div[@class='uk-grid']");
            if (rows != null)
            {
                foreach (var row in rows)
                {
                    var divs = row.SelectNodes("./div[contains(@class, 'uk-width-1-2')]");
                    if (divs != null && divs.Count == 2)
                    {
                        var label = HtmlEntity.DeEntitize(divs[0].InnerText.Trim().TrimEnd(':'));
                        var value = HtmlEntity.DeEntitize(divs[1].InnerText.Trim());
                        specs[label] = value;
                    }
                }
            }

            return new Dictionary<string, string>
            {
                ["Marka"] = specs.GetValueOrDefault("Marka"),
                ["Model"] = specs.GetValueOrDefault("Model"),
                ["Karoserija"] = specs.GetValueOrDefault("Karoserija"),
                ["Naslov"] = titleNode?.InnerText.Trim(),
                ["Cena"] = priceNode?.InnerText.Trim(),
                ["Godište"] = specs.GetValueOrDefault("Godište"),
                ["Kilometraža"] = specs.GetValueOrDefault("Kilometraža"),
                ["Gorivo"] = specs.GetValueOrDefault("Gorivo"),
                ["Kubikaža"] = specs.GetValueOrDefault("Kubikaža"),
                ["Snaga motora"] = specs.GetValueOrDefault("Snaga motora"),
                ["Menjač"] = specs.GetValueOrDefault("Menjač"),
                ["Pogon"] = specs.GetValueOrDefault("Pogon"),
                ["Emisiona klasa"] = specs.GetValueOrDefault("Emisiona klasa"),
                ["Broj vrata"] = specs.GetValueOrDefault("Broj vrata"),
                ["Broj sedišta"] = specs.GetValueOrDefault("Broj sedišta"),
                ["Registrovan do"] = specs.GetValueOrDefault("Registrovan do"),
                ["Stanje"] = specs.GetValueOrDefault("Stanje"),
                ["Emisiona klasa"] = specs.GetValueOrDefault("Emisiona klasa motora"),
                ["Klima"] = specs.GetValueOrDefault("Klima"),
                ["Boja"] = specs.GetValueOrDefault("Boja"),
                ["Materijal enterijera"] = specs.GetValueOrDefault("Materijal enterijera"),
                ["Boja enterijera"] = specs.GetValueOrDefault("Boja enterijera"),
                ["Plivajući zamajac"] = specs.GetValueOrDefault("Plivajući zamajac"),
                ["Strana volana"] = specs.GetValueOrDefault("Strana volana"),
                ["Ocjena"] = "0",
                ["Link"] = url
            };
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[GREŠKA] {url}: {ex.Message}");
            return null;
        }
    }

    static bool IsCriteriaFulfilled(Dictionary<string, string> car)
    {
        if (!int.TryParse(car.GetValueOrDefault("Kilometraža")?.Replace(".", "").Replace(" km", "").Trim(), out int km) || km > 200000)
            return false;

        if (!int.TryParse(car.GetValueOrDefault("Kubikaža")?.Replace(".", "").Replace(" cm3", "").Replace("cm", "").Trim(), out int cm3) || cm3 > 2000)
            return false;

        var powerStr = car.GetValueOrDefault("Snaga motora");
        if (powerStr != null && powerStr.Contains("/"))
        {
            var parts = powerStr.Split('/');
            if (!int.TryParse(parts[1].Split('(')[0].Trim(), out int power) || power < 50)
                return false;
        }

        var emission = car.GetValueOrDefault("Emisiona klasa")?.ToLower();
        if  (!(emission.Contains("euro 4") || emission.Contains("euro 5") || emission.Contains("euro 6")))
            return false;

        var transmission = car.GetValueOrDefault("Menjač")?.ToLower();
        if (!transmission.Contains("manuelni"))
            return false;

        var airConditioning = car.GetValueOrDefault("Menjač")?.ToLower();
        if (airConditioning.Contains("nema"))
            return false;

        var wheel = car.GetValueOrDefault("Menjač")?.ToLower();
        if (wheel.Contains("desni"))
            return false;

        if (!int.TryParse(car.GetValueOrDefault("Broj vrata")?.Trim(), out int doors) || doors < 4)
            return false;

        return true;
    }


    static string CalculateRating(Dictionary<string, string> car)
    {
        try
        {
            int rating = 0;

            if (double.TryParse(car.GetValueOrDefault("Cena")?.Replace(".", "").Replace("€", "").Trim(), out double price))
            {
                rating += price <=  4500 ? 5 : 3;
            }

            if (int.TryParse(car.GetValueOrDefault("Godište")?.Replace(".", "").Trim(), out int age))
            {
                rating += age >= 2020 ? 5 : 
                          age >= 2015 ? 4 : 
                          age >= 2010 ? 3 : 1;
            }

            if (car.TryGetValue("Kilometraža", out string kmStr))
            {
                kmStr = kmStr.Replace(".", "").Replace(" km", "").Trim();
                if (int.TryParse(kmStr, out int km))
                {
                    rating += km <= 50000 ? 5 : 
                              km <= 100000 ? 4 :
                              km <= 150000 ? 3 :
                              km <= 200000 ? 2 : 1;
                }
            }

            if (car.TryGetValue("Kubikaža", out string kubikStr))
            {
                kubikStr = kubikStr.Replace(".", "").Replace(" cm3", "").Replace("cm", "").Trim();
                if (int.TryParse(kubikStr, out int cm3))
                {
                    rating += cm3 <= 1000 ? 3 :
                              cm3 <= 1300 ? 4 :
                              cm3 <= 1600 ? 5 : 2;
                }
            }

            if (car.TryGetValue("Snaga motora", out string powerStr))
            {
                var match = System.Text.RegularExpressions.Regex.Match(powerStr, @"\d+");
                if (match.Success && int.TryParse(match.Value, out int power))
                {
                    rating += power >= 80 ? 5 :
                              power >= 70 ? 4 :
                              power >= 60 ? 3 : 2;
                }
            }

            if (car.TryGetValue("Emisiona klasa", out string em))
            {
                rating += em.Contains("Euro 6") ? 5 : 
                          em.Contains("Euro 5") ? 4 : 0;
            }


            return rating > 0 ? rating.ToString("0") : "N/A";
        }
        catch
        {
            return "N/A";
        }
    }


    static string AddQueryString(string uri, List<KeyValuePair<string, string>> parameters)
    {
        var sb = new StringBuilder(uri);
        sb.Append(uri.Contains("?") ? "&" : "?");

        foreach (var p in parameters)
        {
            sb.Append(Uri.EscapeDataString(p.Key));
            sb.Append('=');
            sb.Append(Uri.EscapeDataString(p.Value));
            sb.Append('&');
        }

        if (parameters.Count > 0)
            sb.Length--;

        return sb.ToString();
    }

    static int RandomTime(int minMs, int maxMs)
    {
        return new Random().Next(minMs, maxMs);
    }
}