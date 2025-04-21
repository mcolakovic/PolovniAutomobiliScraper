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

        for (int page = 1; page <= 202; page++)
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
                            if (carData != null)
                            {
                                carList.Add(carData);
                                seenLinks.Add(fullUrl); // beležimo da smo obradili ovaj oglas
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
                            if (carData != null)
                            {
                                carList.Add(carData);
                                seenLinks.Add(fullUrl); // beležimo da smo obradili ovaj oglas
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
                ["Link"] = url
            };
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[GREŠKA] {url}: {ex.Message}");
            return null;
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
            sb.Length--; // remove last &

        return sb.ToString();
    }

    static int RandomTime(int minMs, int maxMs)
    {
        return new Random().Next(minMs, maxMs);
    }
}