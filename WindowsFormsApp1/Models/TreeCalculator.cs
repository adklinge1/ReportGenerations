using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Models
{
    public class TreeCalculator
    {
        private Dictionary<string, double> _treeTypeToPriceFactor;
        private Dictionary<string, double> _palmTypeToPriceFactor;

        public async Task LoadTreePricesAsync()
        {
            _treeTypeToPriceFactor = _treeTypeToPriceFactor ?? await FetchTreePricesAsync();
            _palmTypeToPriceFactor = ExcelReader.ExcelReader.TryReadPalmSpecies().ToDictionary(kcp => kcp.Key, kvp => kvp.Value.SpeciesRate);
        }

        public double? TryToGetTreePrice(Tree tree)
        {
            if (tree == null || string.IsNullOrWhiteSpace(tree.Species))
            {
                return null;
            }

            if (IsPalmTree(tree))
            {
                return CalculatePalmTreeValue(tree);
            }

            if (_treeTypeToPriceFactor?.ContainsKey(tree.Species.Trim()) != true)
            {
                return null;
            }

            return CalculateTreeValue(tree);
        }
        
        static double CalculateTreeSize(Tree tree)
        {
            if (tree.HasMultipleStem)
            {
                // TODO: implement Multi-stem pricing
                return 0;
            }

            // Single trunk logic
            return 3.14 * ((tree.StemDiameter * tree.StemDiameter) / 4);
        }

        private double CalculateTreeValue(Tree tree)
        {
            double treeSize = CalculateTreeSize(tree);
            double treeFactor = _treeTypeToPriceFactor[tree.Species.Trim()];

            if (tree.LocationRate > 0 && tree.HealthRate > 0 && treeSize > 0)
            {
                // In the agriculture department the tree health and location are numbers between [0-1]
                double healthNormalized = (double)tree.HealthRate / 5;
                double locationNormalized = (double)tree.LocationRate / 5;

                // Actual calculation logic
                return 20 * treeFactor * locationNormalized * healthNormalized * treeSize;
            }

            // Return a default value or handle the case when inputs are not valid
            return 0.0;
        }

        private double CalculatePalmTreeValue(Tree tree)
        {
            // In the agriculture department the tree health and location are numbers between [0-1]
            double healthNormalized = (double)tree.HealthRate / 5;
            double locationNormalized = (double)tree.LocationRate / 5;
            double palmValue = _palmTypeToPriceFactor[tree.Species];
            
            return 1500 * palmValue * tree.Height * healthNormalized * locationNormalized;
        }

        private bool IsPalmTree(Tree tree)
        {
            return _palmTypeToPriceFactor?.ContainsKey(tree.Species) == true;
        }


        public async Task<Dictionary<string, double>> FetchTreePricesAsync()
        {
            string url = "https://www.gov.il/Apps/Moag/TreeValueCalculator/treecalc.html";
            HtmlDocument htmlDocument = new HtmlDocument();

            try
            {
                string htmlContent = await DownloadHtmlAsync(url);

                htmlDocument.LoadHtml(htmlContent);

                Dictionary<string, double> treeToFactorMapping = ParseSelectOptions(htmlDocument, "tree-type");

                if (treeToFactorMapping != null && treeToFactorMapping.Any())
                {
                    return treeToFactorMapping;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}. Usig static local tree price mapping");
            }

            try
            {
                // Get the absolute path to the HTML file
                string absoluteFilePath = Path.Combine(Directory.GetCurrentDirectory(), @"TreeCalculator\StaticHtmlFile.html");

                // Read the HTML content
                string htmlContent = File.ReadAllText(absoluteFilePath);
                htmlDocument.LoadHtml(htmlContent);

                Dictionary<string, double> treeToFactorMapping = ParseSelectOptions(htmlDocument, "tree-type");

                return treeToFactorMapping;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        static async Task<string> DownloadHtmlAsync(string url)
        {
            // Explicitly set TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            using (HttpClientHandler handler = new HttpClientHandler { 
                AllowAutoRedirect = true,
                UseDefaultCredentials = false,
                AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate
            })
            using (HttpClient httpClient = new HttpClient(handler))
            {
                httpClient.Timeout = TimeSpan.FromSeconds(60); // Increase timeout further
                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36");
                httpClient.DefaultRequestHeaders.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8");
                httpClient.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.9,he;q=0.8");
                httpClient.DefaultRequestHeaders.Add("Cache-Control", "no-cache");
                
                Console.WriteLine($"Attempting to download HTML from {url}...");
                
                try {
                    // Download the HTML content
                    HttpResponseMessage response = await httpClient.GetAsync(url);
                    
                    // Check if the request was successful
                    response.EnsureSuccessStatusCode();
                    
                    Console.WriteLine("Successfully received response");
                    
                    // Read and return the HTML content as a string
                    return await response.Content.ReadAsStringAsync();
                }
                catch (TaskCanceledException ex) {
                    Console.WriteLine($"Task cancelled: {ex.Message}. Inner exception: {ex.InnerException?.Message}");
                    throw;
                }
                catch (Exception ex) {
                    Console.WriteLine($"Error downloading HTML: {ex.GetType().Name} - {ex.Message}");
                    throw;
                }
            }
        }

        static Dictionary<string, double> ParseSelectOptions(HtmlDocument document, string selectId)
        {
            Dictionary<string, double> optionsDictionary = new Dictionary<string, double>();

            HtmlNode selectNode = document.DocumentNode.SelectSingleNode($"//select[@id='{selectId}']");

            if (selectNode != null)
            {
                foreach (var optionNode in selectNode.SelectNodes("option"))
                {
                    string valueStr = optionNode.GetAttributeValue("value", "");
                    string text = optionNode.InnerText.Trim();

                    if (!string.IsNullOrEmpty(valueStr) && double.TryParse(valueStr, out double value))
                    {
                        optionsDictionary[text] = value;
                    }
                }
            }

            return optionsDictionary;
        }
    }
}
