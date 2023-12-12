using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Runtime.InteropServices.JavaScript;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace WebPx.Trep
{
    public class FileDownloader
    {
        private HttpClientHandler clientHandler; 

        public FileDownloader(CookieContainer cookieContainer)
        {
            clientHandler = new HttpClientHandler
            {
                AllowAutoRedirect = true,
                UseCookies = true,
                CookieContainer = cookieContainer,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,

            };

            _client = CreateClient("https://primeraeleccion.trep.gt/docs/d1/1687736870/");
            _simpleProofClient = CreateClient("https://app.simpleproof.com/");
        }

        private HttpClient CreateClient(string url)
        {
            var client = new HttpClient(clientHandler);
            client.BaseAddress = new Uri(url);
            foreach (var pihv in GetUserAgentValues())
                client.DefaultRequestHeaders.UserAgent.Add(pihv);
            return client;
        }

        private readonly HttpClient _client;
        private readonly HttpClient _simpleProofClient;

        private IEnumerable<ProductInfoHeaderValue> GetUserAgentValues()
        {
            yield return new ProductInfoHeaderValue("Mozilla", "5.0");
            yield return new ProductInfoHeaderValue("(Windows NT 10.0; Win64; x64)");
            yield return new ProductInfoHeaderValue("AppleWebKit", "537.36");
            yield return new ProductInfoHeaderValue("(KHTML, like Gecko)");
            yield return new ProductInfoHeaderValue("Chrome", "114.0.0.0");
            yield return new ProductInfoHeaderValue("Safari", "537.36");
        }

        public async Task<Attestation?> GetAttestation(string hash)
        {
            
            var request = new HttpRequestMessage(HttpMethod.Options, "/api/proof/attestation");
            request.Headers.Add("Origin", "https://verify.simpleproof.com");
            request.Headers.Add("Access-Control-Request-Method", "POST");
            request.Headers.Add("Access-Control-Request-Headers", "content-type");
            request.Headers.Referrer = new Uri("https://verify.simpleproof.com");
            request.Headers.Accept.Clear();
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            using var result = await _simpleProofClient.SendAsync(request);

            request = new HttpRequestMessage(HttpMethod.Post, "/api/proof/attestation");
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/plain"));
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            request.Content = JsonContent.Create(new { category = "P-000024", hash = hash, num = 0 });
            using var result2 = await _simpleProofClient.SendAsync(request);

            var str = await result2.Content.ReadAsStringAsync();
            var jsonNode = JsonNode.Parse(str);
            var srcfile = jsonNode["srcfile"].ToString();
            //var srcSize = jsonNode["srcSize"];
            var receptionDate = jsonNode["reception_date"]!.ToString()!;
            var attestation = new Attestation()
            {
                SrcFile = Convert.FromBase64String(srcfile),
                ReceptionDate = DateTime.ParseExact(receptionDate, "yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InstalledUICulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal)
            };
            return attestation;
        }

        public async Task<string> GetFile(string path)
        {
            var request = new HttpRequestMessage(HttpMethod.Get, "https://API.com/api");
            using var result = await _client.GetAsync(path);
            var content = await result.Content.ReadAsByteArrayAsync();
            var s = Encoding.ASCII.GetString(content);
            return s;
        }

        public async Task<byte[]> DownloadAsync(string path)
        {
            var request = new HttpRequestMessage(HttpMethod.Get, "https://API.com/api");
            using var result = await _client.GetAsync(path);
            var content = await result.Content.ReadAsByteArrayAsync();
            return content;
        }

        public async Task<byte[]?> DownloadAsync(int centro, int doc, int mesa, string codigo, string? image)
        {

            var uri = $"cdt{centro:0000}{doc:00}/{codigo}.jpg";
            if (!string.IsNullOrEmpty(image))
                uri = $"/docs/{image}";

            retry:

            var request = new HttpRequestMessage(HttpMethod.Get, "https://API.com/api");
            using (var result = await _client.GetAsync(uri))
            {
                bool hasData = false;
                if (result.IsSuccessStatusCode)
                {
                    var bytes = await result.Content.ReadAsByteArrayAsync();
                    hasData = bytes is { Length: > 0 };
                    if (hasData)
                    {
                        retryCount = 0;
                        isRetrying = false;
                        return bytes;
                    }
                }

                if (hasData)
                    await File.AppendAllTextAsync("FailedFiles.txt", $"{doc},{centro},{mesa},{uri} Empty file.{Environment.NewLine}");
                else
                {
                    var delay = result.ReasonPhrase == "Accepted";
                    if (delay)
                    {
                        if (retryCount > 0)
                            seconds += 10;
                        isRetrying = true;
                        retryCount++;
                        await Task.Delay(seconds * 1000);
                        goto retry;
                    }

                    await File.AppendAllTextAsync("FailedFiles.txt", $"{doc},{centro},{mesa},{uri} Error: {result.ReasonPhrase}.{Environment.NewLine}");
                }
            }

            return null;
        }

        private static bool isRetrying = false;
        private static int retryCount = 0;
        private static int seconds = 30;
    }

    public sealed class Attestation
    {
        public Attestation()
        {
            
        }

        public DateTime? ReceptionDate { get; set; }
        public byte[]? SrcFile { get; set; }
    }
}