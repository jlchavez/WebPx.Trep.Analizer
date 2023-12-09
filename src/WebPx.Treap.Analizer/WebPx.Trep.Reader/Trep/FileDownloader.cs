using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace WebPx.Trep
{
    public class FileDownloader
    {
        public FileDownloader(CookieContainer cookieContainer)
        {
            HttpClientHandler v = new HttpClientHandler
            {
                AllowAutoRedirect = true,
                UseCookies = true,
                CookieContainer = cookieContainer,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,

            };

            _client = new HttpClient(v);
            _client.BaseAddress = new Uri("https://primeraeleccion.trep.gt/docs/d1/1687736870/");
            foreach (var pihv in GetUserAgentValues())
                _client.DefaultRequestHeaders.UserAgent.Add(pihv);
        }

        private readonly HttpClient _client;

        private IEnumerable<ProductInfoHeaderValue> GetUserAgentValues()
        {
            yield return new ProductInfoHeaderValue("Mozilla", "5.0");
            yield return new ProductInfoHeaderValue("(Windows NT 10.0; Win64; x64)");
            yield return new ProductInfoHeaderValue("AppleWebKit", "537.36");
            yield return new ProductInfoHeaderValue("(KHTML, like Gecko)");
            yield return new ProductInfoHeaderValue("Chrome", "114.0.0.0");
            yield return new ProductInfoHeaderValue("Safari", "537.36");
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
}