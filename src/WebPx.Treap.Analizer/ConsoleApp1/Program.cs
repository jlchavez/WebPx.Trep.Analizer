using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;

namespace ConsoleApp1
{
    internal class Program
    {
        private static IEnumerable<ProductInfoHeaderValue> GetUserAgentValues()
        {
            yield return new ProductInfoHeaderValue("Mozilla", "5.0");
            yield return new ProductInfoHeaderValue("(Windows NT 10.0; Win64; x64)");
            yield return new ProductInfoHeaderValue("AppleWebKit", "537.36");
            yield return new ProductInfoHeaderValue("(KHTML, like Gecko)");
            yield return new ProductInfoHeaderValue("Chrome", "114.0.0.0");
            yield return new ProductInfoHeaderValue("Safari", "537.36");
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            var cookieContainer = new CookieContainer();
            HttpClientHandler v = new HttpClientHandler
            {
                AllowAutoRedirect = true,
                UseCookies = true,
                CookieContainer = cookieContainer,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,

            };

            var _client = new HttpClient(v);
            _client.BaseAddress = new Uri("https://app.simpleproof.com/");
            foreach (var pihv in GetUserAgentValues())
                _client.DefaultRequestHeaders.UserAgent.Add(pihv);

            var request = new HttpRequestMessage(HttpMethod.Options, "/api/proof/attestation");
            request.Headers.Add("Origin", "https://verify.simpleproof.com");
            request.Headers.Add("Access-Control-Request-Method", "POST");
            request.Headers.Add("Access-Control-Request-Headers", "content-type");
            request.Headers.Referrer = new Uri("https://verify.simpleproof.com");
            request.Headers.Accept.Clear();
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            var callOptions = _client.SendAsync(request);
            using var result = callOptions.Result;

            Console.WriteLine(result.StatusCode);

            request = new HttpRequestMessage(HttpMethod.Post, "/api/proof/attestation");
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/plain"));
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            request.Content = JsonContent.Create(new { category = "P-000024", hash = "b345a21809f4840ef83f7b28818182254bb2e5bee3225362e6598266a6889b36", num = 0 });
            callOptions = _client.SendAsync(request);
            using var result2 = callOptions.Result;
            Console.WriteLine(result2.StatusCode);

            var str = result2.Content.ReadAsStringAsync().Result;
            Console.WriteLine(str);
        }
    }
}