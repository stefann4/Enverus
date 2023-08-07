using Enverus;
using System.Net;
using System.Net.Http;


var conversion = new Conversion();

var handler = new HttpClientHandler
{
    AllowAutoRedirect = true,
    ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true,
};

using (HttpClient client = new(handler))
{
    client.DefaultRequestHeaders.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3");
    var excelUrl = "https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl/static-files/7240366e-61cc-4acb-89bf-86dc1a0dffe8";
    try
    {
        client.BaseAddress = new Uri("https://bakerhughesrigcount.gcs-web.com/");

        var response = await client.GetAsync(excelUrl);

        if (response.IsSuccessStatusCode)
        {
            var file = await response.Content.ReadAsByteArrayAsync();

            File.WriteAllBytes("Book1.xlsx", file);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
    }
}


Console.WriteLine("Conversion starts...");

conversion.ConvertExcelToCsv();

Console.WriteLine("Conversion finished at " + DateTime.Now);
Console.ReadKey();