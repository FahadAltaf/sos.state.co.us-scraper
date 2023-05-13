using HtmlAgilityPack;
using PuppeteerExtraSharp;
using PuppeteerExtraSharp.Plugins.ExtraStealth;
using PuppeteerSharp;
using System.Web;

List<string> ids = new List<string>() { "20231474876" };
var agents = await FetchRegisteredAgentDetails(ids);
Console.ReadKey();



async Task<List<RegisteredAgent>> FetchRegisteredAgentDetails(List<string> ids)
{
    var extra = new PuppeteerExtra();
    extra.Use(new StealthPlugin());
    using var browserFetcher = new BrowserFetcher();
    await browserFetcher.DownloadAsync(BrowserFetcher.DefaultChromiumRevision);
    var browser = await extra.LaunchAsync(new LaunchOptions()
    {
        Headless = false,
        DefaultViewport = null
    });

    var page = await browser.NewPageAsync();
    List<RegisteredAgent> agents = new List<RegisteredAgent>();
    foreach (var id in ids)
    {
        await page.GoToAsync($"https://www.sos.state.co.us/biz/BusinessEntityDetail.do?quitButtonDestination=BusinessEntityCriteriaExt&fileId={id}&masterFileId=");
        await page.WaitForTimeoutAsync(2000);

        var doc = new HtmlDocument();
        doc.LoadHtml(await page.GetContentAsync());
        RegisteredAgent agent = new RegisteredAgent();
        var regNameNode = doc.DocumentNode.SelectSingleNode("//*[@id=\"application\"]/table/tbody/tr/td[2]/table/tbody/tr[3]/td/form/table[1]/tbody/tr[1]/td/table/tbody/tr[7]/td/table/tbody/tr[2]/td[2]");
        if (regNameNode != null)
        {
            agent.Name = HttpUtility.HtmlDecode(regNameNode.InnerText);
        }

        var mailingNode = doc.DocumentNode.SelectSingleNode("//*[@id=\"application\"]/table/tbody/tr/td[2]/table/tbody/tr[3]/td/form/table[1]/tbody/tr[1]/td/table/tbody/tr[7]/td/table/tbody/tr[4]");
        if (mailingNode != null)
        {
            var text = mailingNode.InnerText;
            agent.FullAddress = text;
            var parts = text.Split(",", StringSplitOptions.RemoveEmptyEntries);
            try
            {
                switch (parts.Length)
                {
                    case 4:
                        agent.Address = parts[0];
                        agent.City = parts[1];
                        var stateZip = parts[2].Trim();
                        agent.State = stateZip.Substring(0, 2);
                        agent.Zip = stateZip.Substring(2);
                        agent.Country = parts[3];
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to split address");
            }
        }

        if(string.IsNullOrEmpty(agent.Name))
        {
            agents.Add(agent);
        }
    }
    await browser.CloseAsync();
    return agents;
}


public class RegisteredAgent
{
    public string Name { get; set; }
    public string FullAddress { get; set; }
    public string Address { get; set; }
    public string City { get; set; }
    public string State { get; set; }
    public string Zip { get; set; }
    public string Country { get; set; }
}