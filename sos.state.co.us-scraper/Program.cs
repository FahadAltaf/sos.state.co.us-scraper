using HtmlAgilityPack;
using PuppeteerExtraSharp;
using PuppeteerExtraSharp.Plugins.ExtraStealth;
using PuppeteerSharp;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Web;
using Dapper;

List<ColoradoDataset> entries = new List<ColoradoDataset>();
using (var connection = new SqlConnection(Environment.GetEnvironmentVariable("WebConnectionString")))
{
    entries = connection.Query<ColoradoDataset>("Select * from ColoradoDatasets where agentFirstName is NULL and agentCompany is NULL order by entityid desc").ToList();
}

int pages = (entries.Count + 5 - 1) / 5;
List<Task> tasks = new List<Task>();
for (int count = 1; count <= pages; ++count)
{
    int index = count - 1;
    var data = entries.Skip(index * 5).Take(5).ToList();

    Task newTask = Task.Factory.StartNew(() => { FetchRegisteredAgentDetails(data).Wait(); });

    tasks.Add(newTask);

    if (count % 10 == 0 || count == pages)
    {
        foreach (Task task in tasks)
        {
            while (!task.IsCompleted)
            { }
        }
    }
}

Console.ReadKey();



async Task FetchRegisteredAgentDetails(List<ColoradoDataset> data)
{
    var extra = new PuppeteerExtra();
    extra.Use(new StealthPlugin());
    using var browserFetcher = new BrowserFetcher();
    await browserFetcher.DownloadAsync(BrowserFetcher.DefaultChromiumRevision);
    var browser = await extra.LaunchAsync(new LaunchOptions()
    {
        DefaultViewport = null,
        Headless = true,
        Args = new string[] { "--proxy-server=" + Environment.GetEnvironmentVariable("DCProxyServer", EnvironmentVariableTarget.Process) }
    });
    var page = await browser.NewPageAsync();
    await page.AuthenticateAsync(new Credentials
    {
        Username = Environment.GetEnvironmentVariable("ProxyUsername", EnvironmentVariableTarget.Process),
        Password = Environment.GetEnvironmentVariable("ProxyPassword", EnvironmentVariableTarget.Process)
    });
    using (var connection = new SqlConnection(Environment.GetEnvironmentVariable("WebConnectionString")))
    {
        foreach (var entry in data)
        {
            try
            {
                await page.GoToAsync($"https://www.sos.state.co.us/biz/BusinessEntityDetail.do?quitButtonDestination=BusinessEntityCriteriaExt&fileId={entry.entityid}&masterFileId=");
                await page.WaitForTimeoutAsync(5000);

                var doc = new HtmlDocument();
                doc.LoadHtml(await page.GetContentAsync());
                
                var regNameNode = doc.DocumentNode.SelectSingleNode("//*[@id=\"application\"]/table/tbody/tr/td[2]/table/tbody/tr[3]/td/form/table[1]/tbody/tr[1]/td/table/tbody/tr[7]/td/table/tbody/tr[2]/td[2]");
                if (regNameNode != null)
                {
                    entry.agentCompany = HttpUtility.HtmlDecode(regNameNode.InnerText);
                }

                var mailingNode = doc.DocumentNode.SelectSingleNode("//*[@id=\"application\"]/table/tbody/tr/td[2]/table/tbody/tr[3]/td/form/table[1]/tbody/tr[1]/td/table/tbody/tr[7]/td/table/tbody/tr[4]/td[2]");
                if (mailingNode != null)
                {
                    var text = mailingNode.InnerText.Replace("\t", "");
                    entry.agentMailingFullAddress = text;
                    var parts = text.Split(",", StringSplitOptions.RemoveEmptyEntries);
                    try
                    {
                        string stateZip = "";
                        switch (parts.Length)
                        {
                            case 4:
                               entry. agentMailingAddress = parts[0];
                              entry.  agentMailingCity = parts[1];
                                 stateZip = parts[2].Trim();
                                entry.agentMailingState = stateZip.Substring(0, 2);
                                entry.agentMailingZip = stateZip.Substring(2);
                                entry.agentMailingCountry = parts[3];
                                break;
                            case 5:
                                entry.agentMailingAddress = parts[0]+" "+parts[1];
                                entry.agentMailingCity = parts[2];
                                 stateZip = parts[3].Trim();
                                entry.agentMailingState = stateZip.Substring(0, 2);
                                entry.agentMailingZip = stateZip.Substring(2);
                                entry.agentMailingCountry = parts[4];
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

                if (!string.IsNullOrEmpty(entry.agentCompany))
                {
                    string sql = @"UPDATE       ColoradoDatasets
SET                agentMailingFullAddress =@agentMailingFullAddress, agentMailingAddress =@agentMailingAddress, agentMailingCity =@agentMailingCity, agentMailingState =@agentMailingState, agentMailingZip =@agentMailingZip, agentMailingCountry =@agentMailingCountry, agentCompany =@agentCompany
where Id=@Id";
                    connection.Execute(sql, entry);
                    Console.WriteLine("Updated => " + entry.entityid);
                }
                else
                    Console.WriteLine("Agent name not found for " + entry.entityid);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to scrape {entry.entityid}. Reason: {ex.Message}");
            }
        }
        await browser.CloseAsync();
    }
}

public class ColoradoDataset
{
    public Int64 Id { get; set; }
    public string entityid { get; set; }
    public string entityname { get; set; }
    public string principaladdress1 { get; set; }
    public string principalcity { get; set; }
    public string principalstate { get; set; }
    public string principalzipcode { get; set; }
    public string principalcountry { get; set; }
    public string entitystatus { get; set; }
    public string jurisdictonofformation { get; set; }
    public string entitytype { get; set; }
    public string agentCompany { get; set; }
    public string agentfirstname { get; set; }
    public string agentlastname { get; set; }
    public string agentprincipaladdress1 { get; set; }
    public string agentprincipalcity { get; set; }
    public string agentprincipalstate { get; set; }
    public string agentprincipalzipcode { get; set; }
    public string agentprincipalcountry { get; set; }

    public string agentMailingFullAddress { get; set; }
    public string agentMailingAddress { get; set; }
    public string agentMailingCity { get; set; }
    public string agentMailingState { get; set; }
    public string agentMailingZip { get; set; }
    public string agentMailingCountry { get; set; }

    public DateTime entityformdate { get; set; }
    public string agentmiddlename { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public string JobId { get; set; }
}