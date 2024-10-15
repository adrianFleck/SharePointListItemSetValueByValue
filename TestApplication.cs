using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Net;
using System.Text.Json;
using File = System.IO.File;

namespace TestApp;

public class TestApplication
{
    public async Task Run()
    {
        // Load config from config.json
        var fileContent = await File.ReadAllTextAsync("config.json");
        // Deserialize the config
        var config = JsonDocument.Parse(fileContent);

        // Constants
        const string listUrl = "/Lists/MyList";
        const int listItemId = 1;
        const string termFieldName = "MyTermField";
        const string newTermGuid = "6ed6db23-43a6-46b0-aa5a-42775e76dac0";

        // Get the SharePoint site URL, app service user name, password, and client ID from the config
        var sharePointSiteUrl = config.RootElement.GetProperty("SharePoint__SiteUrl").GetString();
        var sharePointAppServiceUserName = config.RootElement.GetProperty("SharePoint__AppServiceUserName").GetString();
        var sharePointAppServiceUserPassword = config.RootElement.GetProperty("SharePoint__AppServiceUserPassword").GetString();
        var sharePointAppServiceClientId = config.RootElement.GetProperty("SharePoint__AppServiceClientId").GetString();

        Console.WriteLine($"Connecting to SharePoint site \"{sharePointSiteUrl}\"");
        var passwordSecureString = new NetworkCredential("", sharePointAppServiceUserPassword).SecurePassword;
        using var authManager = PnP.Framework.AuthenticationManager.CreateWithCredentials(sharePointAppServiceClientId, sharePointAppServiceUserName, passwordSecureString);
        using var context = await authManager.GetContextAsync(sharePointSiteUrl);
        context.ExecutingWebRequest += (sender, e) => e.WebRequestExecutor.WebRequest.UserAgent = "NONISV|PhoenixContact|Salesnet/1.0";

        var web = context.Web;
        context.Load(web);
        await context.ExecuteQueryAsync();
        Console.WriteLine($"Connected to site: {web.Title}");

        var list = web.GetListByUrl(listUrl);
        context.Load(list);
        await context.ExecuteQueryAsync();
        Console.WriteLine($"Connected to list: {list.Title}");

        var item = list.GetItemById(listItemId);
        context.Load(item);
        await context.ExecuteQueryRetryAsync();
        Console.WriteLine($"Loaded item \"{item.FieldValues["Title"]}\"");

        var listField = list.Fields.GetByInternalNameOrTitle(termFieldName);
        var taxKeywordField = context.CastTo<TaxonomyField>(listField);

        var termValue = new TaxonomyFieldValue
        {
            Label = null,
            TermGuid = newTermGuid,
            WssId = -1
        };
        taxKeywordField.SetFieldValueByValue(item, termValue);
        taxKeywordField.Update();

        item.Update();
        await context.ExecuteQueryRetryAsync();
        Console.WriteLine($"Updated item \"{item.FieldValues["Title"]}\" with term \"{newTermGuid}\"");
    }

    public async Task CopyFileAsync(ClientRuntimeContext context, string sourceUrl, string targetUrl, bool overwriteFile, CancellationToken cancellationToken = default)
    {
        try
        {
            MoveCopyUtil.CopyFile(context, sourceUrl, targetUrl, overwriteFile, new MoveCopyOptions());
            await context.ExecuteQueryRetryAsync();
        }
        catch (Exception ex)
        {
            throw;
        }
    }
}