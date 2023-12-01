using System.ComponentModel;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Net;
using System.Net.Mail;
using System.Drawing;
using System;
using System.Text.Encodings.Web;
using Newtonsoft.Json;
using Microsoft.Identity.Client;
using System.Globalization;

// This sample cosnole app does the following
// - Uses a Service Principal to get an access token for Azure DevOps
// - Please see this link on how to use Service Principal and Managed Identities to caccess Azure DevOps
// -- https://learn.microsoft.com/en-us/azure/devops/integrate/get-started/authentication/service-principal-managed-identity?view=azure-devops
// - Call an existing query to get all the GitHub related item [All GitHub Items upto [Ge]]
// - We use Query's Id [ec56ded8-96df-4df4-9c65-108edbcc3f54] to run the query through APIS
// - Please see the link below to know more about the API for ADO Queries
// -- https://learn.microsoft.com/en-us/rest/api/azure/devops/wit/wiql/get?view=azure-devops-rest-7.1&tabs=HTTP
// We then use the id's from query to get the details of each item using the folloiwng API
// -- https://learn.microsoft.com/en-us/rest/api/azure/devops/wit/work-items/get-work-items-batch?view=azure-devops-rest-7.1&tabs=HTTP
// Finally we build an HTML file with a atble insight that contains all the items

const string APP_CLIENT_ID = "";
const string APP_CLIENT_SECRET = "";
const string AD_TENANT_ID = "";
const string ADO_ORG_NAME = "";
const string ADO_PROJECT_NAME = "";
const string ADO_QUERY_ID = "";

//use the httpclient
using (var client = new HttpClient())
{
    try
    {
        // Create the Confidential Client Application
        IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(APP_CLIENT_ID)
                   .WithClientSecret(APP_CLIENT_SECRET)
                   .WithAuthority(new Uri(string.Format(CultureInfo.InvariantCulture, "https://login.microsoftonline.com/{0}", AD_TENANT_ID)))
                   .Build();

        // Use the debfualt App client id for Azure DevOps [499b84ac-1321-427f-aa17-267ca6975798] to create the scope
        string AdoAppClientID = "499b84ac-1321-427f-aa17-267ca6975798/.default";
        string[] scopes = new string[] { AdoAppClientID };

        // Now get the the Access token
        var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();

        // Build the ADO REST API Urls
        client.BaseAddress = new Uri(string.Format("https://dev.azure.com/{0}/", ADO_ORG_NAME));  //url of your organization
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
        //client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);

        // Assignt the access token
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);

        // Call the ADO API to get the list of GitHub backlog items
        HttpResponseMessage response = client.GetAsync(string.Format("{0}/_apis/wit/wiql/{1}?api-version=5.1", ADO_PROJECT_NAME, ADO_QUERY_ID)).Result;

        //check to see if we have a successful response
        if (response.IsSuccessStatusCode)
        {
            // Get the response as a string
            string rootData = response.Content.ReadAsStringAsync().Result;

            // Convert the response to C# Object using Newtonsoft.JSON library
            // This response will give us a list of all the GitHub Backlog work item Ids
            Root root = JsonConvert.DeserializeObject<Root>(rootData);

            // Now we need to do an other query to get the details that we need to each item
            // For that query, we will need the comma separated list of the all item ids and the name of the fields that we need
            string ids = "";
            foreach (var item in root.workItems)
            {
                ids = string.IsNullOrEmpty(ids) ? ids + item.id.ToString() : ids + "," + item.id.ToString();
            }

            // Here are the name of the fields that we need
            string postData = "{'ids': [" + ids + "],'fields': ['System.Id','System.Title','System.Description']}";
            var content = new StringContent(postData);

            // Build the request and call the REST API
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            response = client.PostAsync("_apis/wit/workitemsbatch?api-version=7.1-preview.1", content).Result;

            // Get the response as a string
            string details = response.Content.ReadAsStringAsync().Result;

            // Convert the response to C# Object using Newtonsoft.JSON library
            // This response will give us the requested details about all the GitHub Backlog work item Ids
            DetailedRoot detailedRoot = JsonConvert.DeserializeObject<DetailedRoot>(details);

            // Now build the HTML with an JTML table

            // Initial HTML tag and style
            string htmlData = "<HTML><style>table, th, td {border: 1px solid black;border-collapse: collapse;}</style>";
            
            // Add the P tag for Item count
            htmlData = htmlData + "<p>Total Number of Items = " + detailedRoot.value.Count.ToString() +"</p>";

            // HTML Table Headers
            string htmlTable = "<table><tr><th>Id</th><th>Title</th><th>Description</th></tr>";

            // Loop through each item to create an HTML table row for each
            foreach (var item in detailedRoot.value)
            {
                // Build HTML Table Row
                string htmlTableRow = string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>", 
                    item.fields.SystemId, item.fields.SystemTitle, item.fields.SystemDescription);
                htmlTable = htmlTable + htmlTableRow;
            }

            // HTML table closing tag
            htmlTable = htmlTable + "</table>";

            // HTML close tag
            htmlData = htmlData + htmlTable + "</HTML>";

            // Write out the file 
            File.WriteAllText(string.Format("repotable{0}.html", Guid.NewGuid()), htmlData);

            // Write the HTML data to console output
            Console.WriteLine(htmlData);

            // Press any key to close
            Console.WriteLine("Press any key to close");
            Console.Read();
        }
        else
        {
            // Write error to console if the response is not 200
            Console.WriteLine(response.StatusCode.ToString());
        }
    }
    catch(Exception ex) 
    {
        // Write error to console if there is an exception
        Console.WriteLine(ex.Message);
    }
}

// All the classes below are generated using the online JSON to C# tools
public class Column
{
    public string referenceName { get; set; }
    public string name { get; set; }
    public string url { get; set; }
}
public class Field
{
    public string referenceName { get; set; }
    public string name { get; set; }
    public string url { get; set; }
}
public class Root
{
    public string queryType { get; set; }
    public string queryResultType { get; set; }
    public DateTime asOf { get; set; }
    public List<Column> columns { get; set; }
    public List<SortColumn> sortColumns { get; set; }
    public List<WorkItem> workItems { get; set; }
}
public class SortColumn
{
    public Field field { get; set; }
    public bool descending { get; set; }
}

public class WorkItem
{
    public int id { get; set; }
    public string url { get; set; }
}

// Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
public class CommentVersionRef
{
    public int commentId { get; set; }
    public int version { get; set; }
    public string url { get; set; }
}

public class Fields
{
    [JsonProperty("System.Id")]
    public int SystemId { get; set; }

    [JsonProperty("System.Title")]
    public string SystemTitle { get; set; }

    [JsonProperty("System.Description")]
    public string SystemDescription { get; set; }
}

public class DetailedRoot
{
    public int count { get; set; }
    public List<Value> value { get; set; }
}

public class Value
{
    public int id { get; set; }
    public int rev { get; set; }
    public Fields fields { get; set; }
    public string url { get; set; }
    public CommentVersionRef commentVersionRef { get; set; }
}


