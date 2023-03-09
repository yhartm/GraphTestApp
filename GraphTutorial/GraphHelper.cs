using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Newtonsoft.Json;


public class GraphHelper
{
    private static Settings? _settings;
    private static UsernamePasswordCredential? _credential;
    private static GraphServiceClient? _userClient;
    private static HttpClient _httpClient;
    private static AuthProvider provider;
    
    private static string username = "o365_test@peakboard.com";
    private static string password = "I3oP5Q%J1im0qOzY";

    private static string accessToken;
    private static string refreshToken;
    private static string tokenLifetime;
    private static long millis;

    private const string AUTHORIZATION_URL = "https://login.microsoftonline.com/{0}/oauth2/v2.0/devicecode";  
    private const string ALL_SCOPE_AUTHORIZATIONS = "user.read offline_access";
    private const string TOKEN_ENDPOINT_URL = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token";

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        /*_credential = new DeviceCodeCredential(
            deviceCodePrompt,
            settings.TenantId, 
            settings.ClientId);*/

        _credential = new UsernamePasswordCredential(username, password, settings.TenantId, settings.ClientId);
        
        _userClient = new GraphServiceClient(_credential, settings.GraphUserScopes);

    }

    public static async Task InitGraph(Settings settings)
    {
        _settings = settings;
        _httpClient = new HttpClient();
        
        // authorize
        string deviceCode = await AuthorizeAsync();
        Console.WriteLine("Press enter to proceed after authentication");
        Console.ReadLine();
        
        // get tokens
        await GetTokensAsync(deviceCode);
        
        // init Authentication Provider
        provider = new AuthProvider(accessToken, _settings.GraphUserScopes);
        _userClient = new GraphServiceClient(_httpClient, provider);

    }

    private static async Task<string> AuthorizeAsync()
    {
        // generate url for http request
        string url = string.Format(AUTHORIZATION_URL, _settings.TenantId);
        
        // generate body for http request
        Dictionary<string, string> values = new Dictionary<string, string>
        {
            {"client_id", _settings.ClientId},
            {"scope", ALL_SCOPE_AUTHORIZATIONS}
        };
        
        FormUrlEncodedContent data = new FormUrlEncodedContent(values);
        
        // make http request to get device code for authentication
        HttpResponseMessage response = await _httpClient.PostAsync(url, data);
        string jsonString = await response.Content.ReadAsStringAsync();
        var authorizationResponse = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);
        
        // get device code and authentication message
        authorizationResponse.TryGetValue("device_code", out var deviceCode);
        authorizationResponse.TryGetValue("message", out var message);
        
        Console.WriteLine(message);

        return deviceCode;
    }

    private static async Task GetTokensAsync(string deviceCode)
    {
        // generate url for http request
        string url = string.Format(TOKEN_ENDPOINT_URL, _settings.TenantId);
        
        // generate body for http request
        var values = new Dictionary<string, string>
        {
            { "grant_type", "urn:ietf:params:oauth:grant-type:device_code" },
            { "client_id", _settings.ClientId },
            { "device_code", deviceCode }
        };
        
        FormUrlEncodedContent data = new FormUrlEncodedContent(values);
        
        // make http request to get access token and refresh token
        HttpResponseMessage response = await _httpClient.PostAsync(url, data);
        string jsonString = await response.Content.ReadAsStringAsync();
        var tokenResponse = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);
        
        // store token values
        tokenResponse.TryGetValue("refresh_token", out refreshToken);
        tokenResponse.TryGetValue("access_token", out accessToken);
        tokenResponse.TryGetValue("expires_in", out tokenLifetime);
        millis = DateTimeOffset.Now.ToUnixTimeMilliseconds();
    }

    private static async Task RefreshTokensAsync()
    {
        // generate url for http request
        string url = String.Format(TOKEN_ENDPOINT_URL, _settings.TenantId);
        
        // generate body for http requestd
        var values = new Dictionary<string, string>
        {
            { "client_id", _settings.ClientId },
            { "grant_type", "refresh_token" },
            { "scope", ALL_SCOPE_AUTHORIZATIONS },
            { "refresh_token", refreshToken }
        };

        var data = new FormUrlEncodedContent(values);
        
        // make http request to get new tokens
        HttpResponseMessage response = await _httpClient.PostAsync(url, data);
        string jsonString = await response.Content.ReadAsStringAsync();
        var tokenResponse = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);
        
        tokenResponse.TryGetValue("access_token", out accessToken);
        tokenResponse.TryGetValue("refresh_token", out refreshToken);
        tokenResponse.TryGetValue("expires_in", out tokenLifetime);
        millis = DateTimeOffset.Now.ToUnixTimeMilliseconds();
        
        // refresh access token in auth-provider
        provider.Token = accessToken;

    }

    public static async Task CheckTokenLifetimeAsync()
    {
        long temp = DateTimeOffset.Now.ToUnixTimeMilliseconds();
        if (temp - millis > Int32.Parse(tokenLifetime))
        {
            Console.Write("Refreshing Tokens...");
            await RefreshTokensAsync();
            Console.WriteLine("Done!");
        }
    }


    /*public static async Task Authenticate()
    {
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        await _credential.AuthenticateAsync(context);
    }*/

    public static async Task<string> GetUserTokenAsync()
    {
        /*/ null checking
        _ = _credential ??
            throw new NullReferenceException("Graph has not been initialized for user auth");
        
        _ = _settings?.GraphUserScopes ?? throw new ArgumentNullException("Argument 'scopes' cannot be null");
        
        // request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);


        var response = await _credential.GetTokenAsync(context);*/
        return accessToken;
    }

    public static Task<User?> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new NullReferenceException("Graph has not been initialized for user auth");
        

        return _userClient.Me
            .GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Select = new string []{ "displayName", "mail", "userPrincipalMail" };
            });
            
    }

    public static Task<MessageCollectionResponse?> GetInboxAsync()
    {
        // Ensure client isn't null
        _ = _userClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            .MailFolders["Inbox"]
            .Messages
            .GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Select = new string[] { "from", "isRead", "receivedDateTime", "subject" };
                requestConfiguration.QueryParameters.Top = 25;
                requestConfiguration.QueryParameters.Orderby = new string[] { "receivedDateTime desc" };
            });
    }

    public static async Task SendMailAsync(string subject, string body, string recipient)
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new NullReferenceException("Graph has not been initialized for user auth");
        
        // create request body
        var requestBody = new SendMailPostRequestBody
        {
            Message = new Message
            {
                Subject = subject,
                Body = new ItemBody()
                {
                    Content = body,
                    ContentType = BodyType.Text
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient
                        }
                    }
                }
            }
        };
        
        // send mail
        await _userClient.Me
            .SendMail
            .PostAsync(requestBody);
    }

    public static Task<EventCollectionResponse?> GetEventsAsync()
    {
        // Ensure client isn't null
        _ = _userClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");
        
        // get Dates for request config
        string start = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssZ");
        string end = DateTime.Now.AddDays(7).ToString("yyyy-MM-ddThh:mm:ssZ");
        
        return _userClient.Me
                .CalendarView
                .GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Select =
                        new string[] { "subject", "start", "end", "location" };
                    requestConfiguration.QueryParameters.StartDateTime = start;
                    requestConfiguration.QueryParameters.EndDateTime = end;
                });
    }

    public static Task<PersonCollectionResponse?> GetPeopleAsync()
    {
        // Ensure client isn't null
        _ = _userClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me.People.GetAsync();
    }

    public static Task<ContactCollectionResponse?> GetContactsAsync()
    {
        // Ensure client isn't null
        _ = _userClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            .Contacts
            .GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Select = new string[]
                    { "displayName", "emailAddresses", "businessAddress" };
            });
    }

    public static Task<TodoTaskListCollectionResponse?> GetTaskListsAsync()
    {
        // Ensure client isn't null
        _ = _userClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");
    
        // get task lists
        return _userClient.Me.Todo.Lists.GetAsync();
    }

    private static Task<TodoTaskCollectionResponse?> GetTasksAsync(string id)
    {
        // Ensure client isn't null
        _ = _userClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");
        
        // get tasks of specified task list
        return _userClient.Me.Todo.Lists[id].Tasks.GetAsync((requestConfiguration) =>
        {
            requestConfiguration.QueryParameters.Orderby = new[] { "dueDateTime/dateTime" };
        });
    }

    public static async Task<List<TodoTaskCollectionResponse>> GetTodosAsync()
    {
        // get TaskListCollection
        var lists = await GetTaskListsAsync();
        List<TodoTaskCollectionResponse> tasks = new List<TodoTaskCollectionResponse>();
        
        // iterate over TaskListCollection and get TaskCollections of each list
        foreach (var list in lists.Value)
        {
            var task = await GetTasksAsync(list.Id);
            
            // add TaskCollections into result list
            tasks.Add(task);
        }
        
        // return list of all TaskCollections
        return tasks;
    }
    
    public static async Task AddTodoAsync(string title, DateTime date)
    {
        // Ensure client isn't null
        _ = _userClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");
        
        string id = null;
        var lists = await GetTaskListsAsync();
        // Search for TaskList with name 'Aufgaben' and get its ID
        foreach (var list in lists.Value)
        {
            if (list.DisplayName == "Aufgaben")
            {
                id = list.Id;
                break;
            }
        }
        
        // create request body
        var requestBody = new TodoTask
        {
            Title = title,
            DueDateTime = new DateTimeTimeZone
            {
                DateTime = date.ToString("yyyy-MM-ddThh:mm:ssZ"),
                TimeZone = "UTC"
            }
        };
        
        // Add Task to TaskList 'Aufgaben' if it exists respectively if ID isn't null
        if (id != null)
        {
            await _userClient.Me.Todo.Lists[id].Tasks.PostAsync(requestBody);
        }
        

    }

    public static Task<int?> GetUserCountAsync()
    {
        // Ensure client isn't null
        _ = _userClient ?? throw new NullReferenceException("Graph has not been initialized for user auth");
        
        return _userClient.Users.Count.GetAsync((requestConfiguration) =>
        {
            requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
        });
    }
     
}