using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;


public class GraphHelper
{
    private static Settings? _settings;
    private static DeviceCodeCredential? _deviceCodeCredential;
    private static GraphServiceClient? _userClient;

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt, settings.TenantId, settings.ClientId);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }
    

    public static async Task<string> GetUserTokenAsync()
    {
        // null checking
        _ = _deviceCodeCredential ??
            throw new NullReferenceException("Graph has not been initialized for user auth");
        
        _ = _settings?.GraphUserScopes ?? throw new ArgumentNullException("Argument 'scopes' cannot be null");
        
        // request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
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