
// See https://aka.ms/new-console-template for more information

using OpenQA.Selenium;
using OpenQA.Selenium.Edge;

Console.WriteLine(".NET Graph Tutorial\n");

var settings = Settings.LoadSettings();
WebDriver? edgedriver = null;

// Initialize Graph
// InitializeGraph(settings);
await GraphHelper.InitGraph(settings);

// Greet the user by name
//await GreetUserAsync();
/*var count = await GraphHelper.GetUserCountAsync();
Console.WriteLine($"Currently are {count} users registered.");*/

int choice = -1;

if (edgedriver != null) edgedriver.Close();


while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. List my inbox");
    Console.WriteLine("3. Send mail");
    Console.WriteLine("4. Other API calls");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    await GraphHelper.CheckTokenLifetimeAsync();

    switch(choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            await DisplayAccessTokenAsync();
            break;
        case 2:
            // List emails from user's inbox
            await ListInboxAsync();
            break;
        case 3:
            // Send an email message
            await SendMailAsync();
            break;
        case 4:
            // Run any Graph code
            await AdvancedApiCalls();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}

void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForUserAuth(settings,
        (info, cancel) =>
        {
            // Display the device code message to
            // the user. This tells them
            // where to go to sign in and provides the
            // code to use.
            Console.WriteLine(info.Message);
            
            // write URL & Code into file
            //WriteToFile($"{info.VerificationUri}\n{info.UserCode}");
            
            // if Selenium doesn't work
            // Process.Start(new ProcessStartInfo(info.VerificationUri.ToString()) { UseShellExecute = true });
            
            // open web browser via Selenium
            edgedriver = OpenWebBrowser(info.VerificationUri.ToString(), info.UserCode);

            return Task.FromResult(0);
        });
}

async Task GreetUserAsync()
{
    try
    {
        var user = await GraphHelper.GetUserAsync();
        Console.WriteLine($"Hello, {user?.DisplayName}!");
        // For Work/school accounts, email is in Mail property
        // Personal accounts, email is in UserPrincipalName
        Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user: {ex.Message}");
    }
}

async Task DisplayAccessTokenAsync()
{
    try
    {
        var userToken = await GraphHelper.GetUserTokenAsync();
        Console.WriteLine($"User token: {userToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user access token: {ex.Message}");
    }
}

async Task ListInboxAsync()
{
    try
    {
        var messagePage = await GraphHelper.GetInboxAsync();

        if (messagePage?.Value != null)
            foreach (var message in messagePage.Value)
            {
                Console.WriteLine($"Message: {message.Subject ?? "NO SUBJECT"}");
                Console.WriteLine($"    from: {message.From?.EmailAddress?.Name}");
                Console.WriteLine($"    Status: {(message.IsRead!.Value ? "Read" : "Unread")}");
                Console.WriteLine($"    Received: {message.ReceivedDateTime?.ToLocalTime().ToString()}");
            }
        
    }
    catch (Exception e)
    {
        Console.WriteLine($"Error getting user's inbox: {e.Message}");
    }
}

async Task SendMailAsync()
{
    try
    {
        var user = await GraphHelper.GetUserAsync();
        var userEmail = user?.Mail ?? user?.UserPrincipalName;

        if (string.IsNullOrEmpty(userEmail))
        {
            Console.WriteLine("Couldn't get your email address, canceling...");
            return;
        }

        await GraphHelper.SendMailAsync("Testing Microsoft Graph", "Hello world", userEmail);

        Console.WriteLine("Mail sent");
    }
    catch (Exception e)
    {
        Console.WriteLine($"Error sending mail: {e.Message}");
    }
}

async Task MakeGraphCallAsync()
{
    // Get calendar events
    try
    {
        var events = await GraphHelper.GetEventsAsync();

        foreach (var ev in events.Value)
        {
            Console.WriteLine($"Subject: {ev.Subject ?? "NO SUBJECT"}");
            Console.WriteLine($"    Start: {ev.Start?.DateTime}");
            Console.WriteLine($"    End: {ev.End?.DateTime}");
            Console.WriteLine($"    Location: {ev.Location?.DisplayName}");
        }
    }
    catch (Exception e)
    {
        Console.WriteLine($"Error receiving events: {e.Message}");
    }
}

async Task GetPeopleAsync()
{
    // get people user works with
    try
    {
        var people = await GraphHelper.GetPeopleAsync();
        Console.WriteLine("People I work with:");

        foreach (var person in people.Value)
        {
            Console.WriteLine($"Name: {person.DisplayName}");
            Console.WriteLine($"    Email: {person?.UserPrincipalName}");
        }
        
    }
    catch (Exception e)
    {
            Console.WriteLine($"Error receiving drive: {e.Message}");
    }
}

async Task GetContactsAsync()
{
    // get user contacts
    try
    {
        var contacts = await GraphHelper.GetContactsAsync();
        Console.WriteLine("My Contacts:");

        foreach (var contact in contacts.Value)
        {
            Console.WriteLine($"Name: {contact?.DisplayName}");
            Console.WriteLine($"    Email: {contact?.EmailAddresses?[0].Address}");
            Console.WriteLine($"    Business Address: {contact?.BusinessAddress?.CountryOrRegion}");
        }
        
    }
    catch (Exception e)
    {
        Console.WriteLine($"Error receiving drive: {e.Message}");
    }
}

async Task GetTodosAsync()
{
    // get users tasks
    try
    {
        var tasks = await GraphHelper.GetTodosAsync();
        Console.WriteLine("My Todos:");

        foreach (var tasklist in tasks)
        {
            foreach (var task in tasklist.Value)
            {
                Console.WriteLine($"TODO: {task.Title}");
                Console.WriteLine($"    Due: {task.DueDateTime?.DateTime}");
            }
            
        }
        
    }
    catch (Exception e)
    {
        Console.WriteLine($"Error receiving drive: {e.Message}");
    }
}

async Task AddTodoAsync()
{
    // add a task 
    try
    {
        Console.WriteLine("Task: ");
        string title = Console.ReadLine() ?? "temp";
        Console.Write("Date: Today + ");
        int offset = Int32.Parse(Console.ReadLine() ?? "0");

        DateTime date = DateTime.Now.AddDays(offset);

        await GraphHelper.AddTodoAsync(title, date);
    }
    catch (Exception e)
    {
        Console.WriteLine($"Error receiving drive: {e.Message}");
    }
}

void WriteToFile(string message)
{
    // write authentication url & code into file
    var fs = File.Open("C:/Users/Yannis/Documents/Peakboard/Graph/TestApp/GraphTutorial/output", FileMode.Open);
    StreamWriter sw = new StreamWriter(fs);
    sw.WriteLine(message);
    sw.Close();
    fs.Close();
}

WebDriver OpenWebBrowser(string url, string code)
{
    // open authentication browser window with Selenium 
    
    // init driver
    var driver = new EdgeDriver(@"C:\Users\YannisHartmann\OneDrive - Peakboard GmbH\MS_Graph\Edge_Driver\edgedriver_win64");
    
    // navigate to microsoft graph website
    driver.Navigate().GoToUrl(url);
    
    // input authentication code
    IWebElement textfield = driver.FindElement(By.Id("otc"));
    textfield.SendKeys(code);
    return driver;
}

async Task AdvancedApiCalls()
{
    // extends program main loop
    bool valid = false;
    while (!valid)
    {
        Console.WriteLine("Please choose one of the following options:");
        Console.WriteLine("calendar");
        Console.WriteLine("people");
        Console.WriteLine("contacts");
        Console.WriteLine("todos");
        Console.WriteLine("new todo");
        Console.WriteLine("_");

        string input = Console.ReadLine() ?? "_";

        switch (input)
        {
            case "calendar":
                await MakeGraphCallAsync();
                valid = true;
                break;
            case "people":
                await GetPeopleAsync();
                valid = true;
                break;
            case "contacts":
                await GetContactsAsync();
                valid = true;
                break;
            case "todos":
                await GetTodosAsync();
                valid = true;
                break;
            case "new todo":
                await AddTodoAsync();
                valid = true;
                break;
            case "auth": 
                //await GraphHelper.Authenticate();
                valid = true;
                break;
            case "_":
                valid = true;
                break;
            default:
                Console.WriteLine("Invalid choice");
                break;
        }
    }
}


