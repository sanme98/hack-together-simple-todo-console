// Licensed under the MIT license.
using GraphTutorial;
using Microsoft.Graph.Models;

Console.WriteLine("Simple Todo Console\n");

var settings = Settings.LoadSettings();

var graphHelper = new GraphHelper(settings);

// Greet the user by name
await GreetUserAsync();

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. Clear access token cache");
    Console.WriteLine("3. Clear screen");
    Console.WriteLine("4. List All To Do Task Lists");
    Console.WriteLine("5. List All To Do Tasks");
    Console.WriteLine("6. Display To Do Task");
    Console.WriteLine("7. Add To Do Task");
    Console.WriteLine("8. Update To Do Task");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

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
            // Run clear token cache
            await ClearCacheAsync();
            break;
        case 3:
            Console.Clear();
            await GreetUserAsync();
            break;
        case 4:
            await ListTodoTaskListsAsync();
            break;
        case 5:
            await ListTodoTasksAsync();
            break;
        case 6:
            await DisplayTodoTaskAsync();
            break;
        case 7:
            await AddTodoTaskAsync();
            break;
        case 8:
            await UpdateTodoTaskAsync();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}

async Task GreetUserAsync()
{
    try
    {
        var user = await graphHelper.GetUserAsync();
        Console.WriteLine($"Hello, {user?.DisplayName}!");
        // For Work/school accounts, email is in Mail property
        // Personal accounts, email is in UserPrincipalName
        Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");

        var upcomingTasks = await graphHelper.GetUpcomingDueTodoTasksAsync();
        Console.WriteLine($"{upcomingTasks?.Count} upcoming task(s) due within two days");
        for (int i = 1; i <= upcomingTasks.Count; i++)
        {
            Console.WriteLine($"  {i}. {upcomingTasks[i - 1].Title ?? "NO Name"}" + $" Due: {upcomingTasks[i - 1].DueDateTime.ToDateTime().ToLocalTime().ToString("yyyy-MM-dd")}");
        }
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
        var userToken = await graphHelper.GetUserTokenAsync();
        Console.WriteLine($"User token: {userToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user access token: {ex.Message}");
    }
}

async Task ListTodoTaskListsAsync()
{
    try
    {
        var taskListPage = await graphHelper.GetTodoTaskListsAsync();
        ShowTodoListMessage(taskListPage);

        while (true)
        {
            var moreAvailable = !string.IsNullOrEmpty(taskListPage.OdataNextLink);

            if (moreAvailable)
            {
                Console.Write($"\nMore Todo Lists available? {moreAvailable}. Get next page (y/n)? ");
            }
            else
            {
                Console.WriteLine($"\nMore Todo Lists available? {moreAvailable}");
                break;
            }
            if (Console.ReadLine() == "y")
            {
                taskListPage = await graphHelper.GetTodoTaskListNextPageAsync(taskListPage.OdataNextLink);
                ShowTodoListMessage(taskListPage);
            }
            else
            {
                break;
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user's todo lists: {ex.Message}");
    }
}

void ShowTodoListMessage(TodoTaskListCollectionResponse taskListPage)
{
    if (taskListPage?.Value == null)
    {
        Console.WriteLine("No results returned.");
        return;
    }

    // Output each message's details
    foreach (var list in taskListPage.Value)
    {
        Console.WriteLine($"{list.DisplayName ?? "NO Name"}" + (list.WellknownListName.HasValue ? $" ({list.WellknownListName})" : ""));
    }
}

async Task ListTodoTasksAsync()
{
    try
    {
        string taskListId = await GetTaskListId();

        var taskPage = await graphHelper.GetTodoTasksAsync(taskListId);
        ShowTodoMessage(taskPage);

        while (true)
        {
            var moreAvailable = !string.IsNullOrEmpty(taskPage.OdataNextLink);

            if (moreAvailable)
            {
                Console.Write($"\nMore Todo Lists available? {moreAvailable}. Get next page (y/n)? ");
            }
            else
            {
                Console.WriteLine($"\nMore Todo Lists available? {moreAvailable}");
                break;
            }
            if (Console.ReadLine() == "y")
            {
                taskPage = await graphHelper.GetTodoTaskNextPageAsync(taskPage.OdataNextLink);
                ShowTodoMessage(taskPage);
            }
            else
            {
                break;
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user's inbox: {ex.Message}");
    }
}

void ShowTodoMessage(TodoTaskCollectionResponse taskPage)
{
    if (taskPage?.Value == null)
    {
        Console.WriteLine("No results returned.");
        return;
    }

    // Output each message's details
    foreach (var task in taskPage.Value)
    {
        Console.WriteLine($"Title: {task.Title ?? "NO Name"}");
        Console.WriteLine($"  Created DateTime: {task.CreatedDateTime.Value.ToLocalTime()}");
        Console.WriteLine($"  Due Date: {task.DueDateTime?.ToDateTime().ToLocalTime().ToString("yyyy-MM-dd")}");
        Console.WriteLine($"  Status: {task.Status}");
    }
}

async Task<string> GetTaskListId()
{
    string taskListId;
    Console.Write("Please enter Task List Name (empty for 'Tasks'): ");
    while (true)
    {
        var taskListName = Console.ReadLine();
        if (string.IsNullOrEmpty(taskListName))
        {
            taskListName = "Tasks";
        }
        var taskListPage = await graphHelper.SearchTodoTaskListsAsync(taskListName);
        if (taskListPage?.Value == null || taskListPage.Value.Count == 0)
        {
            Console.Write("No results returned. Please enter Task List Name (empty for 'Tasks'): ");
            continue;
        }
        else if (taskListPage.Value.Count > 1)
        {
            Console.WriteLine($"More than one result found. First one '{taskListPage.Value[0].DisplayName}' selected.");
        }

        taskListId = taskListPage.Value[0].Id;
        break;
    }

    return taskListId;
}

async Task AddTodoTaskAsync()
{
    try
    {
        string taskListId = await GetTaskListId();
        Console.Write("Please enter title: ");
        string title = Console.ReadLine();
        
        Console.Write("Please enter Due Date (yyyy-MM-dd HH:MM): ");
        var strDueDate = Console.ReadLine();
        DateTime dueDate;
        if (!DateTime.TryParse(strDueDate, out dueDate))
        {
            dueDate = DateTime.MinValue;
        }

        await graphHelper.AddTaskToList(taskListId, title, dueDate);

        Console.WriteLine("To do added.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error adding to do: {ex.Message}");
    }
}

async Task UpdateTodoTaskAsync()
{
    try
    {
        string taskListId = await GetTaskListId();

        string taskId, taskTitle;
        Console.Write("Please enter Task Title to search: ");
        while (true)
        {
            taskTitle = Console.ReadLine();
            var taskPage = await graphHelper.SearchTodoTasksAsync(taskListId, taskTitle);
            if (taskPage?.Value == null || taskPage.Value.Count == 0)
            {
                Console.Write("No results returned. Please enter Task Title: ");
                continue;
            }
            else if (taskPage.Value.Count > 1)
            {
                Console.WriteLine($"More than one result found. Please select one from the TOP 10.");
                for (int i = 0; i < 10 && i < taskPage.Value.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {taskPage.Value[i].Title}");
                    Console.WriteLine($"  Created DateTime: {taskPage.Value[i].CreatedDateTime.Value.ToLocalTime()}");
                    Console.WriteLine($"  Due Date: {taskPage.Value[i].DueDateTime?.ToDateTime().ToLocalTime().ToString("yyyy-MM-dd")}");
                    Console.WriteLine($"  Status: {taskPage.Value[i].Status}");
                }
                string selected = Console.ReadLine();
                int selectedIndex;
                if (!int.TryParse(selected, out selectedIndex) || selectedIndex <= 0 || selectedIndex > taskPage.Value.Count)
                {
                    Console.WriteLine("Invalid index selected!");
                    return;
                }
                taskId = taskPage.Value[selectedIndex - 1].Id;
                taskTitle = taskPage.Value[selectedIndex - 1].Title;
                break;
            }

            taskId = taskPage.Value[0].Id;
            taskTitle = taskPage.Value[0].Title;
            break;
        }

        Console.Write("Please enter new Due Date (yyyy-MM-dd HH:MM): ");
        var strDueDate = Console.ReadLine();
        DateTime dueDate;
        if (!DateTime.TryParse(strDueDate, out dueDate))
        {
            dueDate = DateTime.MaxValue;
        }

        Console.Write("Please enter status 0) NotStarted, 1) InProgress, 2) Completed, 3) WaitingOnOthers, 4) Deferred) (default: 0): ");        
        var strStatus = Console.ReadLine();
        if (!(strStatus == "1" || strStatus == "2" || strStatus == "3" || strStatus == "4"))
        {
            strStatus = "0";
        }
        await graphHelper.UpdateTodoTask(taskListId, taskId, dueDate, strStatus);

        Console.WriteLine($"To do '{taskTitle}' updated.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error adding to do: {ex.Message}");
    }
}

async Task DisplayTodoTaskAsync()
{
    try
    {
        string taskListId = await GetTaskListId();

        string taskTitle;
        TodoTask todoTask;
        Console.Write("Please enter Task Title to search: ");
        while (true)
        {
            taskTitle = Console.ReadLine();
            var taskPage = await graphHelper.SearchTodoTasksAsync(taskListId, taskTitle);
            if (taskPage?.Value == null || taskPage.Value.Count == 0)
            {
                Console.Write("No results returned. Please enter Task Title: ");
                continue;
            }
            else if (taskPage.Value.Count > 1)
            {
                Console.WriteLine($"More than one result found. Please select one from the TOP 10.");
                for (int i = 0; i < 10 && i < taskPage.Value.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {taskPage.Value[i].Title}");
                    Console.WriteLine($"  Created DateTime: {taskPage.Value[i].CreatedDateTime.Value.ToLocalTime()}");
                    Console.WriteLine($"  Due Date: {taskPage.Value[i].DueDateTime?.ToDateTime().ToLocalTime().ToString("yyyy-MM-dd")}");
                    Console.WriteLine($"  Status: {taskPage.Value[i].Status}");
                }
                string selected = Console.ReadLine();
                int selectedIndex;
                if (!int.TryParse(selected, out selectedIndex) || selectedIndex <= 0 || selectedIndex > taskPage.Value.Count)
                {
                    Console.WriteLine("Invalid index selected!");
                    return;
                }
                todoTask = taskPage.Value[selectedIndex - 1];
                break;
            }

            todoTask = taskPage.Value[0];
            break;
        }

        Console.WriteLine($"Title: {todoTask.Title}");
        Console.WriteLine($"  Created DateTime: {todoTask.CreatedDateTime.Value.ToLocalTime()}");
        Console.WriteLine($"  Due Date: {todoTask.DueDateTime?.ToDateTime().ToLocalTime().ToString("yyyy-MM-dd")}");
        Console.WriteLine($"  Start DateTime: {todoTask.StartDateTime?.ToDateTime().ToLocalTime()}");
        Console.WriteLine($"  Completed DateTime: {todoTask.CompletedDateTime?.ToDateTime().ToLocalTime()}");
        Console.WriteLine($"  Status: {todoTask.Status}");
        Console.WriteLine($"  Importance: {todoTask.Importance.Value}");
        Console.WriteLine($"  Categories: {string.Join(", ", todoTask.Categories)}");
        Console.WriteLine($"  Reminder On?: {todoTask.IsReminderOn}");
        Console.WriteLine($"  Reminder DateTime: {todoTask.ReminderDateTime?.ToDateTime().ToLocalTime()}");
        Console.WriteLine($"  Recurrence: {todoTask.Recurrence?.Pattern.Type}");
        Console.WriteLine($"  Has Attachments: {todoTask.HasAttachments}");
        Console.WriteLine($"  Note: {todoTask.Body?.Content.Trim()}");
        Console.WriteLine($"  Last Modified DateTime: {todoTask.LastModifiedDateTime.Value.ToLocalTime()}");        
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error adding to do: {ex.Message}");
    }
}

async Task ClearCacheAsync()
{
    await graphHelper.ClearCacheAsync();
}
