// Licensed under the MIT license.

using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Authentication;
using TaskStatus = Microsoft.Graph.Models.TaskStatus;

namespace GraphTutorial;

public class GraphHelper
{
    // Settings object
    private readonly Settings _settings;
    // Client configured with user authentication
    private readonly GraphServiceClient _userClient;

    public GraphHelper(Settings settings)
    {
        _settings = settings;

        _userClient = GetClient(settings.TenantId, settings.ClientId);
    }

    private GraphServiceClient GetClient(string tenantId, string clientId)
    {
        // Multi-tenant apps can use "common",
        // single-tenant apps must use the tenant ID from the Azure portal
        var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(clientId, tenantId));
        return new GraphServiceClient(authenticationProvider);
    }

    public async Task<string> GetUserTokenAsync()
    {
        var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(_settings.ClientId, _settings.TenantId));
        return await authenticationProvider.AccessTokenProvider.GetAuthorizationTokenAsync(new Uri("https://graph.microsoft.com"));
    }

    public Task<User?> GetUserAsync()
    {
        return _userClient.Me.GetAsync((config) =>
        {
            // Only request specific properties
            config.QueryParameters.Select = new[] { "displayName", "mail", "userPrincipalName" };
        });
    }

    public Task<TodoTaskListCollectionResponse?> GetTodoTaskListsAsync()
    {
        return _userClient.Me
            .Todo
            .Lists
            .GetAsync((config) =>
            {
                config.QueryParameters.Top = 5;
            });
    }

    public async Task<TodoTaskListCollectionResponse?> GetTodoTaskListNextPageAsync(string odataNextLink)
    {
        var nextPageRequest = new Microsoft.Graph.Me.Todo.Lists.ListsRequestBuilder(odataNextLink, _userClient.RequestAdapter);
        return await nextPageRequest.GetAsync();
    }

    public Task<TodoTaskListCollectionResponse?> SearchTodoTaskListsAsync(string taskListName)
    {
        return _userClient.Me
            .Todo
            .Lists
            .GetAsync((config) =>
            {
                config.QueryParameters.Filter = $"contains(displayName,'{taskListName}')";
            });
    }

    public Task<TodoTaskCollectionResponse?> GetTodoTasksAsync(string listId)
    {
        return _userClient.Me
            .Todo
            .Lists[listId]
            .Tasks
            .GetAsync((config) =>
            {
                config.QueryParameters.Top = 5;
                config.QueryParameters.Orderby = new[] { "completedDateTime/dateTime" };
            });
    }

    /// <summary>
    /// Upcoming due task within 24 hours from now
    /// </summary>
    /// <returns></returns>
    public async Task<List<TodoTask>?> GetUpcomingDueTodoTasksAsync()
    {
        var taskListPage = await _userClient.Me
            .Todo
            .Lists
            .GetAsync();

        if (taskListPage?.Value == null || taskListPage.Value.Count == 0)
        {
            return null;
        }

        DateTime date = DateTime.Today;
        // Specify that the date is unspecified
        date = DateTime.SpecifyKind(date, DateTimeKind.Unspecified);
        // Convert it to UTC time
        date = date.ToUniversalTime();

        List<TodoTask>? listTodoTask = new List<TodoTask>();
        foreach (var list in taskListPage.Value)
        {
            var currentPageResult = await _userClient.Me
            .Todo
            .Lists[list.Id]
            .Tasks
            .GetAsync((config) =>
            {
                config.QueryParameters.Filter = $"dueDateTime/dateTime ge '{date.ToString("yyyy-MM-ddTHH:mm:ss")}' and dueDateTime/dateTime le '{date.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ss")}'";
                config.QueryParameters.Orderby = new[] { "completedDateTime/dateTime", "dueDateTime/dateTime" };
            });

            listTodoTask.AddRange(currentPageResult.Value);
        }

        return listTodoTask;
    }

    public async Task<TodoTaskCollectionResponse?> GetTodoTaskNextPageAsync(string odataNextLink)
    {
        var nextPageRequest = new Microsoft.Graph.Me.Todo.Lists.Item.Tasks.TasksRequestBuilder(odataNextLink, _userClient.RequestAdapter);
        return await nextPageRequest.GetAsync();
    }

    public Task<TodoTaskCollectionResponse?> SearchTodoTasksAsync(string listId, string taskTitle)
    {
        return _userClient.Me
            .Todo
            .Lists[listId]
            .Tasks
            .GetAsync((config) =>
            {
                config.QueryParameters.Top = 10;
                config.QueryParameters.Filter = $"contains(title,'{taskTitle}')";
                config.QueryParameters.Orderby = new[] { "completedDateTime/dateTime" };
            });
    }

    public async Task AddTaskToList(string listId, string title, DateTime dueDate)
    {
        var requestBody = new TodoTask
        {
            Title = title,
            DueDateTime = dueDate == DateTime.MinValue ? null : dueDate.ToDateTimeTimeZone(),
        };
        await _userClient.Me.Todo.Lists[listId].Tasks.PostAsync(requestBody);
    }

    public async Task UpdateTodoTask(string listId, string taskId, DateTime dueDate, string status)
    {
        var requestBody = new TodoTask();
        if (dueDate != DateTime.MaxValue) requestBody.DueDateTime = dueDate == DateTime.MinValue ? null : dueDate.ToDateTimeTimeZone();
        if (status == "0") requestBody.Status = TaskStatus.NotStarted;
        else if (status == "1") requestBody.Status = TaskStatus.InProgress;
        else if (status == "2") requestBody.Status = TaskStatus.Completed;
        else if (status == "3") requestBody.Status = TaskStatus.WaitingOnOthers;
        else if (status == "4") requestBody.Status = TaskStatus.Deferred;

        await _userClient.Me.Todo.Lists[listId].Tasks[taskId].PatchAsync(requestBody);
    }

    public async Task ClearCacheAsync()
    {
        await new TokenProvider(_settings.ClientId, _settings.TenantId).ClearCacheAsync();
    }
}
