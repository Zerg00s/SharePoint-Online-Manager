namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents metadata about a SharePoint list or library.
/// </summary>
public class ListInfo
{
    public Guid Id { get; set; }
    public string Title { get; set; } = string.Empty;
    public string ServerRelativeUrl { get; set; } = string.Empty;
    public int ItemCount { get; set; }
    public bool Hidden { get; set; }
    public DateTime Created { get; set; }
    public DateTime LastItemModifiedDate { get; set; }
    public int BaseTemplate { get; set; }

    /// <summary>
    /// Gets a human-readable list type based on BaseTemplate.
    /// </summary>
    public string ListType => BaseTemplate switch
    {
        100 => "Custom List",
        101 => "Document Library",
        102 => "Survey",
        103 => "Links",
        104 => "Announcements",
        105 => "Contacts",
        106 => "Calendar",
        107 => "Tasks",
        108 => "Discussion Board",
        109 => "Picture Library",
        110 => "Data Sources",
        115 => "Form Library",
        118 => "Wiki Page Library",
        119 => "Custom Workflow Process",
        120 => "Custom Workflow History",
        130 => "Data Connection Library",
        140 => "Workflow History",
        150 => "Gantt Tasks",
        170 => "Promoted Links",
        171 => "App Catalog",
        175 => "Asset Library",
        432 => "Issues List",
        544 => "Facility",
        600 => "External List",
        851 => "Site Pages Library",
        _ => $"List ({BaseTemplate})"
    };

    /// <summary>
    /// Gets the absolute URL for the list given a site URL.
    /// </summary>
    public string GetAbsoluteUrl(string siteUrl)
    {
        var uri = new Uri(siteUrl);
        return $"{uri.Scheme}://{uri.Host}{ServerRelativeUrl}";
    }
}
