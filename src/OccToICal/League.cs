namespace OccToICal;

internal record League
{
    public required string Id { get; set; }
    public required Uri Spreadsheet { get; set; }
    public required TimeSpan GameDuration { get; set; }
    public required ICollection<Team> Teams { get; set; }
}

internal record Team
{
    public required string Name { get; set; }

    public required ICollection<string> Synonyms { get; set; } = [];
}