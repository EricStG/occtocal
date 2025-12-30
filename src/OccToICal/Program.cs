using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using Microsoft.Extensions.Configuration;
using nietras.SeparatedValues;
using OccToICal;
using System.Data;
using System.Globalization;
using System.Web;

var config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json", false)
    .Build();

var leagues = config.GetSection("Leagues").Get<League[]>()!;

using var httpClient = new HttpClient();

var serializer = new CalendarSerializer();

var outPath = args.DefaultIfEmpty(Directory.GetCurrentDirectory()).First();

var localTimeZone = TimeZoneInfo.FindSystemTimeZoneById("America/Toronto");

foreach (var league in leagues)
{
        var leaguePath = Path.Combine(outPath, league.Id);
    Directory.CreateDirectory(leaguePath);

    var calendar = new Ical.Net.Calendar();

    var games = GetGamesAsync(league, httpClient);

    var gamesPerTeam = games.GroupBy(x => x.Team, x => x.Game);

    await foreach (var team in gamesPerTeam)
    {
        var events = team.Select(g => new CalendarEvent
        {
            Uid = $"occ-{league}-{g.StartTime}-{g.Sheet}",
            Summary = g.Title,
            Duration = Duration.FromTimeSpanExact(league.GameDuration),
            Start = new CalDateTime(g.StartTime),
            Location = "Ottawa Curling Club",
            GeographicLocation = new GeographicLocation(45.410914781505205, -75.68995690238111),
            Transparency = g.IsBye ? "TRANSPARENT" : "OPAQUE"
        }).ToArray();

        calendar.Events.AddRange(events);

        var serializedCalendar = serializer.SerializeToString(calendar);

        var teamPath = Path.Combine(leaguePath, $"{team.Key.Name}.ics");
        File.WriteAllText(teamPath, serializedCalendar);
    }

}

async IAsyncEnumerable<(Team Team, Game Game)> GetGamesAsync(League league, HttpClient httpClient)
{
    var uriBuilder = new UriBuilder(league.Spreadsheet);
    var query = HttpUtility.ParseQueryString(league.Spreadsheet.Query);
    var gid = query.Get("gid");

    var segments = league.Spreadsheet.Segments[0..^1];

    uriBuilder.Path = string.Concat(segments) + "pub";
    uriBuilder.Query = $"output=csv&gid={gid}";
    using var csvStream = await httpClient.GetStreamAsync(uriBuilder.Uri);

    // todo: clean those stupid options
    var sepOptions = new SepReaderOptions
    {
        HasHeader = false,
    };
    using var reader = Sep.Reader(_ => sepOptions).From(csvStream);
    int startYear = 0;
    int endYear = 0;
    string leagueName = "";

    bool isSchedule = false;

    var currentDate = DateTime.MinValue;

    var monthNames = CultureInfo.InvariantCulture.DateTimeFormat.AbbreviatedMonthNames.Select(x => x.ToLowerInvariant()).ToArray();

    var sheetMap = new Dictionary<int, string>();

    var day = 0;
    var month = 0;
    var year = 0;

    foreach (var row in reader)
    {
        var firstCol = row[0].ToString();
        if (string.IsNullOrWhiteSpace(leagueName))
        {
            leagueName = firstCol;
            continue;
        }

        if (startYear == 0)
        {
            var values = firstCol.Split('-', StringSplitOptions.TrimEntries);

            if (values.Length != 2)
            {
                continue;
            }

            startYear = int.Parse(values[0]);
            endYear = int.Parse(values[1]);
            continue;
        }

        if (!isSchedule)
        {
            if (firstCol.Equals("date", StringComparison.OrdinalIgnoreCase))
            {
                isSchedule = true;

                for (var i = 1; i < row.ColCount; i++)
                {
                    var sheet = row[i].ToString();
                    if (sheet.StartsWith("sheet", StringComparison.OrdinalIgnoreCase) || sheet.StartsWith("bye", StringComparison.OrdinalIgnoreCase))
                    {
                        sheetMap.Add(i, sheet);
                    }
                }
            }
            continue;
        }

        if (!string.IsNullOrWhiteSpace(firstCol))
        {
            var dateParts = firstCol.Split('-', StringSplitOptions.TrimEntries);
            day = int.Parse(dateParts[0]);
            var monthString = dateParts[1].ToLowerInvariant();
            month = Array.IndexOf(monthNames, monthString) + 1;
            year = month < 7 ? endYear : startYear;
        }

        var secondCol = row[1].ToString();
        if (string.IsNullOrWhiteSpace(secondCol))
        {
            // no time, we are done here
            break;
        }

        var time = TimeOnly.Parse(secondCol);

        var startTime = new DateTime(year, month, day, time.Hour, time.Minute, 0);
        startTime = TimeZoneInfo.ConvertTimeToUtc(startTime, localTimeZone);

        foreach (var sheet in sheetMap)
        {
            var title = row[sheet.Key].ToString();

            var team = league.Teams.Where(x => title.Contains(x.Name) || x.Synonyms.Any(y => title.Contains(y))).FirstOrDefault();

            if (team != default)
            {
                var isBye = sheet.Value.Contains("bye", StringComparison.OrdinalIgnoreCase);
                yield return (team, new Game
                {
                    IsBye = isBye,
                    Sheet = sheet.Value,
                    StartTime = startTime,
                    Title = (isBye ? "Bye - " : string.Empty) + title.Trim('"'),
                });
                break;
            }
        }
    }
}

public record Game
{
    public required bool IsBye { get; set; }
    public required string Sheet { get; set; }
    public required DateTime StartTime { get; set; }
    public required string Title { get; set; }
}