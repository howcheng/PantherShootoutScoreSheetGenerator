﻿using LINQtoCSV;
using Microsoft.Extensions.Logging;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class TeamReaderService : ITeamReaderService
	{
		private readonly AppSettings _settings;
		private readonly IFileReader _fileReader;
		private readonly ILogger _logger;

		public TeamReaderService(AppSettings settings, IFileReader fileReader, ILogger<TeamReaderService> log)
		{
			_settings = settings;
			_fileReader = fileReader;
			_logger = log;
		}

		public IDictionary<string, IEnumerable<Team>> GetTeams()
		{
			_logger.LogInformation("Getting team info");

			if (!_fileReader.DoesFileExist(_settings.PathToInputFile!))
				throw new InvalidOperationException("File does not exist!");

			using (Stream stream = _fileReader.ReadAsStream(_settings.PathToInputFile!))
			using (StreamReader reader = new StreamReader(stream))
			{
				CsvFileDescription inputFileDescription = new CsvFileDescription
				{
					SeparatorChar = ',',
					FirstLineHasColumnNames = true,
					IgnoreUnknownColumns = true,
				};
				CsvContext context = new CsvContext();
				IEnumerable<Team> teams = context.Read<Team>(reader, inputFileDescription).ToArray(); // if we don't materialize this, the division names disappear when we do the GroupBy
				IEnumerable<IGrouping<string, Team>> teamsGrouped = teams.GroupBy(x => x.DivisionName);

				_logger.LogInformation($"Found {teams.Count()} teams across {teamsGrouped.Count()} divisions");

				// divisions should be in the same order as in ShootoutConstants.DivisionNames
				var orderedTeams = new Dictionary<string, IEnumerable<Team>>();
				foreach (string division in ShootoutConstants.DivisionNames)
				{
					if (teamsGrouped.Any(g => g.Key == division))
					{
						orderedTeams[division] = teamsGrouped.First(g => g.Key == division).AsEnumerable();
					}
				}

				return orderedTeams;
			}
		}
	}
}
