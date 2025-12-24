using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the requests used to build the tiebreaker columns in the standings table; these must be done after all the standings tables have been created 
	/// because they may need to span multiple pools (in the case of a 10-team division)
	/// </summary>
	public class TiebreakerColumnsRequestCreator : ITiebreakerColumnsRequestCreator
	{
		private readonly DivisionSheetConfig _config;
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;

		public TiebreakerColumnsRequestCreator(DivisionSheetConfig config, IStandingsRequestCreatorFactory requestCreatorFactory)
		{
			_config = config;
			_requestCreatorFactory = requestCreatorFactory;
		}

		public PoolPlayInfo CreateTiebreakerRequests(PoolPlayInfo info, IEnumerable<Team> poolTeams, int startRowIndex)
		{
			List<Request> requests = new List<Request>(PsoDivisionSheetHelper.MainTiebreakerColumns.Count + 5);

			int startGamesRowNum = startRowIndex + 1; // first row in first round is 3
			Team firstTeam = poolTeams.First();
			string firstTeamSheetCell = $"{ShootoutConstants.SHOOTOUT_SHEET_NAME}!{firstTeam.TeamSheetCell}";

			// Get the score entry rows for this specific pool
			Tuple<int, int> scoreEntryStartAndEnd;
			if (info.ScoreEntryStartAndEndRowNums.ContainsKey(firstTeam.PoolName))
			{
				// Use the first and last rows from this pool's score entry rows
				List<Tuple<int, int>> poolScoreEntryRows = info.ScoreEntryStartAndEndRowNums[firstTeam.PoolName];
				int firstRow = poolScoreEntryRows.First().Item1;
				int lastRow = poolScoreEntryRows.Last().Item2;
				scoreEntryStartAndEnd = new Tuple<int, int>(firstRow, lastRow);
			}
			else
			{
				// Fall back to using the division-wide values if pool info not available
				scoreEntryStartAndEnd = new Tuple<int, int>(info.FirstScoreEntryRowNum, info.LastScoreEntryRowNum);
			}

			// tiebreaker columns
			foreach (string hdr in PsoDivisionSheetHelper.MainTiebreakerColumns)
			{
				TiebreakerRequestCreatorConfig creatorConfig = new TiebreakerRequestCreatorConfig
				{
					SheetId = _config.SheetId,
					StartGamesRowNum = startGamesRowNum,
					EndGamesRowNum = startGamesRowNum + _config.TeamsPerPool - 1,
					SheetStartRowIndex = startRowIndex,
					FirstTeamsSheetCell = firstTeamSheetCell,
					RowCount = _config.TeamsPerPool,
					StandingsStartAndEndRowNums = info.StandingsStartAndEndRowNums,
					ScoreEntryStartAndEndRowNums = scoreEntryStartAndEnd,
				};

				IStandingsRequestCreator requestCreator = _requestCreatorFactory.GetRequestCreator(hdr);
				Request request = requestCreator.CreateRequest(creatorConfig);
				requests.Add(request);
			}

			info.UpdateSheetRequests.AddRange(requests);
			return info;
		}
	}
}
