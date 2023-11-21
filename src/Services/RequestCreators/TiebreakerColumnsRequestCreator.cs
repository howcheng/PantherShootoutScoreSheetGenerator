using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface ITiebreakerColumnsRequestCreator
	{
		PoolPlayInfo CreateTiebreakerRequests(PoolPlayInfo info, IEnumerable<Team> poolTeams, int startRowIndex);
	}

	/// <summary>
	/// Creates the requests used to build the tiebreaker columns; these must be done after all the standings tables have been created because they may need
	/// to span multiple pools (in the case of a 10-team division)
	/// </summary>
	public class TiebreakerColumnsRequestCreator : ITiebreakerColumnsRequestCreator
	{
		private readonly DivisionSheetConfig _config;
		private readonly PsoFormulaGenerator _formulaGenerator;
		private readonly PsoDivisionSheetHelper _helper;
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;

		public TiebreakerColumnsRequestCreator(DivisionSheetConfig config, FormulaGenerator formGen, IStandingsRequestCreatorFactory requestCreatorFactory)
		{
			_config = config;
			_formulaGenerator = (PsoFormulaGenerator)formGen;
			_helper = (PsoDivisionSheetHelper)formGen.SheetHelper;
			_requestCreatorFactory = requestCreatorFactory;
		}

		public PoolPlayInfo CreateTiebreakerRequests(PoolPlayInfo info, IEnumerable<Team> poolTeams, int startRowIndex)
		{
			List<Request> requests = new List<Request>(PsoDivisionSheetHelper.MainTiebreakerColumns.Count + 5);

			int startGamesRowNum = startRowIndex + 1; // first row in first round is 3
			Team firstTeam = poolTeams.First();
			string firstTeamSheetCell = $"{ShootoutConstants.SHOOTOUT_SHEET_NAME}!{firstTeam.TeamSheetCell}";

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
					ScoreEntryStartAndEndRowNums = info.ScoreEntryStartAndEndRowNums,
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
