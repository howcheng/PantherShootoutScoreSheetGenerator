using Google.Apis.Sheets.v4.Data;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the requests for the standings table for a 10-team division
	/// </summary>
	public sealed class StandingsTableRequestCreator10Teams : StandingsTableRequestCreator
	{
		public StandingsTableRequestCreator10Teams(DivisionSheetConfig config, StandingsSheetHelper helper, IStandingsRequestCreatorFactory factory)
			: base(config, helper, factory)
		{
		}

		public override PoolPlayInfo CreateStandingsRequests(PoolPlayInfo info, IEnumerable<Team> poolTeams, int startRowIndex)
		{
			info = base.CreateStandingsRequests(info, poolTeams, startRowIndex);

			// rank, tiebreaker wins, and tiebreaker goal diff columns -- all need to account for both pools
			int firstRowNum = info.StandingsStartAndEndRowNums.First().Item1;
			int pool1StartIdx = firstRowNum - 1;
			int pool2StartIdx = info.StandingsStartAndEndRowNums.Last().Item1 - 1;

			CreateRankRequests(info, firstRowNum, pool1StartIdx, pool2StartIdx);

			return info;
		}

		protected override Request? CreateRequestForStandingsTableColumn(string hdr, int startGamesRowNum, int startRowIndex, int endGamesRowNum, string firstTeamSheetCell)
		{
			switch (hdr)
			{
				// skip the rank requests for now because we need both pools to be built before we can do them
				case Constants.HDR_RANK:
					return null;
				default:
					return base.CreateRequestForStandingsTableColumn(hdr, startGamesRowNum, startRowIndex, endGamesRowNum, firstTeamSheetCell);
			}
		}

		private void CreateRankRequests(PoolPlayInfo info, int firstRowNum, int pool1StartIdx, int pool2StartIdx)
		{
			Request pool1Req = CreateRankRequest(firstRowNum, pool1StartIdx);
			Request pool2Req = CreateRankRequest(firstRowNum, pool2StartIdx);

			info.UpdateSheetRequests.Add(pool1Req);
			info.UpdateSheetRequests.Add(pool2Req);
		}

		private Request CreateRankRequest(int startRowNum, int standingsRowIdx)
		{
			IStandingsRequestCreator creator = _requestCreatorFactory.GetRequestCreator(Constants.HDR_RANK);
			PsoStandingsRequestCreatorConfig config = new()
			{
				SheetId = _config.SheetId,
				StartGamesRowNum = startRowNum,
				SheetStartRowIndex = standingsRowIdx,
				GamesPerRound = _config.GamesPerRound,
				NumberOfRounds = _config.NumberOfGameRounds,
				RowCount = _config.NumberOfTeams, // this would be TeamsPerPool normally
			};
			Request req = creator.CreateRequest(config);
			return req;
		}
	}
}