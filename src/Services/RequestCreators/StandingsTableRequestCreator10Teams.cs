using System.Reflection.Metadata.Ecma335;
using Google.Apis.Sheets.v4.Data;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the requests for the standings table for a 10-team division
	/// </summary>
	public sealed class StandingsTableRequestCreator10Teams : StandingsTableRequestCreator
	{
		public StandingsTableRequestCreator10Teams(DivisionSheetConfig config, PsoDivisionSheetHelper helper, IStandingsRequestCreatorFactory factory) 
			: base(config, helper, factory)
		{
		}

		public override PoolPlayInfo CreateStandingsRequests(PoolPlayInfo info, IEnumerable<Team> poolTeams, int startRowIndex)
		{
			info = base.CreateStandingsRequests(info, poolTeams, startRowIndex);

			int firstRowNum = info.StandingsStartAndEndRowNums.First().Item1;
			int pool1StartIdx = firstRowNum - 1;
			int pool2StartIdx = info.StandingsStartAndEndRowNums.Last().Item1 - 1;
			Request pool1Req = CreateRankRequest(firstRowNum, pool1StartIdx);
			Request pool2Req = CreateRankRequest(firstRowNum, pool2StartIdx);

			info.UpdateSheetRequests.Add(pool1Req);
			info.UpdateSheetRequests.Add(pool2Req);

			return info;
		}

		private Request CreateRankRequest(int startRowNum, int standingsRowIdx)
		{
			IStandingsRequestCreator rankCreator = _requestCreatorFactory.GetRequestCreator(Constants.HDR_RANK);
			PsoStandingsRequestCreatorConfig config = new()
			{
				SheetId = _config.SheetId,
				StartGamesRowNum = startRowNum,
				SheetStartRowIndex = standingsRowIdx,
				GamesPerRound = _config.GamesPerRound,
				NumberOfRounds = _config.NumberOfRounds,
				RowCount = _config.NumberOfTeams, // this would be TeamsPerPool normally
			};
			Request req = rankCreator.CreateRequest(config);
			return req;
		}

		protected override Request? CreateRequestForStandingsTableColumn(string hdr, int startGamesRowNum, int startRowIndex, int endGamesRowNum, string firstTeamSheetCell)
		{
			// skip the rank request for now because we need both pools to be built before we can do it
			if (hdr == Constants.HDR_RANK)
				return null;

			return base.CreateRequestForStandingsTableColumn(hdr, startGamesRowNum, startRowIndex, endGamesRowNum, firstTeamSheetCell);
		}
	}
}
