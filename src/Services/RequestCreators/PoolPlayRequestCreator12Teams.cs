using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class PoolPlayRequestCreator12Teams : PoolPlayRequestCreator
	{
		private readonly PsoFormulaGenerator _formulaGenerator;
		private readonly PsoDivisionSheetHelper12Teams _helper;
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;
		private readonly IPoolWinnersSortedStandingsListRequestCreator _poolWinnersRequestCreator;
		private string _divisionName = string.Empty;

		public PoolPlayRequestCreator12Teams(DivisionSheetConfig config, IScoreSheetHeadersRequestCreator headersCreator, IScoreInputsRequestCreator inputsCreator, IStandingsTableRequestCreator tableCreator
			, FormulaGenerator fg, IStandingsRequestCreatorFactory factory, ITiebreakerColumnsRequestCreator tiebreakersCreator, ISortedStandingsListRequestCreator sortedStandingsCreator
			, IPoolWinnersSortedStandingsListRequestCreator poolWinnersRequestCreator) 
			: base(config, headersCreator, inputsCreator, tableCreator, tiebreakersCreator, sortedStandingsCreator)
		{
			_formulaGenerator = (PsoFormulaGenerator)fg;
			_helper = (PsoDivisionSheetHelper12Teams)_formulaGenerator.SheetHelper;
			_requestCreatorFactory = factory;
			_poolWinnersRequestCreator = poolWinnersRequestCreator;
		}

		public override PoolPlayInfo CreatePoolPlayRequests(PoolPlayInfo info)
		{
			PoolPlayInfo12Teams ret = new PoolPlayInfo12Teams(base.CreatePoolPlayRequests(info));
			_divisionName = info.Pools!.First().First().DivisionName;

			List<UpdateRequest> updateRequests = new List<UpdateRequest>(ret.UpdateValuesRequests);
			// headers for pool winners section
			int startRowIdx = ret.ChampionshipStartRowIndex + 1;
			int startColIdx = _helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_RANK);
			UpdateRequest poolWinnersHeaderRequest = new UpdateRequest(_divisionName)
			{
				RowStart = startRowIdx,
				ColumnStart = startColIdx,
			};
			poolWinnersHeaderRequest.Rows.Add(_helper.CreateHeaderRow(PsoDivisionSheetHelper12Teams.PoolWinnersHeaderRow));
			updateRequests.Add(poolWinnersHeaderRequest);

			// headers for runners-up section
			startRowIdx += ret.Pools!.First().Count();
			
			UpdateRequest runnersUpHeaderRequest = new UpdateRequest(_divisionName)
			{
				RowStart = startRowIdx,
				ColumnStart = startColIdx,
			};
			runnersUpHeaderRequest.Rows.Add(_helper.CreateHeaderRow(PsoDivisionSheetHelper12Teams.RunnersUpHeaderRow));
			updateRequests.Add(runnersUpHeaderRequest);

			ret.UpdateValuesRequests = updateRequests;

			// formulas for pool-winners and runners-up
			ret = CreatePoolWinnersFormulasRequests(ret);

			// sorted standings table for pool winners
			ret = (PoolPlayInfo12Teams)_poolWinnersRequestCreator.CreateSortedStandingsListRequest(ret);
			return ret;
		}

		private PoolPlayInfo12Teams CreatePoolWinnersFormulasRequests(PoolPlayInfo12Teams info)
		{
			int startRowNum = info.ChampionshipStartRowIndex + 3; // start at the same row as the championship section (after the first two header rows)
			int poolCount = _config.NumberOfPools;
			int endRowNum = startRowNum + poolCount - 1;
			int startColIdx = _helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_RANK);
			info.PoolWinnersStartAndEndRowNums.Add(new Tuple<int, int>(startRowNum, endRowNum));
			UpdateRequest poolWinnersRequest = new UpdateRequest(_divisionName)
			{
				RowStart = startRowNum - 1,
				ColumnStart = startColIdx + 1, // +1 for the rank column
			};

			startRowNum = endRowNum + 2;
			endRowNum = startRowNum + poolCount - 1;
			info.PoolWinnersStartAndEndRowNums.Add(new Tuple<int, int>(startRowNum, endRowNum));
			UpdateRequest runnersUpRequest = new UpdateRequest(_divisionName)
			{
				RowStart = startRowNum - 1,
				ColumnStart = startColIdx + 1, // +1 for the rank column
			};

			int poolWinnersStartRowNum = poolWinnersRequest.RowStart + 1;
			int runnersUpStartRowNum = poolWinnersStartRowNum + poolCount + 1;

			// can't use repeated cell requests for team names/points/games played because we don't want the row numbers to auto-increment
			for (int i = 0; i < poolCount; i++)
			{
				IGrouping<string, Team> pool = info.Pools!.ElementAt(i);
				Tuple<int, int> startAndEndRows = info.StandingsStartAndEndRowNums[i];

				int standingsStartRow = startAndEndRows.Item1;
				int standingsEndRow = startAndEndRows.Item2;

				// pool-winners formulas
				GoogleSheetRow poolWinnersRow = new GoogleSheetRow();
				int currentPoolWinnersRowNum = poolWinnersStartRowNum + i;
				poolWinnersRow.AddRange(CreateFormulasExceptRank(currentPoolWinnersRowNum, standingsStartRow, standingsEndRow, true).Select(x => new GoogleSheetCell { FormulaValue = x }));
				poolWinnersRequest.Rows.Add(poolWinnersRow);

				// runners-up formulas
				GoogleSheetRow runnersUpRow = new GoogleSheetRow();
				runnersUpRow.AddRange(CreateFormulasExceptRank(runnersUpStartRowNum + i, standingsStartRow, standingsEndRow, false).Select(x => new GoogleSheetCell { FormulaValue = x }));
				runnersUpRequest.Rows.Add(runnersUpRow);
			}

			info.UpdateValuesRequests.Add(poolWinnersRequest);
			info.UpdateValuesRequests.Add(runnersUpRequest);

			// tiebreaker checkboxes
			IStandingsRequestCreator tiebreakerCreator = _requestCreatorFactory.GetRequestCreator(ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER);
			int winnersFormulaStartRowIdx = poolWinnersStartRowNum - 1;
			int runnersUpFormulaStartRowIdx = runnersUpStartRowNum - 1;
			StandingsRequestCreatorConfig tiebreakerConfig = new StandingsRequestCreatorConfig
			{
				SheetId = _config.SheetId,
				SheetStartRowIndex = winnersFormulaStartRowIdx,
				RowCount = _config.NumberOfPools,
			};
			Request poolWinnersTiebreakerRequest = tiebreakerCreator.CreateRequest(tiebreakerConfig);
			info.UpdateSheetRequests.Add(poolWinnersTiebreakerRequest);

			tiebreakerConfig.SheetStartRowIndex = runnersUpFormulaStartRowIdx;
			Request runnersUpTiebreakerRequest = tiebreakerCreator.CreateRequest(tiebreakerConfig);
			info.UpdateSheetRequests.Add(runnersUpTiebreakerRequest);

			// rank formulas
			info.UpdateSheetRequests.Add(CreateRankWithTiebreakerFormulaRequest(winnersFormulaStartRowIdx));
			info.UpdateSheetRequests.Add(CreateRankWithTiebreakerFormulaRequest(runnersUpFormulaStartRowIdx));
			return info;
		}

		private List<string> CreateFormulasExceptRank(int startRowNum, int standingsStartRowNum, int standingsEndRowNum, bool isPoolWinner)
		{
			int rank = isPoolWinner ? 1 : 2;

			string teamNameFormula = _formulaGenerator.GetPoolWinnersTeamNameFormula(startRowNum, standingsStartRowNum, standingsEndRowNum, rank);
			string pointsFormula = _formulaGenerator.GetPoolWinnersGamePointsFormula(startRowNum, standingsStartRowNum, standingsEndRowNum, rank);
			string gamesPlayedFormula = _formulaGenerator.GetPoolWinnersGamesPlayedFormula(standingsStartRowNum, standingsEndRowNum);

			return new List<string> { teamNameFormula, pointsFormula, gamesPlayedFormula };
		}

		private Request CreateRankWithTiebreakerFormulaRequest(int startRowIdx) 
		{
			IStandingsRequestCreator creator = _requestCreatorFactory.GetRequestCreator(ShootoutConstants.HDR_POOL_WINNER_RANK);
			PsoStandingsRequestCreatorConfig config = new PsoStandingsRequestCreatorConfig
			{
				SheetId = _config.SheetId,
				SheetStartRowIndex = startRowIdx,
				StartGamesRowNum = startRowIdx + 1,
				EndGamesRowNum = startRowIdx + _config.TeamsPerPool,
				RowCount = _config.NumberOfPools,
			};
			Request request = creator.CreateRequest(config);
			return request;
		}
	}
}
