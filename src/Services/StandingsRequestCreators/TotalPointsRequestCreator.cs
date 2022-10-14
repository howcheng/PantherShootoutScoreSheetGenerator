using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class TotalPointsRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		public TotalPointsRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_GAME_PTS)
		{
		}

		public Request CreateRequest(StandingsRequestCreatorConfig cfg)
		{
			PsoStandingsRequestCreatorConfig config = (PsoStandingsRequestCreatorConfig)cfg;
			PsoDivisionSheetHelper helper = (PsoDivisionSheetHelper)_formulaGenerator.SheetHelper;

			StringBuilder sb = new StringBuilder("=");
			int startGamesRowNum = config.StartGamesRowNum;
			for (int i = 0; i < config.NumberOfRounds; i++)
			{
				int endGamesRowNum = startGamesRowNum - 1 + config.GamesPerRound; // if 2 games/round, then last row should be 4 (not 3+2, because it's cells A3:A4 inclusive)
				string homeTeamsCellRange = Utilities.CreateCellRangeString(helper.HomeTeamColumnName, startGamesRowNum, endGamesRowNum, CellRangeOptions.FixRow);
				string awayTeamsCellRange = Utilities.CreateCellRangeString(helper.AwayTeamColumnName, startGamesRowNum, endGamesRowNum, CellRangeOptions.FixRow);
				string homePointsCellRange = Utilities.CreateCellRangeString(helper.HomeTeamPointsColumnName, startGamesRowNum, endGamesRowNum, CellRangeOptions.FixRow);
				string awayPointsCellRange = Utilities.CreateCellRangeString(helper.AwayTeamPointsColumnName, startGamesRowNum, endGamesRowNum, CellRangeOptions.FixRow);

				sb.AppendFormat("SUMIFS({1},{2},\"=\"&{0})+SUMIFS({3},{4},\"=\"&{0})",
					config.FirstTeamsSheetCell,
					homePointsCellRange,
					homeTeamsCellRange,
					awayPointsCellRange,
					awayTeamsCellRange
					);
				if (i < config.NumberOfRounds - 1)
					sb.Append("+");
				startGamesRowNum += config.GamesPerRound + 2; // +2 to account for the blank space and round header
			}

			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, _columnIndex, config.RowCount,
				sb.ToString());
			return request;
		}
	}

}
