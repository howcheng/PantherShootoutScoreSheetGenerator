using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class HeadToHeadComparisonRequestCreatorConfig : PsoStandingsRequestCreatorConfig
	{
		/// <summary>
		/// For the head-to-head columns, we don't have proper column headers, so we need to pass in the column index
		/// </summary>
		public int ColumnIndex { get; set; }
		/// <summary>
		/// Counterpart to <see cref="ScoreBasedStandingsRequestCreatorConfig.FirstTeamsSheetCell"/>
		/// </summary>
		public string OpponentTeamSheetCell { get; set; } = string.Empty;
	}

	/// <summary>
	/// Creates <see cref="Request"/> objects for the head-to-head tiebreaker comparisons
	/// </summary>
	public class HeadToHeadComparisonRequestCreator : IStandingsRequestCreator
	{
		private readonly PsoDivisionSheetHelper Helper;

		public HeadToHeadComparisonRequestCreator(FormulaGenerator formGen)
		{
			Helper = (PsoDivisionSheetHelper)formGen.SheetHelper;
		}

		public string ColumnHeader => throw new NotImplementedException();

		/// <summary>
		/// HACK: Because the head-to-head comparison columns don't have headers, we are just going to use the class name, seeing as how
		/// this guy will be used for all of those columns
		/// </summary>
		/// <param name="columnHeader"></param>
		/// <returns></returns>
		public bool IsApplicableToColumn(string columnHeader) => columnHeader == nameof(HeadToHeadComparisonRequestCreator);

		public Request CreateRequest(StandingsRequestCreatorConfig cfg)
		{
			HeadToHeadComparisonRequestCreatorConfig config = (HeadToHeadComparisonRequestCreatorConfig)cfg;
			string homeTeamFormula = CreateFormula(config, true);
			string awayTeamFormula = CreateFormula(config, false);

			// IF NA when looking in the home team column then use the away team column, if still NA then blank
			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, config.ColumnIndex, config.RowCount,
				$"=IFNA(IFNA({homeTeamFormula}, {awayTeamFormula}), \"\")");
			return request;
		}

		private string CreateFormula(HeadToHeadComparisonRequestCreatorConfig config, bool forHomeTeam)
		{
			string homeTeamsRange = Utilities.CreateCellRangeString(Helper.HomeTeamColumnName, config.StartGamesRowNum, config.EndGamesRowNum, CellRangeOptions.FixRow);
			string awayTeamsRange = Utilities.CreateCellRangeString(Helper.AwayTeamColumnName, config.StartGamesRowNum, config.EndGamesRowNum, CellRangeOptions.FixRow);
			string winningTeamsRange = Utilities.CreateCellRangeString(Helper.WinnerColumnName, config.StartGamesRowNum, config.EndGamesRowNum, CellRangeOptions.FixRow);
			string teamNameCell = config.FirstTeamsSheetCell;
			string opponentNameCell = config.OpponentTeamSheetCell;

			// SWITCH(ArrayFormula(VLOOKUP(Shootout!A17&Shootout!A18,{A$3:A$12&D$3:D$12, P$3:P$12},2,false)), "H", "W", "A", "L", "D", "D")
			// in the row where home team and away team are the given values, get the value from the winning team column and put "W" "L" or "D" as indicated
			// swap the home and away teams when forHomeTeam == false
			return string.Format("SWITCH(ArrayFormula(VLOOKUP({0}!{1}&{0}!{2},{{{3}&{4}, {5}}}, 2, FALSE)), \"{6}\", \"{9}\", \"{7}\", \"{10}\", \"{8}\", \"{8}\")",
				ShootoutConstants.SHOOTOUT_SHEET_NAME,
				forHomeTeam ? teamNameCell : opponentNameCell,
				forHomeTeam ? opponentNameCell : teamNameCell,
				homeTeamsRange,
				awayTeamsRange,
				winningTeamsRange,
				Constants.HOME_TEAM_INDICATOR,
				Constants.AWAY_TEAM_INDICATOR,
				Constants.DRAW_INDICATOR,
				forHomeTeam ? Constants.WIN_INDICATOR : Constants.LOSS_INDICATOR,
				forHomeTeam ? Constants.LOSS_INDICATOR : Constants.WIN_INDICATOR
			);
		}
	}
}
