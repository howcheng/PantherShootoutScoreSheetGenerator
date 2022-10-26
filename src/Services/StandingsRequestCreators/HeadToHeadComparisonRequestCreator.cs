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
		private readonly PsoFormulaGenerator _formulaGenerator;

		public HeadToHeadComparisonRequestCreator(FormulaGenerator formGen)
		{
			_formulaGenerator = (PsoFormulaGenerator)formGen;
		}

		public string ColumnHeader => string.Empty;

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
			string homeTeamFormula = _formulaGenerator.GetHeadToHeadComparisonFormulaForHomeTeam(config.StartGamesRowNum, config.EndGamesRowNum, config.FirstTeamsSheetCell, config.OpponentTeamSheetCell);
			string awayTeamFormula = _formulaGenerator.GetHeadToHeadComparisonFormulaForAwayTeam(config.StartGamesRowNum, config.EndGamesRowNum, config.FirstTeamsSheetCell, config.OpponentTeamSheetCell);

			// IF NA when looking in the home team column then use the away team column, if still NA then blank
			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, config.ColumnIndex, config.RowCount,
				$"=IFNA(IFNA({homeTeamFormula}, {awayTeamFormula}), \"\")");
			return request;
		}
	}
}
