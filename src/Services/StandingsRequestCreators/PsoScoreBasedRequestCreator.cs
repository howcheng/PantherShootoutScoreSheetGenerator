using System.Text;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Base class for classes that create a <see cref="Request"/> to build a column based on the game results in the standings table (e.g., goals scored/conceded)
	/// </summary>
	public abstract class PsoScoreBasedStandingsRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		private Func<int, int, string, string> _formulaGeneratorMethod;

		protected PsoScoreBasedStandingsRequestCreator(FormulaGenerator formGen, string columnHeader, Func<int, int, string, string> formulaGeneratorMethod)
			: base(formGen, columnHeader)
		{
			_formulaGeneratorMethod = formulaGeneratorMethod;
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) => throw new NotImplementedException();

		public override Request CreateRequest(StandingsRequestCreatorConfig cfg)
		{
			PsoStandingsRequestCreatorConfig config = (PsoStandingsRequestCreatorConfig)cfg;

			StringBuilder sb = new StringBuilder("=");
			int startGamesRowNum = config.StartGamesRowNum;
			for (int i = 0; i < config.NumberOfRounds; i++)
			{
				int endGamesRowNum = startGamesRowNum - 1 + config.GamesPerRound; // NOTE: not the same as the EndGamesRowNum in the config object; if 2 games/round, then last row should be 4 (not 3+2, because it's cells A3:A4 inclusive)
				sb.Append(_formulaGeneratorMethod(startGamesRowNum, endGamesRowNum, config.FirstTeamsSheetCell));
				if (i < config.NumberOfRounds - 1)
					sb.Append('+');
				startGamesRowNum += config.GamesPerRound + Constants.ROUND_OFFSET_STANDINGS_TABLE; // +2 to account for the blank space and round header
			}

			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, _columnIndex, config.RowCount,
				sb.ToString());
			return request;
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for number of games played
	/// </summary>
	/// <remarks>
	/// This is different from the <see cref="GamesPlayedRequestCreator"/> because in PSO, we only have one standings table for all pool games
	/// </remarks>
	public class PsoGamesPlayedRequestCreator : PsoScoreBasedStandingsRequestCreator
	{
		public PsoGamesPlayedRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_GAMES_PLAYED, formGen.GetGamesPlayedFormula)
		{
		}
	}

	public class PsoGamesWonRequestCreator : PsoScoreBasedStandingsRequestCreator
	{
		public PsoGamesWonRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_NUM_WINS, formGen.GetGamesWonFormula)
		{
		}
	}

	public class PsoGamesLostRequestCreator : PsoScoreBasedStandingsRequestCreator
	{
		public PsoGamesLostRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_NUM_LOSSES, formGen.GetGamesLostFormula)
		{
		}
	}

	public class PsoGamesDrawnRequestCreator : PsoScoreBasedStandingsRequestCreator
	{
		public PsoGamesDrawnRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_NUM_DRAWS, formGen.GetGamesDrawnFormula)
		{
		}
	}
}
