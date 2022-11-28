using System.Text;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class TotalPointsRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		public TotalPointsRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_GAME_PTS)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			PsoStandingsRequestCreatorConfig config = (PsoStandingsRequestCreatorConfig)cfg;
			PsoDivisionSheetHelper helper = (PsoDivisionSheetHelper)_formulaGenerator.SheetHelper;

			StringBuilder sb = new StringBuilder("=");
			int startGamesRowNum = config.StartGamesRowNum;
			for (int i = 0; i < config.NumberOfRounds; i++)
			{
				int endGamesRowNum = startGamesRowNum + config.GamesPerRound - 1; // if 2 games/round, then last row should be 4 (not 3+2, because it's cells A3:A4 inclusive)

				sb.Append(((PsoFormulaGenerator)_formulaGenerator).GetTotalPointsFormula(startGamesRowNum, endGamesRowNum, config.FirstTeamsSheetCell));
				if (i < config.NumberOfRounds - 1)
					sb.Append('+');
				startGamesRowNum += config.GamesPerRound + 2; // +2 to account for the blank space and round header
			}

			// -1 pt for yellow cards, -2 pts for red cards
			string ycCol = helper.GetColumnNameByHeader(Constants.HDR_YELLOW_CARDS);
			string rcCol = helper.GetColumnNameByHeader(Constants.HDR_RED_CARDS);
			sb.AppendFormat("-{1}{0}-2*{2}{0}", config.StartGamesRowNum, ycCol, rcCol);
			return sb.ToString();
		}
	}
}
