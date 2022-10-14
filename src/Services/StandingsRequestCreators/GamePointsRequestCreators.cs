using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public abstract class PsoGamePointsRequestCreator : StandingsRequestCreator
	{
		public PsoGamePointsRequestCreator(FormulaGenerator formGen, string columnHeader)
			: base(formGen, columnHeader)
		{
		}

		protected Request CreateRequest(StandingsRequestCreatorConfig config, bool forHomeTeam)
		{
			PsoDivisionSheetHelper helper = (PsoDivisionSheetHelper)_formulaGenerator.SheetHelper;

			// =IFS(OR(ISBLANK(A3),P3=""),0,P3="H",6,P3="D",3,P3="A",0)+IFS(B3="",0,B3<=3,B3,B3>3,3)+IFS(C3="",0,C3=0,1,C3>0,0)
			// if team not entered yet, 0; if winner pending, 0; if home wins, 6 pts; if draw, 3 pts; if loss, 0 (can't use ISBLANK for winner cell because there's a formula there)
			// + if home goals not entered yet, 0; if <=3 home goals, +home goals; if >3 home goals, +3
			// + if away goals not entered yet, 0; if 0 away goals, +1; anything else, 0
			// for away team, just swap the home/away parts for the 2nd and 3rd sections

			int startRowNum = config.StartGamesRowNum;
			string homeTeamCell = $"{helper.HomeTeamColumnName}{startRowNum}";
			string winningTeamCell = $"{helper.WinnerColumnName}{startRowNum}";
			string homeGoalsCell = $"{helper.HomeGoalsColumnName}{startRowNum}";
			string awayGoalsCell = $"{helper.AwayGoalsColumnName}{startRowNum}";
			string gamePointsFormula = string.Format("=IFS(OR(ISBLANK({0}),{1}=\"\"),0,{1}=\"{4}\",6,{1}=\"{6}\",3,{1}=\"{5}\",0)+IFS(ISBLANK({2}),0,{2}<=3,{2},{2}>3,3)+IFS(ISBLANK({3}),0,{3}=0,1,{3}>0,0)",
				homeTeamCell,
				winningTeamCell,
				forHomeTeam ? homeGoalsCell : awayGoalsCell,
				forHomeTeam ? awayGoalsCell : homeGoalsCell,
				forHomeTeam ? Constants.HOME_TEAM_INDICATOR : Constants.AWAY_TEAM_INDICATOR,
				forHomeTeam ? Constants.AWAY_TEAM_INDICATOR : Constants.HOME_TEAM_INDICATOR,
				Constants.DRAW_INDICATOR
			);

			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, _columnIndex, config.RowCount,
				gamePointsFormula);
			return request;
		}
	}

	public class HomeGamePointsRequestCreator : PsoGamePointsRequestCreator, IStandingsRequestCreator
	{
		public HomeGamePointsRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_HOME_PTS)
		{
		}

		public Request CreateRequest(StandingsRequestCreatorConfig config)
			=> CreateRequest(config, true);
	}

	public class AwayGamePointsRequestCreator : PsoGamePointsRequestCreator, IStandingsRequestCreator
	{
		public AwayGamePointsRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_AWAY_PTS)
		{
		}

		public Request CreateRequest(StandingsRequestCreatorConfig config)
			=> CreateRequest(config, false);
	}
}
