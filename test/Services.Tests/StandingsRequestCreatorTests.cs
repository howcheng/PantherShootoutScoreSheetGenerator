using Google.Apis.Sheets.v4.Data;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class StandingsRequestCreatorTests
	{
		// NOTE: All expected formulas were copied directly from the 2021 score sheet

		private Fixture _fixture;
		private PsoDivisionSheetHelper _helper;
		private FormulaGenerator _fg;
		private const int TEAMS_PER_POOL = 4;
		private const int START_ROW_NUM = 3;

		public StandingsRequestCreatorTests()
		{
			_fixture = new Fixture();
			_helper = new PsoDivisionSheetHelper(new DivisionSheetConfig { GamesPerRound = 2, NumberOfRounds = 3, TeamsPerPool = TEAMS_PER_POOL});
			_fg = new FormulaGenerator(_helper);
		}

		private void ValidateRequest(Request request, StandingsRequestCreatorConfig config, int columnIndex)
		{
			Assert.Equal(config.SheetId, request.RepeatCell.Range.SheetId);
			Assert.Equal(config.SheetStartRowIndex, request.RepeatCell.Range.StartRowIndex);
			Assert.Equal(config.SheetStartRowIndex + config.RowCount, request.RepeatCell.Range.EndRowIndex);
			Assert.Equal(columnIndex, request.RepeatCell.Range.StartColumnIndex);
		}

		#region Score-based creators (games played, wins, losses, draws, total pts, game pts)
		// we're not actually validating the formulas themselves as those were done in the StandingsGoogleSheetHelper solution
		private void ValidateScoreBasedFormula(Request request, PsoStandingsRequestCreatorConfig config, string formulaStartsWith)
		{
			// unlike the regular season, we only have 1 standings table, so games played, wins, losses, and draws have to calculate for all rounds at once
			string[] parts = request.RepeatCell.Cell.UserEnteredValue.FormulaValue[1..].Split('+'); // trim the leading "="
			Assert.Equal(config.NumberOfRounds * 2, parts.Length); // 1 each for home and away columns
			Assert.All(parts, p => p.StartsWith(formulaStartsWith));
		}

		[Fact]
		public void TestGamesPlayedRequestCreator()
		{
			PsoStandingsRequestCreatorConfig config = _fixture.Create<PsoStandingsRequestCreatorConfig>();
			PsoGamesPlayedRequestCreator creator = new PsoGamesPlayedRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, _helper.GetColumnIndexByHeader(Constants.HDR_GAMES_PLAYED));
			ValidateScoreBasedFormula(request, config, "COUNTIF");
		}

		[Fact]
		public void TestGamesWonRequestCreator()
		{
			PsoStandingsRequestCreatorConfig config = _fixture.Create<PsoStandingsRequestCreatorConfig>();
			PsoGamesWonRequestCreator creator = new PsoGamesWonRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, _helper.GetColumnIndexByHeader(Constants.HDR_NUM_WINS));
			ValidateScoreBasedFormula(request, config, "COUNTIF");
		}

		[Fact]
		public void TestGamesLostRequestCreator()
		{
			PsoStandingsRequestCreatorConfig config = _fixture.Create<PsoStandingsRequestCreatorConfig>();
			PsoGamesLostRequestCreator creator = new PsoGamesLostRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, _helper.GetColumnIndexByHeader(Constants.HDR_NUM_LOSSES));
			ValidateScoreBasedFormula(request, config, "COUNTIF");
		}

		[Fact]
		public void TestGamesDrawnRequestCreator()
		{
			PsoStandingsRequestCreatorConfig config = _fixture.Create<PsoStandingsRequestCreatorConfig>();
			PsoGamesDrawnRequestCreator creator = new PsoGamesDrawnRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, _helper.GetColumnIndexByHeader(Constants.HDR_NUM_DRAWS));
			ValidateScoreBasedFormula(request, config, "COUNTIF");
		}

		[Fact]
		public void TestTotalPointsRequestCreator()
		{
			PsoStandingsRequestCreatorConfig config = _fixture.Create<PsoStandingsRequestCreatorConfig>();
			TotalPointsRequestCreator creator = new TotalPointsRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, _helper.GetColumnIndexByHeader(Constants.HDR_GAME_PTS));
			ValidateScoreBasedFormula(request, config, "SUMIF");
		}

		[Fact]
		public void TestHomePointsRequestCreator()
		{
			// this formula copied from thw 2021 score sheet for 14U Boys
			PsoStandingsRequestCreatorConfig config = _fixture.Build<PsoStandingsRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.Create();
			HomeGamePointsRequestCreator creator = new HomeGamePointsRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, _helper.GetColumnIndexByHeader(Constants.HDR_HOME_PTS));

			const string expected = "=IFS(OR(ISBLANK(A3),W3=\"\"),0,W3=\"H\",6,W3=\"D\",3,W3=\"A\",0)+IFS(ISBLANK(B3),0,B3<=3,B3,B3>3,3)+IFS(ISBLANK(C3),0,C3=0,1,C3>0,0)";
			Assert.Equal(expected, request.RepeatCell.Cell.UserEnteredValue.FormulaValue);
		}

		[Fact]
		public void TestAwayPointsRequestCreator()
		{
			// this formula copied from thw 2021 score sheet for 14U Boys
			PsoStandingsRequestCreatorConfig config = _fixture.Build<PsoStandingsRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.Create();
			AwayGamePointsRequestCreator creator = new AwayGamePointsRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_PTS));

			const string expected = "=IFS(OR(ISBLANK(A3),W3=\"\"),0,W3=\"A\",6,W3=\"D\",3,W3=\"H\",0)+IFS(ISBLANK(C3),0,C3<=3,C3,C3>3,3)+IFS(ISBLANK(B3),0,B3=0,1,B3>0,0)";
			Assert.Equal(expected, request.RepeatCell.Cell.UserEnteredValue.FormulaValue);
		}
		#endregion

		[Fact]
		public void TestRankWithTiebreakerRequestCreator()
		{
			// this formula copied from the 2021 score sheet for 14U Boys
			PsoStandingsRequestCreatorConfig config = _fixture.Build<PsoStandingsRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.With(x => x.RowCount, TEAMS_PER_POOL)
				.Create();
			RankWithTiebreakerRequestCreator creator = new RankWithTiebreakerRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, _helper.GetColumnIndexByHeader(Constants.HDR_RANK));

			const string expected = "=IF(OR(N3 >= 3, COUNTIF(N$3:N$6, N3) = 1), N3, N3 + O3)";
			Assert.Equal(expected, request.RepeatCell.Cell.UserEnteredValue.FormulaValue);
		}

		[Fact]
		public void TestHeadToHeadRequestCreator()
		{
			// this formula copied from thw 2021 score sheet for 14U Boys
			HeadToHeadComparisonRequestCreatorConfig config = _fixture.Build<HeadToHeadComparisonRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.With(x => x.EndGamesRowNum, 14)
				.With(x => x.FirstTeamsSheetCell, "A13")
				.With(x => x.OpponentTeamSheetCell, "A$13")
				.With(x => x.RowCount, TEAMS_PER_POOL)
				.Create();
			HeadToHeadComparisonRequestCreator creator = new HeadToHeadComparisonRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, config.ColumnIndex);

			const string expected = "=IFNA(IFNA(SWITCH(ArrayFormula(VLOOKUP(Shootout!A13&Shootout!A$13,{A$3:A$14&D$3:D$14, W$3:W$14}, 2, FALSE)), \"H\", \"W\", \"A\", \"L\", \"D\", \"D\"), SWITCH(ArrayFormula(VLOOKUP(Shootout!A$13&Shootout!A13,{A$3:A$14&D$3:D$14, W$3:W$14}, 2, FALSE)), \"H\", \"L\", \"A\", \"W\", \"D\", \"D\")), \"\")";
			Assert.Equal(expected, request.RepeatCell.Cell.UserEnteredValue.FormulaValue);
		}

		[Fact]
		public void TestHomeGamePointsRequestCreator()
		{
			// this formula copied from thw 2021 score sheet for 14U Boys
			StandingsRequestCreatorConfig config = _fixture.Build<StandingsRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.Create();
			HomeGamePointsRequestCreator creator = new HomeGamePointsRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, _helper.GetColumnIndexByHeader(Constants.HDR_HOME_PTS));

			const string expected = "=IFS(OR(ISBLANK(A3),W3=\"\"),0,W3=\"H\",6,W3=\"D\",3,W3=\"A\",0)+IFS(ISBLANK(B3),0,B3<=3,B3,B3>3,3)+IFS(ISBLANK(C3),0,C3=0,1,C3>0,0)";
			Assert.Equal(expected, request.RepeatCell.Cell.UserEnteredValue.FormulaValue);
		}

		[Fact]
		public void TestAwayGamePointsRequestCreator()
		{
			// this formula copied from thw 2021 score sheet for 14U Boys
			StandingsRequestCreatorConfig config = _fixture.Build<StandingsRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.Create();
			AwayGamePointsRequestCreator creator = new AwayGamePointsRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			ValidateRequest(request, config, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_PTS));

			const string expected = "=IFS(OR(ISBLANK(A3),W3=\"\"),0,W3=\"A\",6,W3=\"D\",3,W3=\"H\",0)+IFS(ISBLANK(C3),0,C3<=3,C3,C3>3,3)+IFS(ISBLANK(B3),0,B3=0,1,B3>0,0)";
			Assert.Equal(expected, request.RepeatCell.Cell.UserEnteredValue.FormulaValue);
		}
	}
}
