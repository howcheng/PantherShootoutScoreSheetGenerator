using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class StandingsRequestCreatorTests
	{
		private readonly Fixture _fixture;
		private readonly PsoDivisionSheetHelper _helper;
		private readonly PsoFormulaGenerator _fg;
		private const int TEAMS_PER_POOL = 4;
		private const int START_ROW_NUM = 3;
		private const int GAMES_PER_ROUND = 2;
		private const int NUM_ROUNDS = 3;

		public StandingsRequestCreatorTests()
		{
			_fixture = new Fixture();
			_helper = new PsoDivisionSheetHelper(new DivisionSheetConfig { GamesPerRound = GAMES_PER_ROUND, NumberOfRounds = NUM_ROUNDS, TeamsPerPool = TEAMS_PER_POOL });
			_fg = new PsoFormulaGenerator(_helper);
		}

		private void ValidateRequest(Request request, StandingsRequestCreatorConfig config, int columnIndex)
		{
			Assert.Equal(config.SheetId, request.RepeatCell.Range.SheetId);
			Assert.Equal(config.SheetStartRowIndex, request.RepeatCell.Range.StartRowIndex);
			Assert.Equal(config.SheetStartRowIndex + config.RowCount, request.RepeatCell.Range.EndRowIndex);
			Assert.Equal(columnIndex, request.RepeatCell.Range.StartColumnIndex);
		}

		private void ValidateFormula(Request request, StandingsRequestCreatorConfig config, string formula, int columnIndex)
		{
			ValidateRequest(request, config, columnIndex);
			Assert.NotNull(request.RepeatCell.Cell.UserEnteredValue);
			Assert.Contains(formula, request.RepeatCell.Cell.UserEnteredValue.FormulaValue); // Contains because some of the formulas may not include the = sign
		}

		#region Score-based creators (games played, wins, losses, draws, total pts, game pts)
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

			string formula = _fg.GetPsoGamePointsFormulaForHomeTeam(config.StartGamesRowNum);
			ValidateFormula(request, config, formula, _helper.GetColumnIndexByHeader(Constants.HDR_HOME_PTS));
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

			string formula = _fg.GetPsoGamePointsFormulaForAwayTeam(config.StartGamesRowNum);
			ValidateFormula(request, config, formula, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_PTS));
		}
		#endregion

		[Fact]
		public void TestRankWithTiebreakerRequestCreator()
		{
			// this formula derived from the 2021 score sheet for 14U Boys
			PsoStandingsRequestCreatorConfig config = _fixture.Build<PsoStandingsRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.With(x => x.RowCount, TEAMS_PER_POOL)
				.Create();
			StandingsRankWithTiebreakerRequestCreator creator = new StandingsRankWithTiebreakerRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			string formula = _fg.GetRankWithTiebreakerFormula(config.EndGamesRowNum);
			ValidateFormula(request, config, formula, _helper.GetColumnIndexByHeader(Constants.HDR_RANK));
		}

		[Fact]
		public void TestPoolWinnersRankWithTiebreakerRequestCreator()
		{
			// this formula derived from the 2021 score sheet for 14U Boys
			PsoStandingsRequestCreatorConfig config = _fixture.Build<PsoStandingsRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.With(x => x.RowCount, TEAMS_PER_POOL)
				.Create();
			PsoFormulaGenerator fg = new PsoFormulaGenerator(new PsoDivisionSheetHelper12Teams(DivisionSheetConfigFactory.GetForTeams(12)));
			PoolWinnersRankWithTiebreakerRequestCreator creator = new PoolWinnersRankWithTiebreakerRequestCreator(fg);
			Request request = creator.CreateRequest(config);

			string formula = fg.GetPoolWinnersRankWithTiebreakerFormula(config.StartGamesRowNum, config.StartGamesRowNum + TEAMS_PER_POOL - 1);
			ValidateFormula(request, config, formula, fg.SheetHelper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_RANK));
		}

		[Fact]
		public void TestHeadToHeadRequestCreator()
		{
			HeadToHeadComparisonRequestCreatorConfig config = _fixture.Build<HeadToHeadComparisonRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.With(x => x.EndGamesRowNum, 14)
				.With(x => x.FirstTeamsSheetCell, "Shootout!A13")
				.With(x => x.OpponentTeamSheetCell, "Shootout!A$14")
				.With(x => x.RowCount, TEAMS_PER_POOL)
				.Create();
			HeadToHeadComparisonRequestCreator creator = new HeadToHeadComparisonRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			string formula = _fg.GetHeadToHeadComparisonFormulaForHomeTeam(config.StartGamesRowNum, config.EndGamesRowNum, config.FirstTeamsSheetCell, config.OpponentTeamSheetCell);
			ValidateFormula(request, config, formula, config.ColumnIndex);
		}

		[Fact]
		public void TestHomeGamePointsRequestCreator()
		{
			StandingsRequestCreatorConfig config = _fixture.Build<StandingsRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.Create();
			HomeGamePointsRequestCreator creator = new HomeGamePointsRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			string formula = _fg.GetPsoGamePointsFormulaForHomeTeam(config.StartGamesRowNum);
			ValidateFormula(request, config, formula, _helper.GetColumnIndexByHeader(Constants.HDR_HOME_PTS));
		}

		[Fact]
		public void TestAwayGamePointsRequestCreator()
		{
			StandingsRequestCreatorConfig config = _fixture.Build<StandingsRequestCreatorConfig>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.Create();
			AwayGamePointsRequestCreator creator = new AwayGamePointsRequestCreator(_fg);
			Request request = creator.CreateRequest(config);

			string formula = _fg.GetPsoGamePointsFormulaForAwayTeam(config.StartGamesRowNum);
			ValidateFormula(request, config, formula, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_PTS));
		}

		[Fact]
		public void TestOverallRankRequestCreator()
		{
			IEnumerable<Tuple<int, int>> startAndEnd = _fixture.CreateMany<Tuple<int, int>>(2);
			RankRequestCreatorConfig10Teams config = _fixture.Build<RankRequestCreatorConfig10Teams>()
				.With(x => x.StartGamesRowNum, START_ROW_NUM)
				.With(x => x.StandingsStartAndEndRowNums, startAndEnd)
				.Create();
			PsoFormulaGenerator fg = new PsoFormulaGenerator(new PsoDivisionSheetHelper10Teams(DivisionSheetConfigFactory.GetForTeams(10)));
			OverallRankRequestCreator creator = new OverallRankRequestCreator(fg);
			Request request = creator.CreateRequest(config);

			string formula = fg.GetOverallRankFormula(START_ROW_NUM, startAndEnd);
			ValidateFormula(request, config, formula, fg.SheetHelper.GetColumnIndexByHeader(ShootoutConstants.HDR_OVERALL_RANK));
		}
	}
}
