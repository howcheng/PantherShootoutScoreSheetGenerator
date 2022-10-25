using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using Microsoft.Extensions.DependencyInjection;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class PoolPlayRequestCreator12TeamsTests : SheetGeneratorTests
	{
		private static IStandingsRequestCreatorFactory CreateStandingsRequestCreatorFactory(IEnumerable<Team> teams, DivisionSheetConfig config)
		{
			IServiceCollection services = new ServiceCollection();
			services.AddPantherShootoutServices(teams, config);
			IServiceProvider provider = services.BuildServiceProvider();
			return provider.GetRequiredService<IStandingsRequestCreatorFactory>();
		}

		[Fact]
		public void CanCreatePoolPlayRequestsFor12Teams()
		{
			Fixture fixture = new Fixture();
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(12);
			List<Team> teams = CreateTeams(config);
			PoolPlayInfo12Teams info = new PoolPlayInfo12Teams(teams);

			PsoDivisionSheetHelper12Teams helper = new PsoDivisionSheetHelper12Teams(config);
			PsoFormulaGenerator fg = new PsoFormulaGenerator(helper);

			// these mock setups are the same as for the regular PoolPlayRequestCreator test
			Mock<IScoreSheetHeadersRequestCreator> mockHeadersCreator = new Mock<IScoreSheetHeadersRequestCreator>();
			mockHeadersCreator.Setup(x => x.CreateHeaderRequests(It.IsAny<PoolPlayInfo>(), It.IsAny<string>(), It.IsAny<int>(), It.IsAny<IEnumerable<Team>>()))
				.Returns((PoolPlayInfo ppi, string poolName, int rowIdx, IEnumerable<Team> teams) => ppi);

			Mock<IScoreInputsRequestCreator> mockInputsCreator = new Mock<IScoreInputsRequestCreator>();
			mockInputsCreator.Setup(x => x.CreateScoringRequests(It.IsAny<PoolPlayInfo>(), It.IsAny<IEnumerable<Team>>(), It.IsAny<int>(), ref It.Ref<int>.IsAny))
				.Returns(new scoreInputReturns((PoolPlayInfo ppi, IEnumerable<Team> ts, int rnd, ref int idx) =>
				{
					idx += config.GamesPerRound + 2; // +2 accounts for the blank row at the end
					return ppi;
				}));

			Mock<IStandingsTableRequestCreator> mockStandingsCreator = new Mock<IStandingsTableRequestCreator>();
			mockStandingsCreator.Setup(x => x.CreateStandingsRequests(It.IsAny<PoolPlayInfo>(), It.IsAny<IEnumerable<Team>>(), It.IsAny<int>()))
				.Returns((DivisionSheetConfig cfg, PoolPlayInfo ppi, IEnumerable<Team> ts, int idx) => ppi);

			IStandingsRequestCreatorFactory factory = CreateStandingsRequestCreatorFactory(teams, config);
			PoolPlayRequestCreator12Teams creator = new PoolPlayRequestCreator12Teams(config, mockHeadersCreator.Object, mockInputsCreator.Object, mockStandingsCreator.Object, fg, factory);
			info = (PoolPlayInfo12Teams)creator.CreatePoolPlayRequests(info);

			// expect to have 4 values requests (1 header row, 1 for formulas for team name/points/games played, for both winners and runners-up)
			Assert.Equal(4, info.UpdateValuesRequests.Count);

			Action<GoogleSheetRow, List<string>, int> assertPoolWinnersHeaders = (row, hdrs, rowIdx) =>
			{
				IEnumerable<string> headerValues = row.Select(x => x.StringValue);
				headerValues.Should().BeEquivalentTo(hdrs);
			};
			Action<string, Tuple<int, int>> assertMainFormulaRowNums = (formula, startAndEnd) =>
			{
				string cellRange = Utilities.CreateCellRangeString(helper.GetColumnNameByHeader(Constants.HDR_RANK), startAndEnd.Item1, startAndEnd.Item2);
				Assert.Contains(cellRange, formula);
			};
			Action<GoogleSheetRow, bool, int> assertMainFormulas = (row, isWinners, idx) =>
			{
				Assert.Equal(config.NumberOfPools, row.Count);
				Tuple<int, int> startAndEnd = info.StandingsStartAndEndRowNums.ElementAt(idx);
				Assert.All(row, cell => assertMainFormulaRowNums(cell.FormulaValue, startAndEnd)); // we don't need to validate the formulas themselves as that's done elsewhere
			};
			Assert.Collection(info.UpdateValuesRequests
				, rq => assertPoolWinnersHeaders(rq.Rows.Single(), PsoDivisionSheetHelper12Teams.PoolWinnersHeaderRow, info.StandingsStartAndEndRowNums.First().Item1 - 1)
				, rq => assertPoolWinnersHeaders(rq.Rows.Single(), PsoDivisionSheetHelper12Teams.RunnersUpHeaderRow, info.StandingsStartAndEndRowNums.First().Item1 - 1 + config.NumberOfPools)
				, rq => Assert.Collection(rq.Rows
					, row => assertMainFormulas(row, true, 0)
					, row => assertMainFormulas(row, true, 1)
					, row => assertMainFormulas(row, true, 2)
					)
				, rq => Assert.Collection(rq.Rows
					, row => assertMainFormulas(row, false, 0)
					, row => assertMainFormulas(row, false, 1)
					, row => assertMainFormulas(row, false, 2)
					)
				);

			// expect to have 6 sheet requests: 1 for the tiebreaker checkboxes and 2 each for the rank formulas (calculated and tiebreaker) for both winners and runners-up
			Assert.Equal(6, info.UpdateSheetRequests.Count);
			Assert.All(info.UpdateSheetRequests, rq => Assert.NotNull(rq.RepeatCell));

			// confirm that they all have the correct start row index value
			int winnersStartRowIdx = info.StandingsStartAndEndRowNums.First().Item1 - 1;
			int runnerUpStartRowIdx = winnersStartRowIdx + config.NumberOfPools + 1; // the runners-up section starts immediately after the winners section, +1 for the other header row

			IEnumerable<Request> checkboxRequests = info.UpdateSheetRequests.Where(rq => rq.RepeatCell.Cell.DataValidation != null);
			Action<GridRange, bool, int> assertRangeNumbers = (range, isWinners, colIdx) =>
			{
				int rowIdx = isWinners ? winnersStartRowIdx : runnerUpStartRowIdx;
				Assert.Equal(rowIdx, range.StartRowIndex);
				Assert.Equal(rowIdx + config.NumberOfPools, range.EndRowIndex);
				Assert.Equal(colIdx, range.StartColumnIndex);
			};
			Assert.Collection(checkboxRequests
				, rq => assertRangeNumbers(rq.RepeatCell.Range, true, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER))
				, rq => assertRangeNumbers(rq.RepeatCell.Range, false, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER))
				);

			IEnumerable<Request> rankRequests = info.UpdateSheetRequests.Except(checkboxRequests);
			Assert.Collection(rankRequests
				, rq => assertRangeNumbers(rq.RepeatCell.Range, true, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_RANK))
				, rq => assertRangeNumbers(rq.RepeatCell.Range, true, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_CALC_RANK))
				, rq => assertRangeNumbers(rq.RepeatCell.Range, false, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_RANK))
				, rq => assertRangeNumbers(rq.RepeatCell.Range, false, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_CALC_RANK))
				);
		}

		// https://stackoverflow.com/questions/1068095/assigning-out-ref-parameters-in-moq
		delegate PoolPlayInfo scoreInputReturns(PoolPlayInfo ppi, IEnumerable<Team> ts, int rnd, ref int idx);
	}
}
