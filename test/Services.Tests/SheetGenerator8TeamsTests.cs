using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoFixture.Kernel;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class SheetGenerator8TeamsTests : SheetGeneratorTests
	{
		[Fact]
		public void CanCreateChampionshipRequests()
		{
			DivisionSheetConfig config = new DivisionSheetConfig();
			config.SetupForTeams(8);
			PsoDivisionSheetHelper helper = new PsoDivisionSheetHelper(config);
			ChampionshipRequestCreator8Teams creator = new ChampionshipRequestCreator8Teams(ShootoutConstants.DIV_10UB, config, helper);

			List<Team> teams = CreateTeams(config);
			PoolPlayInfo info = new PoolPlayInfo(teams) 
			{ 
				StandingsStartAndEndRowNums = new List<Tuple<int, int>>
				{
					new Tuple<int, int>(3, 6),
					new Tuple<int, int>(16, 19)
				}
			};
			ChampionshipInfo result = creator.CreateChampionshipRequests(info);

			// expect 5 values updates: 1 for the label, 1 subheader and game input row each for the consolation and championship
			Action<UpdateRequest, string, int> assertSubheader = (rq, label, rowIdx) =>
			{
				Assert.Equal(rowIdx, rq.RowStart);
				GoogleSheetRow row = rq.Rows.Single();
				GoogleSheetCell? labelCell = row.SingleOrDefault(x => !string.IsNullOrEmpty(x.StringValue));
				Assert.NotNull(labelCell);
				Assert.StartsWith(label, labelCell!.StringValue);
				Assert.True(labelCell.Bold);
				Assert.All(row, cell => Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.SubheaderRowColor)));
			};
			Action<string, int, Tuple<int, int>> assertFormula = (f, poolRank, items) =>
			{
				string expected = string.Format("=IF(COUNTIF(F{0}:F{1},\"=3\")=4, VLOOKUP({2},{{M{0}:M{1},E{0}:E{1}}},2,FALSE), \"\")", items.Item1, items.Item2, poolRank);
				Assert.Equal(expected, f);
			};
			Action<UpdateRequest, int, int> assertScoreEntry = (rq, poolRank, rowIdx) =>
			{
				Assert.Equal(rowIdx, rq.RowStart);
				GoogleSheetRow row = rq.Rows.Single();
				Assert.Equal(4, row.Count()); // 4 cells
				Assert.Collection(row
					, cell => assertFormula(cell.FormulaValue, poolRank, info.StandingsStartAndEndRowNums.First())
					, cell => Assert.Empty(cell.StringValue)
					, cell => Assert.Empty(cell.StringValue)
					, cell => assertFormula(cell.FormulaValue, poolRank, info.StandingsStartAndEndRowNums.Last())
					);
			};
			Assert.Collection(result.UpdateValuesRequests
				, rq =>
				{
					// label row
					IEnumerable<string> labelRowValues = rq.Rows.Single().Select(x => x.StringValue);
					Assert.Single(labelRowValues.Where(x => !string.IsNullOrEmpty(x)));
					Assert.All(rq.Rows.Single(), cell => Assert.True(cell.Bold));
					Assert.All(rq.Rows.Single(), cell => Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.HeaderRowColor)));
					Assert.Equal(info.ChampionshipStartRowIndex, rq.RowStart);
				}
				, rq => assertSubheader(rq, "3RD-PLACE", info.ChampionshipStartRowIndex + 1)
				, rq => assertScoreEntry(rq, 2, info.ChampionshipStartRowIndex + 2)
				, rq => assertSubheader(rq, "CHAMPIONSHIP", info.ChampionshipStartRowIndex + 3)
				, rq => assertScoreEntry(rq, 1, info.ChampionshipStartRowIndex + 4)
				);
		}
	}
}
