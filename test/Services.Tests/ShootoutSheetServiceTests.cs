using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using Microsoft.Extensions.Logging;
using NuGet.Frameworks;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class ShootoutSheetServiceTests
	{
		[Theory]
		[InlineData(true)]
		[InlineData(false)]
		public void CanCreateShootoutWinnerRequest(bool forWinner)
		{
			// create some data
			Fixture fixture = new Fixture();
			const int QUANTITY = 4;
			List<Team> teams = fixture.Build<Team>()
				.With(x => x.DivisionName, ShootoutConstants.DIV_14UB) // 14UB was what Intellisense suggested *shrug emoji*
				.CreateMany(QUANTITY)
				.ToList();

			ShootoutSheetService service = new ShootoutSheetService(Mock.Of<ISheetsClient>(), Mock.Of<ILogger<ShootoutSheetService>>());
			const int START_ROW_IDX = 1;
			Request request = service.CreateShootoutWinnerFormatRequest(teams, 12345, START_ROW_IDX, forWinner);

			ConditionalFormatRule? rule = request.AddConditionalFormatRule?.Rule;
			Assert.NotNull(rule);

			GridRange range = rule!.Ranges.Single();
			Assert.Equal(START_ROW_IDX, range.StartRowIndex);
			Assert.Equal(QUANTITY + START_ROW_IDX, range.EndRowIndex);
			Assert.Equal(0, range.StartColumnIndex);
			Assert.Equal(6, range.EndColumnIndex);

			Assert.Single(rule.BooleanRule.Condition.Values);		
			ConditionValue condition = rule.BooleanRule.Condition.Values.Single();
			Assert.Contains("$F2", condition.UserEnteredValue);
			const int START_ROW_NUM = START_ROW_IDX + 1;
			const int END_ROW_NUM = START_ROW_IDX + QUANTITY;
			Assert.Contains($"$B${START_ROW_NUM}:$D${END_ROW_NUM}", condition.UserEnteredValue);

			string additionalConditionCheck = $"$A${START_ROW_NUM}:$A${END_ROW_NUM}";
			if (forWinner)
				Assert.DoesNotContain(additionalConditionCheck, condition.UserEnteredValue);
			else
				Assert.Contains(additionalConditionCheck, condition.UserEnteredValue);
		}

		[Theory]
		[InlineData(3, 2)]
		[InlineData(3, 10)]
		public void CanCreateTeamRow(int firstTeamRowNum, int rowIndex)
		{
			Fixture fixture = new Fixture();
			Team team = fixture.Create<Team>();

			ShootoutSheetService service = new ShootoutSheetService(Mock.Of<ISheetsClient>(), Mock.Of<ILogger<ShootoutSheetService>>());
			const int TEAMS_COUNT = 8;
			GoogleSheetRow row = service.CreateTeamRow(team, firstTeamRowNum, rowIndex, TEAMS_COUNT);

			int EXPECTED_START_ROW = rowIndex + 1;
			string totalFormula = $"=SUM(B{EXPECTED_START_ROW}:D{EXPECTED_START_ROW})";
			string rankFormula = $"=RANK(E{EXPECTED_START_ROW},E${firstTeamRowNum}:E${firstTeamRowNum + TEAMS_COUNT - 1})";
			Assert.Collection(row
				, cell => Assert.Equal(team.TeamName, cell.StringValue)
				, cell => Assert.Empty(cell.StringValue)
				, cell => Assert.Empty(cell.StringValue)
				, cell => Assert.Empty(cell.StringValue)
				, cell => Assert.Equal(totalFormula, cell.FormulaValue)
				, cell => Assert.Equal(rankFormula, cell.FormulaValue)
			);

			Assert.Equal($"A{EXPECTED_START_ROW}", team.TeamSheetCell);
		}

		[Fact]
		public async Task TestGenerateSheet()
		{
			// create some data: 4 teams per division
			Fixture fixture = new Fixture();
			const int QUANTITY = 4;
			Dictionary<string, IEnumerable<Team>> allTeams = new Dictionary<string, IEnumerable<Team>>();
			foreach (string division in ShootoutConstants.DivisionNames)
			{
				allTeams.Add(division, fixture.Build<Team>().With(x => x.DivisionName, division).CreateMany(QUANTITY));
			}

			Sheet sheet = new Sheet
			{
				Properties = fixture.Build<SheetProperties>()
					.OmitAutoProperties()
					.With(x => x.SheetId, () => fixture.Create<int>())
					.Create()
			};

			Mock<ISheetsClient> mockClient = new Mock<ISheetsClient>();
			mockClient.Setup(x => x.CreateSpreadsheet(It.IsAny<string>(), It.IsAny<CancellationToken>()));
			mockClient.Setup(x => x.GetOrAddSheet(It.IsAny<string>(), It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<CancellationToken>())).ReturnsAsync(sheet);

			IList<AppendRequest>? appendRequests = null;
			Action<IList<AppendRequest>, CancellationToken> appendCallback = (rqs, ct) => appendRequests = rqs;
			mockClient.Setup(x => x.Append(It.IsAny<IList<AppendRequest>>(), It.IsAny<CancellationToken>())).Callback(appendCallback);

			IEnumerable<Request>? updateRequests = null;
			string? newSheetName = null;
			Action<IEnumerable<Request>, CancellationToken> updateNameCallback = (rqs, ct) => newSheetName = rqs.Single().UpdateSheetProperties.Properties.Title;
			mockClient.Setup(x => x.ExecuteRequests(It.Is<IEnumerable<Request>>(rqs => rqs.Count() == 1), It.IsAny<CancellationToken>())).Callback(updateNameCallback);

			Action<IEnumerable<Request>, CancellationToken> updateCallback = (rqs, ct) => updateRequests = rqs;
			mockClient.Setup(x => x.ExecuteRequests(It.Is<IEnumerable<Request>>(rqs => rqs.Count() > 1), It.IsAny<CancellationToken>())).Callback(updateCallback);

			ShootoutSheetService service = new ShootoutSheetService(mockClient.Object, Mock.Of<ILogger<ShootoutSheetService>>());
			await service.GenerateSheet(allTeams);

			// all teams should have a team sheet cell value
			Assert.All(allTeams, div => Assert.All(div.Value, team => Assert.NotEmpty(team.TeamSheetCell)));

			// verify the data updates
			Assert.NotNull(newSheetName);
			Assert.Equal(ShootoutConstants.SHOOTOUT_SHEET_NAME, newSheetName);

			Assert.NotNull(appendRequests);
			Assert.Equal(allTeams.Count, appendRequests!.Count);
			Assert.All(appendRequests!, rq => Assert.Equal(2 + QUANTITY, rq.Rows.Count)); // 1 row per team + 2 header rows

			Action<GoogleSheetRow, string> assertHeaderRow = (row, div) =>
			{
				Assert.Equal(div, row.First().StringValue);
				Assert.All(row.Skip(1), cell => // only the first cell has a value
				{
					Assert.IsType<string>(cell.CellValue);
					Assert.Empty(cell.StringValue);
				});
				Assert.All(row, cell =>
				{
					Assert.NotNull(cell.GoogleBackgroundColor);
					Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.HeaderRowColor));
				});
			};
			Action<GoogleSheetRow> assertSubheaderRow = row =>
			{
				string[] subheaderValues = row.Select(x => x.StringValue).ToArray();
				subheaderValues.Should().BeEquivalentTo(ShootoutSheetService.HeaderRowColumns);
				Assert.All(row, cell =>
				{
					Assert.NotNull(cell.GoogleBackgroundColor);
					Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.SubheaderRowColor));
				});
			};
			Action<GoogleSheetRow, int> assertTeamRow = (row, rowNum) =>
			{
				// since we tested the team row generation separately, here all we're checking is that the row numbers are correct
				// and we only need to check the row number in one formula because it's the same in the other
				string rankFormula = row.Last().FormulaValue;
				Assert.Contains($"E{rowNum}", rankFormula);
			};
			Action<AppendRequest, string, IEnumerable<Team>, int> assertDivisionRequest = (rq, div, teams, startRow) =>
			{
				assertHeaderRow(rq.Rows.ElementAt(0), div);
				assertSubheaderRow(rq.Rows.ElementAt(1));
				IEnumerable<GoogleSheetRow> teamRows = rq.Rows.Skip(2);
				Assert.Equal(teams.Count(), teamRows.Count());
				Assert.All(teamRows.Select((row, idx) => new { row, idx }), o => assertTeamRow(o.row, startRow + o.idx));
			};

			int counter = 0;
			int startRowNum = 3;
			Assert.All(appendRequests, rq =>
			{
				KeyValuePair<string, IEnumerable<Team>> pair = allTeams.ElementAt(counter);
				assertDivisionRequest(rq, pair.Key, pair.Value, startRowNum + (counter * (QUANTITY + 2)));
				counter += 1;
			});

			// verify the conditional format requests: 2 per division
			IEnumerable<Request> formatRequests = updateRequests!.Where(rq => rq.AddConditionalFormatRule != null);
			Assert.Equal(ShootoutConstants.DivisionNames.Count() * 2, formatRequests.Count());

			Queue<int> expectedRowIndices = new Queue<int>(ShootoutConstants.DivisionNames.Select((div, idx) => idx * QUANTITY + (idx + 1) * 2));

			Action<Request> assertConditionalFormatting = r => Assert.NotNull(r.AddConditionalFormatRule);
			Assert.All(formatRequests, rq => assertConditionalFormatting(rq));

			for (int i = 0; i < formatRequests.Count(); i += 2)
			{
				int expectedRowIdx = expectedRowIndices.Dequeue();
				Request req1 = formatRequests.ElementAt(i);
				Request req2 = formatRequests.ElementAt(i + 1);
				Assert.Equal(expectedRowIdx, req1.AddConditionalFormatRule.Rule.Ranges.First().StartRowIndex);
				Assert.Equal(expectedRowIdx, req2.AddConditionalFormatRule.Rule.Ranges.First().StartRowIndex);
			}

			// verify the resize requests: 5 columns
			IEnumerable<Request> resizeRequests = updateRequests!.Where(rq => rq.UpdateDimensionProperties != null);
			Assert.Equal(ShootoutSheetService.HeaderRowColumns.Count() - 1, resizeRequests.Count());
		}
	}
}