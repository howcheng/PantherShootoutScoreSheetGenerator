using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using Microsoft.Extensions.Logging;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class ShootoutSheetServiceTests
	{
		private const int TEST_TEAMS_PER_DIVISION = 4;

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
			string totalFormula = $"=SUM(B{EXPECTED_START_ROW}:E{EXPECTED_START_ROW})";
			Assert.Collection(row
				, cell => Assert.Equal(team.TeamName, cell.StringValue)
				, cell => Assert.Empty(cell.StringValue)
				, cell => Assert.Empty(cell.StringValue)
				, cell => Assert.Empty(cell.StringValue)
				, cell => Assert.Empty(cell.StringValue)
				, cell => Assert.Equal(totalFormula, cell.FormulaValue)
			);

			Assert.Equal($"A{EXPECTED_START_ROW}", team.TeamSheetCell);
		}

		[Fact]
		public async Task GenerateSheet_SetsTeamSheetCellForAllTeams()
		{
			// Arrange
			var (mockClient, allTeams) = SetupMockClientAndTeams();
			ShootoutSheetService service = new ShootoutSheetService(mockClient.Object, Mock.Of<ILogger<ShootoutSheetService>>());

			// Act
			await service.GenerateSheet(allTeams);

			// Assert
			Assert.All(allTeams, div => Assert.All(div.Value, team => Assert.NotEmpty(team.TeamSheetCell)));
		}

		[Fact]
		public async Task GenerateSheet_RenamesSheetToShootoutSheetName()
		{
			// Arrange
			var (mockClient, allTeams) = SetupMockClientAndTeams();
			List<IEnumerable<Request>> allCapturedRequests = new List<IEnumerable<Request>>();
			mockClient.Setup(x => x.ExecuteRequests(It.IsAny<IEnumerable<Request>>(), It.IsAny<CancellationToken>()))
				.Callback<IEnumerable<Request>, CancellationToken>((rqs, ct) => allCapturedRequests.Add(rqs));

			ShootoutSheetService service = new ShootoutSheetService(mockClient.Object, Mock.Of<ILogger<ShootoutSheetService>>());

			// Act
			await service.GenerateSheet(allTeams);

			// Assert
			// Find the single-request batch that renames the sheet
			IEnumerable<Request>? renameRequest = allCapturedRequests.FirstOrDefault(batch => 
				batch.Count() == 1 && batch.First().UpdateSheetProperties != null);
			Assert.NotNull(renameRequest);
			Assert.Equal(ShootoutConstants.SHOOTOUT_SHEET_NAME, renameRequest!.First().UpdateSheetProperties.Properties.Title);
		}

		[Fact]
		public async Task GenerateSheet_CreatesCorrectNumberOfAppendRequests()
		{
			// Arrange
			var (mockClient, allTeams) = SetupMockClientAndTeams();
			IList<AppendRequest>? capturedAppendRequests = null;
			mockClient.Setup(x => x.Append(It.IsAny<IList<AppendRequest>>(), It.IsAny<CancellationToken>()))
				.Callback<IList<AppendRequest>, CancellationToken>((rqs, ct) => capturedAppendRequests = rqs);

			ShootoutSheetService service = new ShootoutSheetService(mockClient.Object, Mock.Of<ILogger<ShootoutSheetService>>());

			// Act
			await service.GenerateSheet(allTeams);

			// Assert
			Assert.NotNull(capturedAppendRequests);
			Assert.Equal(allTeams.Count, capturedAppendRequests!.Count);
			Assert.All(capturedAppendRequests!, rq => Assert.Equal(2 + TEST_TEAMS_PER_DIVISION, rq.Rows.Count)); // 1 row per team + 2 header rows
		}

		[Fact]
		public async Task GenerateSheet_CreatesHeaderRowsWithCorrectFormatting()
		{
			// Arrange
			var (mockClient, allTeams) = SetupMockClientAndTeams();
			IList<AppendRequest>? capturedAppendRequests = null;
			mockClient.Setup(x => x.Append(It.IsAny<IList<AppendRequest>>(), It.IsAny<CancellationToken>()))
				.Callback<IList<AppendRequest>, CancellationToken>((rqs, ct) => capturedAppendRequests = rqs);

			ShootoutSheetService service = new ShootoutSheetService(mockClient.Object, Mock.Of<ILogger<ShootoutSheetService>>());

			// Act
			await service.GenerateSheet(allTeams);

			// Assert
			Assert.NotNull(capturedAppendRequests);
			foreach (var (request, division) in capturedAppendRequests!.Zip(allTeams.Keys))
			{
				GoogleSheetRow headerRow = request.Rows.First();
				
				// Verify header row content
				Assert.Equal(division, headerRow.First().StringValue);
				Assert.All(headerRow.Skip(1), cell =>
				{
					Assert.IsType<string>(cell.CellValue);
					Assert.Empty(cell.StringValue);
				});

				// Verify header row formatting
				Assert.All(headerRow, cell =>
				{
					Assert.NotNull(cell.GoogleBackgroundColor);
					Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.HeaderRowColor));
				});
			}
		}

		[Fact]
		public async Task GenerateSheet_CreatesSubheaderRowsWithCorrectFormatting()
		{
			// Arrange
			var (mockClient, allTeams) = SetupMockClientAndTeams();
			IList<AppendRequest>? capturedAppendRequests = null;
			mockClient.Setup(x => x.Append(It.IsAny<IList<AppendRequest>>(), It.IsAny<CancellationToken>()))
				.Callback<IList<AppendRequest>, CancellationToken>((rqs, ct) => capturedAppendRequests = rqs);

			ShootoutSheetService service = new ShootoutSheetService(mockClient.Object, Mock.Of<ILogger<ShootoutSheetService>>());

			// Act
			await service.GenerateSheet(allTeams);

			// Assert
			Assert.NotNull(capturedAppendRequests);
			Assert.All(capturedAppendRequests!, request =>
			{
				GoogleSheetRow subheaderRow = request.Rows.ElementAt(1);
				
				// Verify subheader row content
				string[] subheaderValues = subheaderRow.Select(x => x.StringValue).ToArray();
				subheaderValues.Should().BeEquivalentTo(ShootoutSheetHelper.HeaderRowColumns3Rounds);

				// Verify subheader row formatting
				Assert.All(subheaderRow, cell =>
				{
					Assert.NotNull(cell.GoogleBackgroundColor);
					Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.SubheaderRowColor));
				});
			});
		}

		[Fact]
		public async Task GenerateSheet_CreatesTeamRowsWithCorrectFormulas()
		{
			// Arrange
			var (mockClient, allTeams) = SetupMockClientAndTeams();
			IList<AppendRequest>? capturedAppendRequests = null;
			mockClient.Setup(x => x.Append(It.IsAny<IList<AppendRequest>>(), It.IsAny<CancellationToken>()))
				.Callback<IList<AppendRequest>, CancellationToken>((rqs, ct) => capturedAppendRequests = rqs);

			ShootoutSheetService service = new ShootoutSheetService(mockClient.Object, Mock.Of<ILogger<ShootoutSheetService>>());

			// Act
			await service.GenerateSheet(allTeams);

			// Assert
			Assert.NotNull(capturedAppendRequests);
			int startRowNum = 3;
			for (int divisionIndex = 0; divisionIndex < capturedAppendRequests!.Count; divisionIndex++)
			{
				AppendRequest request = capturedAppendRequests[divisionIndex];
				IEnumerable<GoogleSheetRow> teamRows = request.Rows.Skip(2);
				
				Assert.All(teamRows.Select((row, idx) => new { row, idx }), o =>
				{
					int rowNum = startRowNum + (divisionIndex * (TEST_TEAMS_PER_DIVISION + 2)) + o.idx;
					string rankFormula = o.row.Last().FormulaValue;
					Assert.Contains($"E{rowNum}", rankFormula);
				});
			}
		}

		[Fact]
		public async Task GenerateSheet_CreatesColumnResizeRequests()
		{
			// Arrange
			var (mockClient, allTeams) = SetupMockClientAndTeams();
			List<IEnumerable<Request>> allCapturedRequests = new List<IEnumerable<Request>>();
			mockClient.Setup(x => x.ExecuteRequests(It.IsAny<IEnumerable<Request>>(), It.IsAny<CancellationToken>()))
				.Callback<IEnumerable<Request>, CancellationToken>((rqs, ct) => allCapturedRequests.Add(rqs));

			ShootoutSheetService service = new ShootoutSheetService(mockClient.Object, Mock.Of<ILogger<ShootoutSheetService>>());

			// Act
			await service.GenerateSheet(allTeams);

			// Assert
			Assert.NotEmpty(allCapturedRequests);
			
			// Find the batch that contains the resize requests (the one that's not a single rename request)
			IEnumerable<Request>? resizeRequestsBatch = allCapturedRequests.FirstOrDefault(batch => 
				batch.Count() > 1 && batch.All(r => r.UpdateDimensionProperties != null));
			Assert.NotNull(resizeRequestsBatch);
			
			// Should have resize requests for all columns except the first (team name)
			Assert.Equal(ShootoutSheetHelper.HeaderRowColumns4Rounds.Length - 1, resizeRequestsBatch!.Count());
			Assert.All(resizeRequestsBatch, r =>
			{
				Assert.NotNull(r.UpdateDimensionProperties);
				Assert.Equal("COLUMNS", r.UpdateDimensionProperties.Range.Dimension);
			});
		}

		[Fact]
		public async Task GenerateSheet_ReturnsConfigWithCorrectShootoutRowNumbers()
		{
			// Arrange
			var (mockClient, allTeams) = SetupMockClientAndTeams();
			ShootoutSheetService service = new ShootoutSheetService(mockClient.Object, Mock.Of<ILogger<ShootoutSheetService>>());

			// Act
			ShootoutSheetConfig config = await service.GenerateSheet(allTeams);

			// Assert
			Assert.Equal(allTeams.Count, config.ShootoutStartAndEndRows.Count());
			
			IEnumerable<int> expectedRowIndices = ShootoutConstants.DivisionNames.Select((div, idx) => idx * TEST_TEAMS_PER_DIVISION + (idx + 1) * 2);
			for (int i = 0; i < config.ShootoutStartAndEndRows.Count(); i++)
			{
				int expectedRowIdx = expectedRowIndices.ElementAt(i);
				int startRow = expectedRowIdx + 1;
				int endRow = expectedRowIdx + TEST_TEAMS_PER_DIVISION;
				Tuple<int, int> startAndEnd = config.ShootoutStartAndEndRows.ElementAt(i).Value;
				Assert.Equal(startRow, startAndEnd.Item1);
				Assert.Equal(endRow, startAndEnd.Item2);
			}
		}

		[Fact]
		public async Task HideHelperColumns_CreatesGroupAndHidesColumns()
		{
			// Arrange
			Fixture fixture = new Fixture();
			Dictionary<string, IEnumerable<Team>> allTeams = new Dictionary<string, IEnumerable<Team>>();
			foreach (string division in ShootoutConstants.DivisionNames)
			{
				allTeams.Add(division, fixture.Build<Team>().With(x => x.DivisionName, division).CreateMany(TEST_TEAMS_PER_DIVISION));
			}

			ShootoutSheetConfig config = new ShootoutSheetConfig
			{
				SheetId = fixture.Create<int>()
			};

			// Add division configs
			foreach (var division in allTeams)
			{
				DivisionSheetConfig divConfig = DivisionSheetConfigFactory.GetForTeams(division.Value.Count());
				divConfig.DivisionName = division.Key;
				config.DivisionConfigs.Add(division.Key, divConfig);
			}

			Mock<ISheetsClient> mockClient = new Mock<ISheetsClient>();
			IList<Request>? capturedRequests = null;
			mockClient.Setup(x => x.ExecuteRequests(It.IsAny<IEnumerable<Request>>(), It.IsAny<CancellationToken>()))
				.Callback<IEnumerable<Request>, CancellationToken>((rqs, ct) => capturedRequests = rqs.ToList());

			ShootoutSheetService service = new ShootoutSheetService(mockClient.Object, Mock.Of<ILogger<ShootoutSheetService>>());

			// Act
			await service.HideHelperColumns(config);

			// Assert
			Assert.NotNull(capturedRequests);
			Assert.Equal(2, capturedRequests!.Count); // One to create group, one to hide it

			// Verify first request creates a dimension group
			Request groupRequest = capturedRequests[0];
			Assert.NotNull(groupRequest.AddDimensionGroup);
			Assert.Equal("COLUMNS", groupRequest.AddDimensionGroup.Range.Dimension);
			Assert.Equal(config.SheetId, groupRequest.AddDimensionGroup.Range.SheetId);

			// Verify second request hides the columns
			Request hideRequest = capturedRequests[1];
			Assert.NotNull(hideRequest.UpdateDimensionProperties);
			Assert.Equal("COLUMNS", hideRequest.UpdateDimensionProperties.Range.Dimension);
			Assert.Equal(config.SheetId, hideRequest.UpdateDimensionProperties.Range.SheetId);
			Assert.True(hideRequest.UpdateDimensionProperties.Properties.HiddenByUser);
			Assert.Equal("hiddenByUser", hideRequest.UpdateDimensionProperties.Fields);

			// Verify both requests target the same column range
			DivisionSheetConfig anyDivisionConfig = config.DivisionConfigs.Values.First();
			ShootoutSheetHelper helper = new ShootoutSheetHelper(anyDivisionConfig);
			int expectedStartColumn = helper.LastVisibleColumnIndex + 1;
			int expectedEndColumn = helper.SortedStandingsListColumnIndex + anyDivisionConfig.NumberOfTeams;

			Assert.Equal(expectedStartColumn, groupRequest.AddDimensionGroup.Range.StartIndex);
			Assert.Equal(expectedEndColumn, groupRequest.AddDimensionGroup.Range.EndIndex);
			Assert.Equal(expectedStartColumn, hideRequest.UpdateDimensionProperties.Range.StartIndex);
			Assert.Equal(expectedEndColumn, hideRequest.UpdateDimensionProperties.Range.EndIndex);
		}

		#region Helper Methods

		private (Mock<ISheetsClient>, Dictionary<string, IEnumerable<Team>>) SetupMockClientAndTeams()
		{
			Fixture fixture = new Fixture();
			Dictionary<string, IEnumerable<Team>> allTeams = new Dictionary<string, IEnumerable<Team>>();
			foreach (string division in ShootoutConstants.DivisionNames)
			{
				allTeams.Add(division, fixture.Build<Team>().With(x => x.DivisionName, division).CreateMany(TEST_TEAMS_PER_DIVISION));
			}

			Sheet sheet = new Sheet
			{
				Properties = fixture.Build<SheetProperties>()
					.OmitAutoProperties()
					.With(x => x.SheetId, () => fixture.Create<int>())
					.With(x => x.GridProperties, new GridProperties())
					.Create()
			};

			Mock<ISheetsClient> mockClient = new Mock<ISheetsClient>();
			mockClient.Setup(x => x.CreateSpreadsheet(It.IsAny<string>(), It.IsAny<CancellationToken>()));
			mockClient.Setup(x => x.GetOrAddSheet(It.IsAny<string>(), It.IsAny<int?>(), It.IsAny<int?>(), It.IsAny<CancellationToken>()))
				.ReturnsAsync(sheet);
			mockClient.Setup(x => x.Append(It.IsAny<IList<AppendRequest>>(), It.IsAny<CancellationToken>()));
			mockClient.Setup(x => x.ExecuteRequests(It.IsAny<IEnumerable<Request>>(), It.IsAny<CancellationToken>()));
			mockClient.Setup(x => x.AutoResizeColumn(It.IsAny<string>(), It.IsAny<int>()))
				.ReturnsAsync(100);

			return (mockClient, allTeams);
		}

		#endregion
	}
}