using GoogleSheetsHelper;
using Microsoft.Extensions.DependencyInjection;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public abstract class SheetGeneratorTests
	{
		protected const string POOL_A = "A";
		protected const string POOL_B = "B";
		protected const string POOL_C = "C";
		private readonly string[] _pools = new[] { POOL_A, POOL_B, POOL_C };

		// https://stackoverflow.com/questions/1068095/assigning-out-ref-parameters-in-moq
		private delegate PoolPlayInfo ScoreInputReturns(PoolPlayInfo ppi, IEnumerable<Team> ts, int rnd, ref int idx);

		protected readonly Mock<IScoreSheetHeadersRequestCreator> _mockHeadersCreator = new();
		protected readonly Mock<IScoreInputsRequestCreator> _mockInputsCreator = new();
		protected readonly Mock<IStandingsTableRequestCreator> _mockStandingsCreator = new();
		protected readonly Mock<ITiebreakerColumnsRequestCreator> _mockTiebreakerColsCreator = new();
		protected readonly Mock<ISortedStandingsListRequestCreator> _mockSortedStandingsCreator = new();

		protected List<Team> CreateTeams(DivisionSheetConfig config)
		{
			List<Team> ret = new List<Team>();
			Fixture fixture = new Fixture();
			for (int i = 0; i < config.NumberOfPools; i++)
			{
 				int counter = 0;
				string poolName = _pools[i];
				IEnumerable<Team> teams = fixture.Build<Team>()
					.With(x => x.DivisionName, ShootoutConstants.DIV_10UB)
					.With(x => x.PoolName, poolName)
					.With(x => x.TeamSheetCell, () => $"{poolName}{++counter}")
					.CreateMany(config.TeamsPerPool);
				ret.AddRange(teams);
			}
			return ret;
		}

		protected static IStandingsRequestCreatorFactory CreateStandingsRequestCreatorFactory(IEnumerable<Team> teams)
		{
			IServiceCollection services = new ServiceCollection();
			services.AddPantherShootoutServices(teams, new ShootoutSheetConfig());
			IServiceProvider provider = services.BuildServiceProvider();
			return provider.GetRequiredService<IStandingsRequestCreatorFactory>();
		}

		protected Action<UpdateRequest, PoolPlayInfo, int> AssertChampionshipLabelRequest = (rq, info, rowIdx) =>
		{
			// label row
			Assert.Equal(rowIdx, rq.RowStart);
			IEnumerable<string> labelRowValues = rq.Rows.Single().Select(x => x.StringValue);
			Assert.Single(labelRowValues.Where(x => !string.IsNullOrEmpty(x)));
			Assert.All(rq.Rows.Single(), cell => Assert.True(cell.Bold));
			Assert.All(rq.Rows.Single(), cell => Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.HeaderRowColor)));
			Assert.Equal(info.ChampionshipStartRowIndex, rq.RowStart);
		};

		/// <summary>
		/// Arguments: <see cref="UpdateRequest"/>, text of the label, row index
		/// </summary>
		protected Action<UpdateRequest, string, int> AssertChampionshipSubheader = (rq, label, rowIdx) =>
		{
			Assert.Equal(rowIdx, rq.RowStart);
			GoogleSheetRow row = rq.Rows.Single();
			GoogleSheetCell? labelCell = row.SingleOrDefault(x => !string.IsNullOrEmpty(x.StringValue));
			Assert.NotNull(labelCell);
			Assert.StartsWith(label, labelCell!.StringValue);
			Assert.True(labelCell.Bold);
			Assert.All(row, cell => Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.SubheaderRowColor)));
		};

		protected void CreateMocksForPoolPlayRequestCreatorTests(DivisionSheetConfig config)
		{
			_mockHeadersCreator.Setup(x => x.CreateHeaderRequests(It.IsAny<PoolPlayInfo>(), It.IsAny<string>(), It.IsAny<int>(), It.IsAny<IEnumerable<Team>>()))
				.Returns((PoolPlayInfo ppi, string poolName, int rowIdx, IEnumerable<Team> teams) => ppi);

			_mockInputsCreator.Setup(x => x.CreateScoringRequests(It.IsAny<PoolPlayInfo>(), It.IsAny<IEnumerable<Team>>(), It.IsAny<int>(), ref It.Ref<int>.IsAny))
				.Returns(new ScoreInputReturns((PoolPlayInfo ppi, IEnumerable<Team> ts, int rnd, ref int idx) =>
				{
					idx += config.GamesPerRound + 2; // +2 accounts for the blank row at the end
					return ppi;
				}));

			_mockStandingsCreator.Setup(x => x.CreateStandingsRequests(It.IsAny<PoolPlayInfo>(), It.IsAny<IEnumerable<Team>>(), It.IsAny<int>()))
				.Returns((PoolPlayInfo ppi, IEnumerable<Team> ts, int idx) => ppi);

			_mockTiebreakerColsCreator.Setup(x => x.CreateTiebreakerRequests(It.IsAny<PoolPlayInfo>(), It.IsAny<IEnumerable<Team>>(), It.IsAny<int>()))
				.Returns((PoolPlayInfo ppi, IEnumerable<Team> ts, int idx) => ppi);

			_mockSortedStandingsCreator.Setup(x => x.CreateSortedStandingsListRequest(It.IsAny<PoolPlayInfo>()))
				.Returns((PoolPlayInfo ppi) => ppi);
		}
	}
}
