using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public abstract class SheetGeneratorTests
	{
		protected const string POOL_A = "A";
		protected const string POOL_B = "B";
		protected const string POOL_C = "C";
		private readonly string[] _pools = new[] { POOL_A, POOL_B, POOL_C };


		protected List<Team> CreateTeams(DivisionSheetConfig config)
		{
			List<Team> ret = new List<Team>();
			for (int i = 0; i < config.NumberOfPools; i++)
			{
 				int counter = 0;
				string poolName = _pools[counter];
				Fixture fixture = new Fixture();
				IEnumerable<Team> teams = fixture.Build<Team>()
					.With(x => x.DivisionName, ShootoutConstants.DIV_10UB)
					.With(x => x.PoolName, poolName)
					.With(x => x.TeamSheetCell, () => $"{poolName}{++counter}")
					.CreateMany(config.TeamsPerPool);
				ret.AddRange(teams);
			}
			return ret;
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
	}
}
