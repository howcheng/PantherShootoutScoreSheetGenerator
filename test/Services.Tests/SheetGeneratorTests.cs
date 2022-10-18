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

	}
}
