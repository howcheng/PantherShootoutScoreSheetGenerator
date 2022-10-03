using Microsoft.Extensions.Logging;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class TeamReaderServiceTests
	{
		[Fact]
		public void CanReadTeams()
		{
			// our sample file has 2 divisions or 8 teams each: 10U Girls and 10U Boys

			using (FileStream stream = File.OpenRead("files\\PSO teams.csv"))
			{
				Mock<IFileReader> mockReader = new Mock<IFileReader>();
				mockReader.Setup(x => x.DoesFileExist(It.IsAny<string>())).Returns(true);
				mockReader.Setup(x => x.ReadAsStream(It.IsAny<string>())).Returns(stream);

				TeamReaderService service = new TeamReaderService(new AppSettings(), mockReader.Object, Mock.Of<ILogger<TeamReaderService>>());
				IDictionary<string, IEnumerable<Team>> teams = service.GetTeams();

				Assert.Equal(2, teams.Count);
				Assert.All(teams, pair =>
				{
					Assert.Equal(8, pair.Value.Count());
					Assert.All(pair.Value, t => Assert.Equal(pair.Key, t.DivisionName));
				});
			}
		}
	}
}