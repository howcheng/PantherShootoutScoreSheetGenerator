namespace PantherShootoutScoreSheetGenerator.Services
{
    public interface IStandingsTableRequestCreator
    {
        PoolPlayInfo CreateStandingsRequests(DivisionSheetConfig config, PoolPlayInfo info, IEnumerable<Team> poolTeams, int startRowIndex);
    }
}