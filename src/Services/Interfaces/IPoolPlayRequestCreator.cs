namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IPoolPlayRequestCreator
	{
		/// <summary>
		/// Creates all the Google requests for pool play matches (i.e., excluding championship games)
		/// </summary>
		/// <param name="info"></param>
		/// <returns></returns>
		PoolPlayInfo CreatePoolPlayRequests(PoolPlayInfo info);
	}
}