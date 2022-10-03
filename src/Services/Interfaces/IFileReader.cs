namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IFileReader
	{
		bool DoesFileExist(string path);
		Stream ReadAsStream(string path);
	}
}