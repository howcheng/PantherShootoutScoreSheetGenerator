using Google.Apis.Sheets.v4.Data;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public static class Colors
	{
		// #6d9eeb -- blue
		public static System.Drawing.Color HeaderRowColor = System.Drawing.Color.FromArgb(0x6D, 0x9E, 0xEB);
		// #b6d7a7 -- light green
		public static System.Drawing.Color SubheaderRowColor = System.Drawing.Color.FromArgb(0xB6, 0xD7, 0xA7);

		public static Color FirstPlaceColor = new Color
		{
			// Gold: ffd966
			Red = 1,
			Green = (float)0xD9 / 0xFF,
			Blue = (float)0x66 / 0xFF,
			Alpha = 1,
		};
		public static Color SecondPlaceColor = new Color
		{
			// Silver: c0c0c0
			Red = (float)0xC0 / 0xFF,
			Green = (float)0xC0 / 0xFF,
			Blue = (float)0xC0 / 0xFF,
			Alpha = 1,
		};
		public static Color ThirdPlaceColor = new Color
		{
			// Bronze: 7f6000
			Red = (float)0x7F / 0xFF,
			Green = (float)0x60 / 0xFF,
			Blue = 0,
			Alpha = 1,
		};
		public static Color FourthPlaceColor = new Color
		{
			// light purple: b4a7d6
			Red = (float)0xB4 / 0xFF,
			Green = (float)0xA7 / 0xFF,
			Blue = (float)0xD6 / 0xFF,
			Alpha = 1,
		};

		public static Color GetColorForRank(int rank)
		{
			Color color;
			switch (rank)
			{
				case 1:
					color = FirstPlaceColor;
					break;
				case 2:
					color = SecondPlaceColor;
					break;
				case 3:
					color = ThirdPlaceColor;
					break;
				default:
					color = FourthPlaceColor;
					break;
			}
			return color;
		}
	}
}
