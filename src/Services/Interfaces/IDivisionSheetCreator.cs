﻿using Google.Apis.Sheets.v4.Data;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IDivisionSheetCreator
	{
		Task CreateSheet();
	}
}
