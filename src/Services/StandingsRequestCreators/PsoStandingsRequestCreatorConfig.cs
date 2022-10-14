using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class PsoStandingsRequestCreatorConfig : ScoreBasedStandingsRequestCreatorConfig
	{
		public int NumberOfRounds { get; set; }
		public int GamesPerRound { get; set; }
	}
}
