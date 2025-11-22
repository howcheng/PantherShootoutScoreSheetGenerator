# PantherShootoutScoreSheetGenerator
This console application generates a Google Sheet for entering scores for Region 42's Panther Shooutout tournament. The entire spreadsheet is created programmatically using the Google Sheets API. The document is then published to the web and embedded in the Region 42 website for tournament participants to view. Tournament staff use the file directly to enter scores during the event.

## Steps to follow every year
1. Create a CSV file of the teams (see an example in the Services.Tests project)
1. Update the path to the file in the `appSettings.json` file
1. Recompile the app and run it

## Architecture and Process Flow

### Overview
This application generates a comprehensive Google Sheets workbook for managing a multi-division youth soccer tournament. It creates two types of sheets:
1. **Shootout Sheet**: A master sheet listing all teams across all divisions for entering the scores from the shootout rounds
2. **Division Sheets**: Individual sheets for each division containing pool play scoring, standings tables, and championship brackets

The application uses a modular, dependency injection-based architecture with specialized "request creator" services that generate Google Sheets API requests to build different parts of the spreadsheet.

### Entry Point and Main Process Flow

The application starts in `Program.Main()` and follows this sequence:

#### 1. **Configuration and OAuth** (`Program.Run()`)
```
Load appSettings.json → Build shared services → Perform OAuth authentication → Obtain Google credentials
```
- Loads `AppSettings` from configuration file
- Builds a shared service provider with OAuth checker
- Authenticates with Google using OAuth 2.0 (via `GoogleOAuthCliClient` library)
- Creates a `GoogleCredential` from the access token

#### 2. **Data Loading and Spreadsheet Creation**
```
Read team CSV → Group teams by division → Create spreadsheet → Generate Shootout sheet
```
- `TeamReaderService` reads teams from CSV file specified in `appSettings.json`
- Teams are parsed into `Team` objects with properties: `DivisionName`, `TeamName`, `PoolName`
- Teams are grouped by division and ordered according to `ShootoutConstants.DivisionNames`
- `SheetsClient` creates a new Google Spreadsheet
- `ShootoutSheetService` generates the master "Shootout" sheet with team listings

#### 3. **Shootout Sheet Generation** (`ShootoutSheetService.GenerateSheet()`)
The Shootout sheet is built with:
- **Team rows**: One row per team showing team name and score columns for each round
- **Division headers**: Sections grouped by division (e.g., "U10 Boys", "U12 Girls")
- **Score columns**: Columns for Round 1-4 scores (3 rounds for most divisions, 4 for divisions with 5 teams per pool)
- **Total column**: Formula summing scores across all rounds

Returns a `ShootoutSheetConfig` object containing:
- `SheetId`: Google Sheets sheet ID
- `TeamNameCellWidth`: Width for team name columns
- `DivisionConfigs`: Dictionary mapping division names to their configurations
- `ShootoutStartAndEndRows`: Row numbers where each division's score entry will be placed
- `FirstTeamSheetCells`: Cell references for the first team in each division

#### 4. **Division Sheet Generation Loop**
For each division, the application:
```
Build division services → Generate division sheet → Add score entry to Shootout sheet
```

##### 4a. **Build Division-Specific Services** (`BuildServiceProviderForDivisionSheet()`)
Creates a new service provider with:
- `DivisionSheetConfig`: Configuration based on team count (4, 5, 6, 8, 10, or 12 teams)
- Team-count-specific request creators (explained below)
- `DivisionSheetGenerator`: Main orchestrator for building the division sheet

##### 4b. **Division Sheet Creation** (`DivisionSheetGenerator.CreateSheet()`)
The division sheet generator orchestrates the creation of:

**Pool Play Section:**
- Created by `IPoolPlayRequestCreator` (implementations: `PoolPlayRequestCreator`, `PoolPlayRequestCreator10Teams`, `PoolPlayRequestCreator12Teams`)
- Generates:
  - **Score entry fields**: Input cells for match scores (home team, away team, goals for each)
  - **Standings table**: Live-updating table with columns for GP (games played), W, L, D, Pts, GF, GA, GD
  - **Tiebreaker columns**: Hidden helper columns for breaking ties (head-to-head, goals against, etc.)
  - **Sorted standings list**: Live-sorted team rankings based on points and tiebreakers

**Championship Section:**
- Created by `IChampionshipRequestCreator` (implementations vary by team count: `ChampionshipRequestCreator4Teams`, `ChampionshipRequestCreator6Teams`, etc.)
- Generates bracket structure for:
  - Semifinals (if applicable)
  - Championship game
  - Third place game
  - Consolation games

**Winner Formatting:**
- Created by `IWinnerFormattingRequestsCreator`
- Applies conditional formatting to highlight winning teams

Returns a `PoolPlayInfo` (or `ChampionshipInfo`) object containing:
- Row numbers for standings tables
- Row numbers for score entry sections
- Row numbers for championship games
- All accumulated Google Sheets API requests

##### 4c. **Shootout Score Entry** (`ShootoutScoreEntryService.CreateShootoutScoreEntrySection()`)
Creates the score entry section on the Shootout sheet for this division:
- Uses `IShootoutScoreEntryRequestsCreator` (standard or 6-teams variant)
- Generates formulas that pull scores from the division sheet
- Adds tiebreaker columns
- Creates sorted standings list on the Shootout sheet

### Key Domain Objects

#### Configuration Objects
- **`AppSettings`**: Application configuration (file paths, etc.)
- **`ShootoutSheetConfig`**: Configuration for the master Shootout sheet
- **`DivisionSheetConfig`**: Configuration for a division (team count, pools, rounds, etc.)
  - Subclasses: `DivisionSheetConfig4Teams`, `DivisionSheetConfig6Teams`, etc.

#### Data Objects
- **`Team`**: Represents a team (division name, team name, pool name, cell reference)
- **`PoolPlayInfo`**: Contains pools, row numbers, and accumulated requests for pool play
- **`ChampionshipInfo`**: Extends `PoolPlayInfo` with championship game row numbers
- **`SheetRequests`**: Container for Google Sheets API requests (`UpdateRequest` and `Request`)

### Service Architecture

#### Core Services
- **`ITeamReaderService`** (`TeamReaderService`): Reads teams from CSV using `LINQtoCSV` library
- **`IShootoutSheetService`** (`ShootoutSheetService`): Generates the Shootout master sheet
- **`IDivisionSheetGenerator`** (`DivisionSheetGenerator`): Orchestrates division sheet creation
- **`IShootoutScoreEntryService`** (`ShootoutScoreEntryService`): Creates score entry sections on Shootout sheet

#### Request Creator Pattern
The application uses a "request creator" pattern where specialized services generate Google Sheets API requests. These are registered based on the number of teams in a division:

**Pool Play Request Creators:**
- `IPoolPlayRequestCreator`: Main pool play section
- `IStandingsTableRequestCreator`: Standings table with GP, W, L, D, Pts, GF, GA, GD
- `ITiebreakerColumnsRequestCreator`: Tiebreaker calculation columns
- `ISortedStandingsListRequestCreator`: Live-sorted standings

**Championship Request Creators:**
- `IChampionshipRequestCreator`: Championship bracket structure
- Team-count-specific implementations handle different bracket formats

**Score Entry Request Creators:**
- `IShootoutScoreEntryRequestsCreator`: Score entry on Shootout sheet
- Specialized implementation for 6-team divisions: `ShootoutScoreEntryRequestsCreator6Teams`

**Formula Generators:**
- `FormulaGenerator`: Base class for generating Google Sheets formulas
- `PsoFormulaGenerator`: Pool play formulas
- `ShootoutScoringFormulaGenerator`: Shootout sheet formulas

#### Service Registration
The `PsoServicesServiceCollectionExtensions` class contains two key methods:

1. **`AddPantherShootoutServices()`**: Registers division-specific services based on team count
   - Determines which request creators to use (4, 5, 6, 8, 10, or 12 teams)
   - Registers all standings calculators (wins, losses, draws, goal differential, etc.)
   - Registers tiebreaker calculators (head-to-head, goals against, etc.)

2. **`AddShootoutScoreEntryServices()`**: Registers services for Shootout sheet score entry
   - Registers the appropriate `IShootoutScoreEntryRequestsCreator` based on team count

### External Dependencies

#### Google Sheets Integration
- **`GoogleSheetsHelper`**: Custom library for Google Sheets API operations
  - `ISheetsClient`: Interface for sheet operations (create, update, execute requests)
  - `GoogleSheetRow`, `GoogleSheetCell`: Strongly-typed sheet data models

- **`StandingsGoogleSheetsHelper`**: Library for standings table generation
  - `StandingsSheetHelper`: Helper for creating standings-related requests
  - `IStandingsRequestCreator`: Interface for standings-related request creators

#### Authentication
- **`GoogleOAuthCliClient`**: Custom library for OAuth 2.0 authentication
  - `IOAuthChecker`: Manages OAuth flow and token storage

#### Other Libraries
- **`LINQtoCSV`**: CSV parsing for team data
- **`Microsoft.Extensions.DependencyInjection`**: Dependency injection container
- **`Microsoft.Extensions.Configuration`**: Configuration management
- **`Google.Apis.Sheets.v4`**: Google Sheets API client library

### Data Flow Summary

```
CSV File
   ↓
TeamReaderService → Dictionary<Division, Teams>
   ↓
ShootoutSheetService → Creates Shootout Sheet → ShootoutSheetConfig
   ↓
For each division:
   ↓
   DivisionSheetGenerator
      ↓
      ├─→ PoolPlayRequestCreator → Pool Play Section
      ├─→ ChampionshipRequestCreator → Championship Bracket
      └─→ WinnerFormattingRequestsCreator → Conditional Formatting
      ↓
      Returns PoolPlayInfo/ChampionshipInfo
      ↓
   ShootoutScoreEntryService → Adds score entry to Shootout Sheet
   ↓
Complete Spreadsheet
```

### Key Design Patterns

1. **Dependency Injection**: All services use constructor injection for testability
2. **Factory Pattern**: `DivisionSheetConfigFactory` creates team-count-specific configurations
3. **Strategy Pattern**: Different request creators handle different team counts/scenarios
4. **Builder Pattern**: Request creators accumulate API requests before batch execution
5. **Service Provider per Scope**: Different service providers for shared, document, division, and score entry scopes

### Testing Considerations

The architecture is designed for testability:
- All major components implement interfaces
- Request creators can be tested independently
- Formula generation is isolated in dedicated classes
- File I/O is abstracted behind `IFileReader` interface
- Google Sheets API calls are abstracted behind `ISheetsClient` interface

This allows for comprehensive unit testing without requiring actual Google Sheets API calls or file system access.

