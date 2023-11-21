# PantherShootoutScoreSheetGenerator
This console application generates a Google Sheet for entering scores for Region 42's Panther Shooutout tournament

## Steps to follow every year
1. Create a CSV file of the teams (see an example in the Services.Tests project)
1. Update the path to the file in the `appSettings.json` file
1. Recompile the app and run it

## 2024 notes
PoolPlayRequestCreator10Teams and PoolPlayRequestCreator12Teams need to be updated with new tiebreaker columns and the sorted team list (in 2023, all pools had 8 teams).