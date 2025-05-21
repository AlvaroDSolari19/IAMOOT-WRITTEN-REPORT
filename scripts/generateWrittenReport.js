const XLSX = require('xlsx'); 
const FileSystem = require('fs'); 
const PathUtility = require('path'); 

const englishFilePath = PathUtility.join(__dirname, '../data/english.xlsx'); 
const englishWorkbook = XLSX.readFile(englishFilePath);

const englishWorksheet = englishWorkbook.Sheets[englishWorkbook.SheetNames[0]];
const englishRawRows = XLSX.utils.sheet_to_json(englishWorksheet, { header: 1 });
const englishDataRows = englishRawRows.slice(1); 

const englishTeamScores = {}; 

englishDataRows.forEach((scoreRow) => {
    const rawMemoCode = scoreRow[3];
    if (!rawMemoCode || typeof rawMemoCode !== 'string') return; 

    const memoCodeRegex = new RegExp('^(\\d+)([A-Za-z])');
    const parsedMemoCode = rawMemoCode.trim().match(memoCodeRegex);
    if (!parsedMemoCode) return; 

    const teamID = parsedMemoCode[1]; 
    const roleIndicator = parsedMemoCode[2].toUpperCase(); 

    let scoreCategory = null; 
    if (roleIndicator === 'V'){
        scoreCategory = 'victimScores'; 
    } else if (roleIndicator === 'S' || roleIndicator === 'E'){
        scoreCategory = 'stateScores'; 
    } else {
        return; 
    }

    const scoreColumns = scoreRow.slice(4, 10); 
    const totalScore = scoreColumns.reduce((runningTotal, scoreCellValue) => {
        const numericScore = parseFloat(scoreCellValue);
        return runningTotal + (isNaN(numericScore) ? 0 : numericScore); 
    }, 0);

    if (!englishTeamScores[teamID]){ 
        englishTeamScores[teamID] = {
            victimScores: [], 
            stateScores: [], 
            teamLanguage: 'English'
        };
    }

    englishTeamScores[teamID][scoreCategory].push(totalScore); 

})

const englishRankingTable = [
    ["Team ID", "Victim Average", "State Average", "Overall Average"]
]

for (const teamID in englishTeamScores){
    const teamData = englishTeamScores[teamID]; 
    const victimScores = teamData.victimScores; 
    const stateScores = teamData.stateScores; 

    const victimAverage = victimScores.length > 0 ? (victimScores.reduce((accumulatedTotal, currentVictimScore) => {
        return accumulatedTotal + currentVictimScore;
    }, 0) / victimScores.length ) : null; 
    
    const stateAverage = stateScores.length > 0 ? (stateScores.reduce((accumulatedTotal, currentStateScore) => {
        return accumulatedTotal + currentStateScore; 
    }, 0) / stateScores.length ) : null; 

    let overallAverage = null; 
    if (victimAverage !== null && stateAverage !== null){
        overallAverage = (victimAverage + stateAverage) / 2; 
    } else { 
        overallAverage = victimAverage ?? stateAverage; 
    }

    englishRankingTable.push([
        teamID, 
        victimAverage !== null ? victimAverage.toFixed(2) : '', 
        stateAverage !== null ? stateAverage.toFixed(2) : '', 
        overallAverage !== null ? overallAverage.toFixed(2) : ''
    ])
}

const outputWorkbook = XLSX.utils.book_new(); 
const englishSheet = XLSX.utils.aoa_to_sheet(englishRankingTable); 

XLSX.utils.book_append_sheet(outputWorkbook, englishSheet, 'English Rankings'); 

const outputPath = PathUtility.join(__dirname, '../output/WrittenCompetitionResults.xlsx');
XLSX.writeFile(outputWorkbook, outputPath); 

console.log(`English Rankings sheet written to ${outputPath}`);