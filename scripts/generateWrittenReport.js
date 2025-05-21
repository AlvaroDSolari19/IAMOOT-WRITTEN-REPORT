const XLSX = require('xlsx'); 
const PathUtility = require('path'); 

const outputWorkbook = XLSX.utils.book_new(); 

generateRankingSheet('English', 'english.xlsx', 'English Rankings', outputWorkbook); 
generateRankingSheet('Spanish', 'spanish.xlsx', 'Spanish Rankings', outputWorkbook); 
generateRankingSheet('Portuguese', 'portuguese.xlsx', 'Portuguese Rankings', outputWorkbook);

const outputPath = PathUtility.join(__dirname, '../output/WrittenCompetitionResults.xlsx');
XLSX.writeFile(outputWorkbook, outputPath); 
console.log(`All language rankings written to ${outputPath}`);

function generateRankingSheet(languageLabel, inputFilename, outputSheetName, workbookToAppend){

    const inputFilePath = PathUtility.join(__dirname, '../data', inputFilename); 
    const inputWorkbook = XLSX.readFile(inputFilePath);

    const inputWorksheet = inputWorkbook.Sheets[inputWorkbook.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(inputWorksheet, { header: 1 });
    const dataRows = rawRows.slice(1); 

    const teamScores = {}; 

    dataRows.forEach((scoreRow) => {
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

        if (!teamScores[teamID]){ 
            teamScores[teamID] = {
                victimScores: [], 
                stateScores: [], 
                teamLanguage: languageLabel
            };
        }

        teamScores[teamID][scoreCategory].push(totalScore); 

    })

    const rankingTable = [
        ["Team ID", "Victim Average", "State Average", "Overall Average"]
    ]

    for (const teamID in teamScores){
        const teamData = teamScores[teamID]; 
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

        rankingTable.push([
            teamID, 
            victimAverage !== null ? victimAverage.toFixed(2) : '', 
            stateAverage !== null ? stateAverage.toFixed(2) : '', 
            overallAverage !== null ? overallAverage.toFixed(2) : ''
        ])
    }

    const sheet = XLSX.utils.aoa_to_sheet(rankingTable);
    XLSX.utils.book_append_sheet(workbookToAppend, sheet, outputSheetName); 

}