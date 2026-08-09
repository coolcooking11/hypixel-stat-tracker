import { google } from "googleapis";
process.on("unhandledRejection", (error) => {
    console.error("Error:\n", error.message || error)
    process.exit(1)
})
// processes newline characters in google private key
const googleKey = process.env.GOOGLE_KEY.replace(/\\n/g, '\n')
const auth = new google.auth.GoogleAuth({
    credentials: {
        client_email: process.env.GOOGLE_EMAIL,
        private_key: googleKey
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
})
const sheets = google.sheets({ version: 'v4', auth })
let metadata = await sheets.spreadsheets.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
})
// list of spreadsheet tab names
let tabs = metadata.data.sheets.map((sheet) => sheet.properties.title)
if (tabs.includes("Config")) {
    // if the spreadsheet has been initialized
    const response = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.SPREADSHEET_ID,
        range: "Config!B2:B6"
    })
    if (response.data.values[0][0] === "Yes") {
        await createTab("Bedwars Solos")
    }
    if (response.data.values[1][0] === "Yes") {
        await createTab("Bedwars Duos")
    }
    if (response.data.values[2][0] === "Yes") {
        await createTab("Bedwars Threes")
    }
    if (response.data.values[3][0] === "Yes") {
        await createTab("Bedwars Fours")
    }
    if (response.data.values[4][0] === "Yes") {
        await createTab("Bedwars 4v4")
    }
} else {
    // initializes the spreadsheet if it has not been initalized
    initalize()
}
// creates config file
async function initalize() {
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId: process.env.SPREADSHEET_ID,
        requestBody: {
            requests: [
                {
                    updateSheetProperties: {
                        properties: {
                            sheetId: 0,
                            title: "Config",
                        },
                        fields: 'title',
                    },
                },
                {
                    setDataValidation: {
                        range: {
                            sheetId: 0,
                            startRowIndex: 1,
                            endRowIndex: 6,
                            startColumnIndex: 1,
                            endColumnIndex: 2
                        },
                        rule: {
                            condition: {
                                type: "ONE_OF_LIST",
                                values: [
                                    { userEnteredValue: "Yes" },
                                    { userEnteredValue: "No" },
                                ]
                            },
                            showCustomUi: true,
                            strict: true
                        }
                    }
                },
                {
                    updateDimensionProperties: {
                        range: {
                            sheetId: 0,
                            dimension: "COLUMNS",
                            startIndex: 0,
                            endIndex: 1
                        },
                        properties: {
                            pixelSize: 104
                        },
                        fields: "pixelSize"
                    }
                }
            ]
        }
    })
    await sheets.spreadsheets.values.append({
        spreadsheetId: process.env.SPREADSHEET_ID,
        range: "Config!A:B",
        valueInputOption: "USER_ENTERED",
        requestBody: {
            values: [["Mode", "Enabled"], ["Bedwars Solos", "No"], ["Bedwars Duos", "No"], ["Bedwars Threes", "No"], ["Bedwars Fours", "No"], ["Bedwars 4v4", "No"]]
        }
    })
}
// creates tab with headers
async function createTab(name) {
    // if tab already exists, update tab with data
    if (tabs.includes(name)) {
        updateTab(name)
        return
    }
    // order of tabs if all modes are selected
    const tabOrder = ["Bedwars Solos", "Bedwars Duos", "Bedwars Threes", "Bedwars Fours", "Bedwars 4v4", "Config"]
    const relativeIndex = tabOrder.indexOf(name) // where the new tab should be relative to the others
    const followingTabs = tabOrder.slice(relativeIndex + 1) // tabs that follow the new tab in the master order
    const nextTab = followingTabs.find((title) => tabs.includes(title)) // finds next existing tab
    // creates tab
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId: process.env.SPREADSHEET_ID,
        requestBody: {
            requests: [
                {
                    addSheet: {
                        properties: {
                            title: name,
                            index: tabs.indexOf(nextTab)
                        }
                    }
                }
            ]
        }
    })
    // adds headers
    await sheets.spreadsheets.values.append({
        spreadsheetId: process.env.SPREADSHEET_ID,
        range: `${name}!A:AA`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
            values: [['Date', 'Lifetime Kills', 'Lifetime Deaths', 'Lifetime KDR', 'Lifetime Final Kills', 'Lifetime Final Deaths', 'Lifetime FKDR', 'Lifetime Beds Broken', 'Lifetime Beds Lost', 'Lifetime BBLR', 'Lifetime Wins', 'Lifetime Losses', 'Lifetime WLR', 'Lifetime Games Played', 'Session Kills', 'Session Deaths', 'Session KDR', 'Session Final Kills', 'Session Final Deaths', 'Session FKDR', 'Session Beds Broken', 'Session Beds Lost', 'Session BBLR', 'Session Wins', 'Session Losses', 'Session WLR', 'Session Games Played']]
        }
    })
    // updates current tab order
    metadata = await sheets.spreadsheets.get({
        spreadsheetId: process.env.SPREADSHEET_ID,
    })
    tabs = metadata.data.sheets.map((sheet) => sheet.properties.title)
    // resizes columns to fit data
    const tab = metadata.data.sheets.find((sheet) => sheet.properties.title === name)
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId: process.env.SPREADSHEET_ID,
        requestBody: {
            requests: [
                {
                    updateDimensionProperties: {
                        range: {
                            sheetId: tab.properties.sheetId,
                            dimension: "COLUMNS",
                            startIndex: 0,
                            endIndex: 27
                        },
                        properties: {
                            pixelSize: 150
                        },
                        fields: "pixelSize"
                    }
                }
            ]
        }
    })
}
// writes data to a tab
async function updateTab(name) {
    const res = await fetch(`https://api.hypixel.net/v2/player?uuid=${process.env.UUID}`, {headers: {"API-Key": process.env.HYPIXEL_KEY}})
    const data = await res.json()
    // defines relevant stats
    let kills
    let deaths
    let finalKills
    let finalDeaths
    let bedsBroken
    let bedsLost
    let wins
    let losses
    let gamesPlayed
    if (name === "Bedwars Solos") {
        kills = data.player.stats.Bedwars.eight_one_kills_bedwars || 0
        deaths = data.player.stats.Bedwars.eight_one_deaths_bedwars || 0
        finalKills = data.player.stats.Bedwars.eight_one_final_kills_bedwars || 0
        finalDeaths = data.player.stats.Bedwars.eight_one_final_deaths_bedwars || 0
        bedsBroken = data.player.stats.Bedwars.eight_one_beds_broken_bedwars || 0
        bedsLost = data.player.stats.Bedwars.eight_one_beds_lost_bedwars || 0
        wins = data.player.stats.Bedwars.eight_one_wins_bedwars || 0
        losses = data.player.stats.Bedwars.eight_one_losses_bedwars || 0
        gamesPlayed = data.player.stats.Bedwars.eight_one_games_played_bedwars || 0
    } else if (name === "Bedwars Duos") {
        kills = data.player.stats.Bedwars.eight_two_kills_bedwars || 0
        deaths = data.player.stats.Bedwars.eight_two_deaths_bedwars || 0
        finalKills = data.player.stats.Bedwars.eight_two_final_kills_bedwars || 0
        finalDeaths = data.player.stats.Bedwars.eight_two_final_deaths_bedwars || 0
        bedsBroken = data.player.stats.Bedwars.eight_two_beds_broken_bedwars || 0
        bedsLost = data.player.stats.Bedwars.eight_two_beds_lost_bedwars || 0
        wins = data.player.stats.Bedwars.eight_two_wins_bedwars || 0
        losses = data.player.stats.Bedwars.eight_two_losses_bedwars || 0
        gamesPlayed = data.player.stats.Bedwars.eight_two_games_played_bedwars || 0
    } else if (name === "Bedwars Threes") {
        kills = data.player.stats.Bedwars.four_three_kills_bedwars || 0
        deaths = data.player.stats.Bedwars.four_three_deaths_bedwars || 0
        finalKills = data.player.stats.Bedwars.four_three_final_kills_bedwars || 0
        finalDeaths = data.player.stats.Bedwars.four_three_final_deaths_bedwars || 0
        bedsBroken = data.player.stats.Bedwars.four_three_beds_broken_bedwars || 0
        bedsLost = data.player.stats.Bedwars.four_three_beds_lost_bedwars || 0
        wins = data.player.stats.Bedwars.four_three_wins_bedwars || 0
        losses = data.player.stats.Bedwars.four_three_losses_bedwars || 0
        gamesPlayed = data.player.stats.Bedwars.four_three_games_played_bedwars || 0
    } else if (name === "Bedwars Fours") {
        kills = data.player.stats.Bedwars.four_four_kills_bedwars || 0
        deaths = data.player.stats.Bedwars.four_four_deaths_bedwars || 0
        finalKills = data.player.stats.Bedwars.four_four_final_kills_bedwars || 0
        finalDeaths = data.player.stats.Bedwars.four_four_final_deaths_bedwars || 0
        bedsBroken = data.player.stats.Bedwars.four_four_beds_broken_bedwars || 0
        bedsLost = data.player.stats.Bedwars.four_four_beds_lost_bedwars || 0
        wins = data.player.stats.Bedwars.four_four_wins_bedwars || 0
        losses = data.player.stats.Bedwars.four_four_losses_bedwars || 0
        gamesPlayed = data.player.stats.Bedwars.four_four_games_played_bedwars || 0
    } else if (name === "Bedwars 4v4") {
        kills = data.player.stats.Bedwars.two_four_kills_bedwars || 0
        deaths = data.player.stats.Bedwars.two_four_deaths_bedwars || 0
        finalKills = data.player.stats.Bedwars.two_four_final_kills_bedwars || 0
        finalDeaths = data.player.stats.Bedwars.two_four_final_deaths_bedwars || 0
        bedsBroken = data.player.stats.Bedwars.two_four_beds_broken_bedwars || 0
        bedsLost = data.player.stats.Bedwars.two_four_beds_lost_bedwars || 0
        wins = data.player.stats.Bedwars.two_four_wins_bedwars || 0
        losses = data.player.stats.Bedwars.two_four_losses_bedwars || 0
        gamesPlayed = data.player.stats.Bedwars.two_four_games_played_bedwars || 0
    }
    const timeZone = process.env.TIMEZONE || "America/New_York"
    const date = new Date()
    date.setDate(date.getDate() - 1)
    const yesterday = new Intl.DateTimeFormat("en-CA", {
        timeZone,
        year: "numeric",
        month: "2-digit",
        day: "2-digit"
    }).format(date)
    // gets all data in the sheet
    const response = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.SPREADSHEET_ID,
        range: `${name}!A:AA`
    })
    // gets the last row of the sheet (previous day's data)
    let previousRow = response.data.values.at(-1)
    // variable containing today's data
    let newData = [[yesterday, kills, deaths, calculateRatio(kills, deaths), finalKills, finalDeaths, calculateRatio(finalKills, finalDeaths), bedsBroken, bedsLost, calculateRatio(bedsBroken, bedsLost), wins, losses, calculateRatio(wins, losses), gamesPlayed, "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"]]
    // if there is data present
    if (previousRow[0] !== "Date") {
        // parses all data into integers
        previousRow = previousRow.map((value) => parseInt(value))
        let sessionGamesPlayed = gamesPlayed - previousRow[13];
        // if no games were played today, don't log anything
        if (sessionGamesPlayed === 0) {
            return
        }
        // calculates session stats
        let sessionKills = kills - previousRow[1]
        let sessionDeaths = deaths - previousRow[2]
        let sessionFinalKills = finalKills - previousRow[4]
        let sessionFinalDeaths = finalDeaths - previousRow[5]
        let sessionBedsBroken = bedsBroken - previousRow[7]
        let sessionBedsLost = bedsLost - previousRow[8]
        let sessionWins = wins - previousRow[10]
        let sessionLosses = losses - previousRow[11]
        // replaces hyphens with real data
        newData[0][14] = sessionKills
        newData[0][15] = sessionDeaths
        newData[0][16] = calculateRatio(sessionKills, sessionDeaths)
        newData[0][17] = sessionFinalKills
        newData[0][18] = sessionFinalDeaths
        newData[0][19] = calculateRatio(sessionFinalKills, sessionFinalDeaths)
        newData[0][20] = sessionBedsBroken
        newData[0][21] = sessionBedsLost
        newData[0][22] = calculateRatio(sessionBedsBroken, sessionBedsLost)
        newData[0][23] = sessionWins
        newData[0][24] = sessionLosses
        newData[0][25] = calculateRatio(sessionWins, sessionLosses)
        newData[0][26] = sessionGamesPlayed
    }
    await sheets.spreadsheets.values.append({
        spreadsheetId: process.env.SPREADSHEET_ID,
        range: `${name}!A:AA`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
            values: newData
        }
    })
}
function calculateRatio(numerator, denominator) {
    if (denominator === 0) {
        return "-"
    }
    return (numerator / denominator).toFixed(2)
}