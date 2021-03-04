/**
 * Author: Mason Housenga (masonrh2@illinois.edu) 
 * 
 * This is the Google Apps Script project which is bound to the sPHENIX "Blocks Database" Google Sheet
 * 
 * This file defines a trigger function which, when active, adds formulas to blocks if they are being tested
 * and freezes the same values in place if they are not.
*/

// THIS IS THE REAL DATABASE

// CONSTANTS
/** name of sector 13-64 sheet */
const blocks1364 = 'Blocks1364DB'
/** sheets to check for status changes */
const checkSheets = [blocks1364]
/** the alphabetic column for status */
const statusColumn = 'D'
/** blocks (DBNs) start at THIS row (should be AFTER headers, e.g.) */
const dataStartRow = 2
/** add formulas for blocks in these statuses (and otherwise freeze values in place) */
const addFormulaStatuses = [5, '5a', '5b', '5c', '5m', '5r']
/** 
 * Array of "run" objects which represent connected runs of formulas to update
 * Runs are connected (contiguous) columns of formulas (or output for formulas) NOT including any data to be entered manually
 * Specify the start column using startCol (crucial; one-based)
 * FormulaValues defines what the formulas should be starting at startCol and continuing into the next columns
 * Put the formula (including '=') in formulaValues for columns which should have a formula using @row in place of the cell's own row number
 * Values which are null will DELETE the cell's value (needed for cells which display formula output but are not themselves formulas)
 * Some changes need to be made if other any other sheet other than blocks1364 needs formulas
*/
const formulaRuns = [
  {
    formulaValues: [
      /*AU*/ "=INDEX(FiberCountingDataDump!B$2:N,MAX(ROW(FiberCountingDataDump!$A$2:$A)*(Blocks1364DB!$A@row=FiberCountingDataDump!$A$2:$A))-1)",
      /*AV*/ null,
      /*AW*/ null,
      /*AX*/ null,
      /*AY*/ null,
      /*AZ*/ null,
      /*BA*/ null,
      /*BB*/ null,
      /*BC*/ null,
      /*BD*/ null,
      /*BE*/ null,
      /*BF*/ null,
      /*BG*/ null,
    ],
    startCol: letterToColumn('AU'),
    numFormulas: null
  },
  {
    formulaValues: [
      /*BM*/ "=INDEX(FiberCountingDataDump!O$2:O,MAX(ROW(FiberCountingDataDump!$A$2:$A)*(Blocks1364DB!$A@row=FiberCountingDataDump!$A$2:$A))-1)",
      /*BN*/ "=INDEX(ScintillationDataDump!B$2:F,MAX(ROW(ScintillationDataDump!$A$2:$A)*(Blocks1364DB!$A@row=ScintillationDataDump!$A$2:$A))-1)",
      /*BO*/ null,
      /*BP*/ null,
      /*BQ*/ null,
      /*BR*/ null
    ],
    startCol: letterToColumn('BM'),
    numFormulas: null
  }
]

// CALCULATIONS (for other constants)
for (const run of formulaRuns) {
  let count = 0
  for (let formula of run.formulaValues) {
    if (formula !== null) { count++ }
  }
  run.numFormulas = count
}
/** the first column to grab for formulas */
let startFormulaCol = Infinity
/** the last column to grab for formulas */
let endFormulaCol = -1
/** the number of columns to grab for formulas (max formula col - min formula col + 1) */
let numFormulaCols
for (let run of formulaRuns) {
  if (run.startCol < startFormulaCol) {
    startFormulaCol = run.startCol
  }
  let thisEndCol = run.startCol + run.formulaValues.length - 1
  if (thisEndCol > endFormulaCol) {
    endFormulaCol = thisEndCol
  }
}
numFormulaCols = endFormulaCol - startFormulaCol + 1
// Logger.log(`start: ${startFormulaCol}, end: ${endFormulaCol}, num: ${numFormulaCols}`)

// TRIGGERS
/**
 * The event handler triggered when editing the spreadsheet
 * Also called when manually triggered or when the data entry app changes data in the database
 * 
 * The parameter is an object with the following properties:
 *    e.range (Range object for the cell or cells which were edited)
 *    e.source (Spreadsheet object (the google sheets file to which this script is bound))
 *    e.user (User object of active user, if available)
 * If the edited range is a single cell, the event object also has:
 *    e.oldValue (cell value prior to edit)
 *    e.value (new cell value after the edit)
 * If this was called by another function:
 *    e.caller (where this was called from, if available)
 * 
 * @param {Event} e The onEdit event.
 */
function installableOnEdit (e)
{
  // first determine if this we should be checking for block status changes in this sheet
  let sheet = e.range.getSheet()
  let thisSheetName = sheet.getName()
  // check if we should look for status changes in this sheet
  if (checkSheets.includes(thisSheetName)) {
    let startCol = e.range.getColumn()
    let numCols = e.range.getNumColumns()
    let statusCol = letterToColumn(statusColumn)
    // determine if the status column is in the range of edited cells
    if (statusCol >= startCol && statusCol < startCol + numCols) {
      // then update the formula values in the entire range accordingly
      // log type of update, if available
      if (e.caller) {
        Logger.log(`Starting update from ${e.caller}`)
      } else {
        Logger.log('Starting update from manual edit (or unknown source)')
      }
      // log the user's email if available
      if (e.user && e.user.getEmail() !== "") {
        Logger.log(`This edit was made by ${e.user.getEmail()}`)
      } else {
        Logger.log('(unable to determine which user made this edit)')
      }
      let startRow = e.range.getRow() // 1-based, for use with getValues, e.g.
      let numRows = e.range.getNumRows() // 1-based, for use with getValues, e.g.
      // Logger.log(`startRow: ${startRow}, numRows: ${numRows}`)
      let formulaCount = 0
      let staticCount = 0
      let newStaticCount = 0
      let newFormulaCount = 0
      // get the statuses, formulas, and values of the entire range at once
      // this is MUCH faster than many small requests
      let allStatuses = sheet.getRange(startRow, statusCol, numRows, 1).getValues()
      let allCurrentFormulas = sheet.getRange(startRow, startFormulaCol, numRows, numFormulaCols).getFormulas()
      let allValues = sheet.getRange(startRow, startFormulaCol, numRows, numFormulaCols).getValues()
      // loop through the rows in this range
      for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
        let row = startRow + rowOffset // 1-based, for use with getValues, e.g.
        if (row < dataStartRow) { continue } // ignore changes to headers
        // loop through the connected runs of formulas to update
        for (let i = 0; i < formulaRuns.length; i++) {
          let status = allStatuses[rowOffset][0]
          let run = formulaRuns[i]
          let numFormulas = run.numFormulas
          let setRange = sheet.getRange(row, run.startCol, 1, run.formulaValues.length) // cells in this run of formulas
          let statusCell = sheet.getRange(row, statusCol) // range containing status cell for this row (for setting note)
          let currentFormulas = allCurrentFormulas[rowOffset]
            .slice(run.startCol - startFormulaCol, run.startCol + run.formulaValues.length - startFormulaCol)
          let currentValues = allValues[rowOffset]
            .slice(run.startCol - startFormulaCol, run.startCol + run.formulaValues.length - startFormulaCol)
          // check if this status should have formulas or static values
          if (addFormulaStatuses.includes(status)) {
            // Logger.log(`determined status ${status} was in ${addFormulaStatuses}`)
            // check if any cells which should have formulas are missing formulas
            let addFormulas = false
            for (let z = 0; z < run.formulaValues.length; z++) {
              if (run.formulaValues[z] !== null && currentFormulas[z] === '') {
                addFormulas = true
                break
              }
            }
            // if any cell which should have a formula does not have a formula, write formulas for this range
            if (addFormulas) {
              let formulas = []
              // fill in formula strings (or null) for this run of formulas
              for (let j = 0; j < run.formulaValues.length; j++) {
                let formula
                if (run.formulaValues[j] !== null) {
                  // then replace all instances of "@row" in the formula with this range's row number
                  formula = ''
                  let split = run.formulaValues[j].split("@row")
                  formula += split[0]
                  for (let k = 1; k < split.length; k++) {
                    formula += row + split[k]
                  }
                } else {
                  // set formula to null if we don't want this cell to have a formula
                  // this clears the cell, which allows another formula's values to be displayed here
                  formula = null
                }
                formulas.push(formula)
              }
              setRange.setFormulas([formulas]) // setFormulas wants a 2D array of values to set
              Logger.log(`Sheet ${thisSheetName}, row ${row}: added formulas (${i + 1}/${formulaRuns.length})`)
              newFormulaCount += numFormulas
            } else {
              // skip adding formulas since there are formulas there already (assume that they are correct)
              // Logger.log(`Sheet ${thisSheetName}, row ${row}: skipped adding formulas (${i + 1}/${formulaRuns.length})`)
            }
            formulaCount += numFormulas
          } else {
            // Logger.log(`determined status ${status} was NOT in ${addFormulaStatuses}`)
            // check if any formulas are present
            let replaceValues = false
            for (let z = 0; z < run.formulaValues.length; z++) {
              if (currentFormulas[z] !== '') {
                // Logger.log(`found nonempty formula; replacing values (${currentFormulas[z]} !== '')`)
                replaceValues = true
                break
              }
            }
            // if any cell has a formula, then freeze all values for this range in place
            if (replaceValues) {
              // then replace formulas with static values (whatever this cell is or whatever its function is evaluated to right now)
              // let values = setRange.getValues()
              setRange.setValues([currentValues])
              Logger.log(`Sheet ${thisSheetName}, row ${row}: replaced formulas with static values (${i + 1}/${formulaRuns.length})`)
              newStaticCount += numFormulas
            } else {
              // skip replacing values since there were no formulas in this range
              // Logger.log(`Sheet ${thisSheetName}, row ${row}: skipped replacing values (${i + 1}/${formulaRuns.length})`)
            }
            staticCount += numFormulas
          }
          // add a timestamp for this row to the status column in this row
          statusCell.setNote(timeStamp())
        }
      }
      if (numRows > 1) { Logger.log(`COMPLETE: ${numRows} rows updated (removed ${newStaticCount} formulas, added ${newFormulaCount} formulas)`) }
      if (e.updateAll) { Logger.log(`Now in sheet ${thisSheetName}: ${formulaCount} active formulas, ${staticCount} frozen`) }
    }
  }
  function timeStamp () {
    if (e.caller) {
      return `Formulas were last updated by the script on ${new Date} (edit was from ${e.caller})`
    } else {
      return `Formulas were last updated by the script on ${new Date} (manual edit/unknown source)`
    }
  }
};

/**
 * updates ALL formulas in blocks1364 sheet
 */
function updateAllFormulas () {
  // call onEdit with the entire status column
  // this pretends like the status of every block was changed, which updates all blocks' formulas based on its current status
  let range = SpreadsheetApp.getActive().getSheetByName(blocks1364).getRange(`${statusColumn}${dataStartRow}:${statusColumn}`)
  let e = {range, caller: 'updateAllFormulas', user: Session.getActiveUser(), updateAll: true}
  installableOnEdit(e)
}

/**
 * updates all formulas in specified sheet between startRow and endRow (inclusive)
 * @param {Spreadsheet} spreadsheet
 * @param {Number} startRow
 * @param {Number} endRow
 * @param {User} user
 */
function updateRowFormulas (spreadsheet, startRow, endRow, user) {
  // call onEdit for the status column of the specified rows
  // this pretends like the status of every block was changed
  // Logger.log(`${statusColumn}${startRow}:${statusColumn}${endRow}`)
  let range = spreadsheet.getRange(`${statusColumn}${startRow}:${statusColumn}${endRow}`)
  let e = {range, caller: 'updateRowFormulas', user}
  installableOnEdit(e)
}

// HELPER FUNCTIONS
/**
 * @param {Number} the (one-based) column index
 * @return {String} the alphabetic column representation (e.g. C or AB)
 */
function columnToLetter (column) {
  let temp = ''
  let letter = ''
  while (column > 0) {
    temp = (column - 1) % 26
    letter = String.fromCharCode(temp + 65) + letter
    column = (column - temp - 1) / 26
  }
  return letter
}
/**
 * @param {String} the alphabetic column representation (e.g. C or AB)
 * @return {Number} the (one-based) column index
 */
function letterToColumn (_letter) {
  let letter = _letter.toUpperCase()
  let column = 0
  const length = letter.length
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1)
  }
  return column
}
