
/**
 *  Sort function
 */

function sorting(e) {
  
  const row = e.range.getRow()
  const column = e.range.getColumn()
  const ss = e.source
  const cs = ss.getActiveSheet()
  const cs1 = cs.getSheetName()
  const lastRow = cs.getLastRow();
  
  
  //  only sort if you're on the sheet named Sheet1 and row is not the header row and the column is on column 8

   if (!(cs1 === "Sheet1" && column === 8 && row >= 2)) return

    // getRange refers to 2nd row, 2nd column, 

    const range = cs.getRange(2,2,cs.getLastRow()-1,7)

    // sort column 8 in descening order (largest to smallest), followe by column 6 in ascending order (smallest to largest)

    range.sort([
      {column: 8, ascending: false},
      {column: 6, ascending: true}
      ])
}

/**
 *  When the spreadsheet is edited, run the sort function
 */

function onEdit(e) {
  sorting(e)
}


