//Funcions that changing history sheet

/**
 * Deletes all the information contained in SS_HISTORY
 * @see {@link SS_HISTORY}
 */
function deleteAllHistory() {
  SS_HISTORY.getRange(HISTORY_WEEKS_ROW, HISTORY_MEMBER_COL, 500, 500).deleteCells(SpreadsheetApp.Dimension.ROWS);
}
