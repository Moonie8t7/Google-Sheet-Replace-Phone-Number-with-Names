/**
 * @author u/IAmMoonie <https://www.reddit.com/user/IAmMoonie/>
 * @file https://www.reddit.com/r/sheets/comments/10maqi8/anyway_to_automate_find_and_replace/
 * @desc Replaces all instances of a phone number with a corresponding name on the active sheet.
 * @license MIT
 * @version 1.0
 */

/**
 * A dictionary of phone numbers to names.
 *
 * @typedef {Object} PHONENUMBER_TO_NAME
 * @property {string} "123-555-1289" - The phone number
 * @property {string} "Luke Skywalker" - The corresponding name
 *
 * @example
 * // To add a new phone number and name to the dictionary
 * "999-999-9999": "Obi-Wan Kenobi"
 */
const PHONENUMBER_TO_NAME = {
  "123-555-1289": "Luke Skywalker",
  "555-555-5678": "Frodo Baggins",
  "555-555-4321": "Aragorn",
  "555-555-1111": "Darth Vader",
  "555-555-2222": "Gollum"
};

/**
 * When the spreadsheet is opened, a menu is created with two items: "Replace Numbers With Names" and
 * "Help".
 * @function onOpen
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("▶︎ Scripts")
    .addItem("Replace Numbers With Names", "replaceNumberWithName_")
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Help")
        .addItem("Instructions", "showInstructions_")
    )
    .addToUi();
}

/**
 * It shows a dialog box with instructions on how to use the script
 * @function showInstructions_
 */
function showInstructions_() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "Instructions: \n\n1. To update the dictionary of phone numbers to names, please open the script editor and modify the PHONENUMBER_TO_NAME object. \n\n2. Make sure to save the changes before running the script again. \n\n3. To run the script, go to the ▶︎ Scripts menu and select 'Replace Numbers With Names'"
  );
}

/**
 * Replaces phone numbers with names in the active sheet.
 * @function replaceNumberWithName_
 */
function replaceNumberWithName_() {
  try {
    /* Getting the active sheet, getting the range of data in the sheet, and then getting the values of
    the range. */
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    let values = range.getValues();

    /* Getting the keys of the PHONENUMBER_TO_NAME object. */
    const phoneNumbers = Object.keys(PHONENUMBER_TO_NAME);

    /* Creating a regular expression that will match any of the phone numbers in the
    PHONENUMBER_TO_NAME object. */
    const phoneNumberRegex = new RegExp(
      `(\\+?\\d{0,2}[-\\s]?)?(${phoneNumbers.join(
        "|"
      )})(\\d|\\s|\\+|\\-|\\(|\\))*`,
      "g"
    );

    /* Replacing the phone numbers with the corresponding name. */
    values = values.map((row) =>
      row.map((cell) =>
        typeof cell === "string"
          ? cell.replace(phoneNumberRegex, (match) => {
              let matchPhoneNumber = match.replace(/\D/g, "");
              return PHONENUMBER_TO_NAME[
                phoneNumbers.find((phoneNumber) =>
                  matchPhoneNumber.endsWith(phoneNumber.replace(/\D/g, ""))
                )
              ];
            })
          : cell
      )
    );
    /* Setting the values of the range to the values that were just modified. */
    range.setValues(values);
    SpreadsheetApp.flush();
  } catch (error) {
    /* Logging the error to the console and then alerting the user that an error occurred. */
    console.error(error);
    SpreadsheetApp.getUi().alert(
      "An error occurred while running the script. Please check the logs."
    );
  }
}
