 /**
 * The admissionsObject
 * @param {Object} par The main parameter object.
 * @return {Object} The appObject Object.
 */
function admissionsObject(par) {
  "use strict";
  var objectName = "admissionObject";
  var sendEmails = par.sendEmails;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //Set up the dynamic cell and col references
  var configRange = ss.getSheetByName("Configuration").getRange("A:G");
  var RunCol = configRange.getCell(2,7).getValue();
  var VerifiedApplicantsSheet = configRange.getCell(3,3).getValue();
  var VolatileRandomOrderSourceCol = configRange.getCell(3,7).getValue();
  var VolatileRandomOrderTargetCol = configRange.getCell(4,7).getValue();
  var ParentRecordCol = configRange.getCell(5,7).getValue();
  var ControlSheet = configRange.getCell(6,3).getValue();
  var UnlockedCell = configRange.getCell(6,7).getValue();
  var CloseAcceptanceCell = configRange.getCell(7,7).getValue();
  var SelectionProcessStartedCell = configRange.getCell(8,7).getValue();
  var RunSelectionLastStateCell = configRange.getCell(9,7).getValue();
  var CancelOffersLastStateCell = configRange.getCell(10,7).getValue();
  var DateRunCell = configRange.getCell(11,7).getValue();
  var SpreadSheetIdCell = configRange.getCell(12,7).getValue();
  var SheetIdCell = configRange.getCell(13,7).getValue();
  var OfferedCol = configRange.getCell(14,7).getValue();
  var OfferAcceptedCol = configRange.getCell(15,7).getValue();
  var VerifiedCategoryCol = configRange.getCell(16,7).getValue();
  var RandomOrderCol = configRange.getCell(17,7).getValue();
  var EmailCol = configRange.getCell(18,7).getValue();
  var ApplicantSheet = configRange.getCell(19,3).getValue();
  var ApplicantOfferAcceptedCol = configRange.getCell(19,7).getValue();
  var ApplicantOfferAcceptedCell = configRange.getCell(20,7).getValue();
  var VolatileRandomOrderTargetCell = configRange.getCell(21,7).getValue();
  var ProcessSheet = configRange.getCell(22,3).getValue();
  var StartSelectionProcessDoneDateCell = configRange.getCell(22,7).getValue();
  var RunSelectionProcessDoneCell = configRange.getCell(23,7).getValue();
  var RunSelectionProcessDoneDateCell = configRange.getCell(24,7).getValue();
  var CloseAcceptanceDoneCell = configRange.getCell(25,7).getValue();
  var CloseAcceptanceDoneDateCell = configRange.getCell(26,7).getValue();
  var AdmissionsNoticeSheet = configRange.getCell(27,3).getValue();
  var EmailFromCell = configRange.getCell(27,7).getValue();
  var CloseOfferEmailSubjectCell = configRange.getCell(28,7).getValue();
  var CloseOfferEmailBodyCell = configRange.getCell(29,7).getValue();
  var RunSelectionEmailSubjectCell = configRange.getCell(30,7).getValue();
  var RunSelectionEmailBodyCell = configRange.getCell(31,7).getValue();
  var CommunicationsSheet = configRange.getCell(32,3).getValue();
  var ApplicantApplicationWithdrawnCol = configRange.getCell(33,7).getValue();
  var StartSelectionAllowReRunCell = configRange.getCell(34,7).getValue();
  var RunSelectionAllowReRunCell = configRange.getCell(35,7).getValue();
  var CloseOfferAllowReRunCell = configRange.getCell(36,7).getValue();
  var SendEmailCell = configRange.getCell(37,7).getValue();
  var TreatSiblingsJointlyCell = configRange.getCell(38,7).getValue();
  var ApplicantOfferWithdrawnCol = configRange.getCell(39,7).getValue();
  var ApplicantOfferWithdrawnCell = configRange.getCell(40,7).getValue();
  var ApplicantOfferFromOtherSchoolCol = configRange.getCell(41,7).getValue();
  var ApplicantOfferFromOtherSchoolCell = configRange.getCell(42,7).getValue();
  
  /**
  * run some code
  */
  function run(e) {
    // code here
    var thisSheet = ss.getSheetByName(ControlSheet);
    var unlocked = thisSheet.getRange(UnlockedCell).isChecked();
    var closeAcceptance = thisSheet.getRange(CloseAcceptanceCell).isChecked();
    var run = thisSheet.getRange(RunCol).isChecked();
    var runLastDate = thisSheet.getRange(DateRunCell).getValue();
    var runSelectionLastState = thisSheet.getRange(RunSelectionLastStateCell).getValue();
    var cancelOffersLastState = thisSheet.getRange(CancelOffersLastStateCell).getValue();
    var treatSiblingsJointly = thisSheet.getRange(TreatSiblingsJointlyCell).getValue();
    
    //automatically move any withdrawn applications to the correct verified category
    runApplicationWithdrawn();
    
    //clear previous ranking and offers when all applications have been unverified
    if (!thisSheet.getRange(SelectionProcessStartedCell).isChecked()) {
      resetProcess()
    }
    
    //clear update process timestamps
    var processCheckCol = RunSelectionProcessDoneCell.substring(0,1);
    var processDateCol = RunSelectionProcessDoneDateCell.substring(0,1);
    for(var i = 0; i < 3; i++) {
      if (ss.getSheetByName(ProcessSheet).getRange(processCheckCol+(i+2)).isChecked()) {   
        updateProcessTimestamp(processDateCol+(i+2))
      }
    }

    //clear the previous ranks and offers if we are closing the acceptance window
    if (unlocked && closeAcceptance && (cancelOffersLastState!=true)) {
      //Logger.log("inside");
      thisSheet.getRange(CancelOffersLastStateCell).setValue("TRUE")
      runApplicationAcceptanceClose();
      if (sendEmails){
        //send communications to everyone who was offered a place
        var emailFrom =  ss.getSheetByName(AdmissionsNoticeSheet).getRange(EmailFromCell).getValue()
        var emailSubject = ss.getSheetByName(ProcessSheet).getRange(CloseOfferEmailSubjectCell).getValue();
        var emailBody = ss.getSheetByName(ProcessSheet).getRange(CloseOfferEmailBodyCell).getValue();
        //sheetName, triggerColumn, triggerValue, emailFrom, emailBcc, emailSubject, emailBody
        sendCommunication(VerifiedApplicantsSheet, "VerifiedCategory", "Offer Not Accepted", emailFrom, "", emailSubject, emailBody)
      }      
    }

    //Reset the runSelection Process so we can run it again
    if (unlocked && run!=true && (runSelectionLastState==true)) {
      //Logger.log("We need to reset runSelection Process");
      thisSheet.getRange(RunSelectionLastStateCell).setValue("FALSE")
    }
    
    //Reset the cancelOffer Process so we can run it again
    if (unlocked && closeAcceptance!=true && (cancelOffersLastState==true)) {
      //Logger.log("We need to reset closeAcceptance");
      thisSheet.getRange(CancelOffersLastStateCell).setValue("FALSE")
    }

    if (unlocked && run && (runSelectionLastState!=true)) {
      thisSheet.getRange(RunSelectionLastStateCell).setValue("TRUE")
      //Logger.log(treatSiblingsJointly);

      runApplicationRank(treatSiblingsJointly);

      //add timestamp
      var formattedDate = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy HH:mm");
      thisSheet.getRange(DateRunCell).setValue(formattedDate).setNumberFormat("dd/MM/yyyy HH:mm")
      //get the sheet id using .getId() and .getSheetId() and place in cells
      thisSheet.getRange(SpreadSheetIdCell).setValue(thisSheet.getParent().getId())
      thisSheet.getRange(SheetIdCell).setValue(ss.getSheetByName(VerifiedApplicantsSheet).getSheetId())
      
      if (sendEmails){
        //send communications to everyone who was offered a place
        var emailFrom =  ss.getSheetByName(AdmissionsNoticeSheet).getRange(EmailFromCell).getValue()
        var emailSubject = ss.getSheetByName(ProcessSheet).getRange(RunSelectionEmailSubjectCell).getValue();
        var emailBody = ss.getSheetByName(ProcessSheet).getRange(RunSelectionEmailBodyCell).getValue();
        //sheetName, triggerColumn, triggerValue, emailFrom, emailBcc, emailSubject, emailBody
        sendCommunication(VerifiedApplicantsSheet, "Offered Place", "YES", emailFrom, "", emailSubject, emailBody)
        sendCommunication(VerifiedApplicantsSheet, "Offered Place", "NO", emailFrom, "", emailSubject, emailBody)
      }
    }
  }

  function runApplicationRank(treatSiblingsJointly){  
    var sourceRange = ss.getSheetByName(VerifiedApplicantsSheet).getRange(VolatileRandomOrderSourceCol);
    var targetRange = ss.getSheetByName(VerifiedApplicantsSheet).getRange(VolatileRandomOrderTargetCol);
    
    //copy over the volatile randown sort so it is saved
    sourceRange.copyTo(targetRange, {contentsOnly:true});
        
    //copy over any sibling rank
    if (treatSiblingsJointly == true) { //if the settings say to treat sibling together like twins
      var ranks = targetRange.getValues() ;
      var parents = ss.getSheetByName(VerifiedApplicantsSheet).getRange(ParentRecordCol).getValues();
      //Browser.msgBox(parents);
      
      for (var row = 0, numRows = parents.length; row < numRows; row++) {
        switch (row) {
          case 0: // row numbers are zero-indexed in the values array
            break;
          default:
            // the row is not one of the ones we are interested in
            if (parents[row] != "" && row!=parents[row]-1){
              //Browser.msgBox(row);
              var duplicatedRank = targetRange.getCell(parents[row],1).getValue();
              //copy over the rank
              targetRange.getCell(row+1,1).setValue(duplicatedRank);
            }
        } // switch
      } // row
    }
  }

  function runApplicationWithdrawn(){  
    var sourceRange = ss.getSheetByName(VerifiedApplicantsSheet).getRange(ApplicantApplicationWithdrawnCol);
    var targetRange = ss.getSheetByName(VerifiedApplicantsSheet).getRange(VerifiedCategoryCol);
        
    var withdraws = sourceRange.getValues() ;
    
    for (var row = 0, numRows = withdraws.length; row < numRows; row++) {
      switch (row) {
        case 0: // row numbers are zero-indexed in the values array
          break;
        default:
          // the row is not one of the ones we are interested in
          if (withdraws[row][0] == true){
            //Logger.log("Application Withdrawn");
            //copy over the category
            targetRange.getCell(row+1,1).setValue("Application Withdrawn");
          }
      } // switch
    } // row
  }

  
  function columnToLetter(column)
  {
    var temp, letter = '';
    while (column > 0)
    {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }
  
  function getColByName(sheetName, columnName){
    var headers = ss.getSheetByName(sheetName).getDataRange().getValues().shift();
    var colindex = headers.indexOf(columnName);
    return colindex+1;
  }
  
  function sendCommunication(sheetName, triggerColumn, triggerValue, emailFrom, emailBcc, emailSubject, emailBody) {
    //e.g triggerColumn - triggerValue
    //Offered Place - YES
    //Waiting List Position > 0
    //VerifiedCategory - Offer Not Accepted  
    
    var colLetter = columnToLetter(getColByName(sheetName, triggerColumn));
    
    var checkRange = ss.getSheetByName(sheetName).getRange(colLetter+":"+colLetter);
    var emailValues = ss.getSheetByName(sheetName).getRange(EmailCol).getValues();
    
    var haystack = checkRange.getValues() ;
    
    for (var row = 0, numRows = haystack.length; row < numRows; row++) {
      switch (row) {
        case 0: // row numbers are zero-indexed in the values array
          break;
        default:
          // the row is not one of the ones we are interested in
          if (haystack[row] == triggerValue){
            //Logger.log(emailValues[row][0], emailSubject,emailBody); 
            var nextRow = ss.getSheetByName(CommunicationsSheet).getLastRow() + 1;
            ss.getSheetByName(CommunicationsSheet).appendRow([nextRow-1, emailValues[row][0], emailFrom, emailBcc, emailSubject, emailBody])
          }
      } // switch
    } // row
    
  }
  
  
  function runApplicationAcceptanceClose(){  
    var sourceRange = ss.getSheetByName(VerifiedApplicantsSheet).getRange(OfferAcceptedCol);
    var targetRange = ss.getSheetByName(VerifiedApplicantsSheet).getRange(VerifiedCategoryCol);
    var emailRange = ss.getSheetByName(VerifiedApplicantsSheet).getRange(EmailCol);
    var randomOrderRange = ss.getSheetByName(VerifiedApplicantsSheet).getRange(RandomOrderCol);
    
    //get accepted offered
    var offerAcceptanceStatus = sourceRange.getValues();
    var offerEmails = emailRange.getValues();
    
    //get if place was offered
    var offers = ss.getSheetByName(VerifiedApplicantsSheet).getRange(OfferedCol).getValues();
    
    for (var row = 0, numRows = offerAcceptanceStatus.length; row < numRows; row++) {
      switch (row) {
        case 0: // row numbers are zero-indexed in the values array
          break;
        default:
          // the row is not one of the ones we are interested in
          //Logger.log("row"+row)
          var cellIsChecked = offerAcceptanceStatus[row];
          if (cellIsChecked!="true" && offerEmails[row]!="" && offers[row]=="YES"){
            //if offer isnt accept replace verified category with "Offer Not Accepted"
            targetRange.getCell(row+1,1).setValue("Offer Not Accepted");
            randomOrderRange.getCell(row+1,1).setValue(0);
          }
      } // switch
    } // row
  }
  
  function resetProcess(){  
    var thisSheet = ss.getSheetByName(ControlSheet);
    
    //reset the accepted applications
    ss.getSheetByName(ApplicantSheet).getRange(ApplicantOfferAcceptedCol).clear({contentsOnly: true});
    ss.getSheetByName(ApplicantSheet).getRange(ApplicantOfferAcceptedCell).setValue("Offer Accepted");

    //reset all the withdrawal and offerFromOtherSchool
    ss.getSheetByName(ApplicantSheet).getRange(ApplicantOfferWithdrawnCol).clear({contentsOnly: true});
    ss.getSheetByName(ApplicantSheet).getRange(ApplicantOfferWithdrawnCell).setValue("Application Withdrawn");
    ss.getSheetByName(ApplicantSheet).getRange(ApplicantOfferFromOtherSchoolCol).clear({contentsOnly: true});
    ss.getSheetByName(ApplicantSheet).getRange(ApplicantOfferFromOtherSchoolCell).setValue("Offer From Another School");
  
    //reset the run toggle
    thisSheet.getRange(RunSelectionLastStateCell).setValue("FALSE")
    
    //reset the cancel offers toggle
    thisSheet.getRange(CancelOffersLastStateCell).setValue("FALSE")
    
    //reset the ranking
    ss.getSheetByName(VerifiedApplicantsSheet).getRange(VolatileRandomOrderTargetCol).clear({contentsOnly: true});
    ss.getSheetByName(VerifiedApplicantsSheet).getRange(VolatileRandomOrderTargetCell).setValue("Volatile Random Order")
    
    //reset all the process steps
    ss.getSheetByName(ProcessSheet).getRange(StartSelectionProcessDoneDateCell).setValue("Updating...");
    ss.getSheetByName(ProcessSheet).getRange(RunSelectionProcessDoneCell).setValue("FALSE")
    ss.getSheetByName(ProcessSheet).getRange(RunSelectionProcessDoneDateCell).setValue("Updating...");
    ss.getSheetByName(ProcessSheet).getRange(CloseAcceptanceDoneCell).setValue("FALSE")
    ss.getSheetByName(ProcessSheet).getRange(CloseAcceptanceDoneDateCell).setValue("Updating...");
    
    
  }
  
  function updateProcessTimestamp(cell){
    var formattedDate = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy HH:mm");
    if (ss.getSheetByName(ProcessSheet).getRange(cell).getValue() == "Updating...") {
      ss.getSheetByName(ProcessSheet).getRange(cell).setValue(formattedDate).setNumberFormat("dd/MM/yyyy HH:mm");
    }
  }
  
  
  function sendEmailsPeriodically(){
    var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    Logger.log("Remaining email quota: " + emailQuotaRemaining);
    var emailsSentThisTime = 0; //initialise this amount of emails
    var runningTotalOfEmailsSent = 0;
    //emailQuotaRemaining = 100; //just for testing
    
    var range = ss.getSheetByName(CommunicationsSheet).getRange("A:G");
    var results = range.getDisplayValues();
    var array = [];
    var formattedDate = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy HH:mm")
    
    var rowNumber = 0
    results.forEach(function(row){ 
      if (rowNumber != 0) {//ignore header
        if (row[6] == "" && !isNaN(row[0]) ) { //filled rows
          if (runningTotalOfEmailsSent<emailQuotaRemaining){ //only send if we havent exceeded quota
            range.getCell(rowNumber+1,7).setValue(formattedDate).setNumberFormat("dd/MM/yyyy HH:mm");
            //Logger.log("Email To: " + row[1] + row[4] + row[5])
            MailApp.sendEmail(row[1], row[4], row[5], {htmlBody: row[5]});
            runningTotalOfEmailsSent++;
          }
        }
      }
      rowNumber++;
    });     
  }
  
  return Object.freeze({
    objectName: objectName,
    run: run,
    sendEmailsPeriodically: sendEmailsPeriodically
  });
}
