var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();

var approvalRequest = getRowsData(sheet);


/** Function that alerts the service organization of the approval request. 
  * The service organization can confirm or deny the request. This response 
  * will be reflected on the active sheet. 
  * 
  * This function is triggered by a new Christian Service Hour Approval Request by the student.
  */
function sendServiceVerification(){
  // Iterate through data
  for (var i = 0; i < approvalRequest.length; i++){
    // Check to see if Service Org email has been sent
    var currentRequest = approvalRequest[i].verified;

    if (currentRequest == '' ){
      
      // If not, construct email (with service details, unique ID number, and simple form)
      var emailBody = buildVerifEmailBody(approvalRequest[i], i);
      var emailSubject = "Service Hour Verification Request - " + approvalRequest[i].name + ' (' + approvalRequest[i].username + ')';

      // Send email to contact email      
      MailApp.sendEmail({
        to: approvalRequest[i].contactEmail,
        subject: emailSubject,
        htmlBody: emailBody,
      });
      // Mark that row as having already sent email (with "Pending" or something)
      var theRange = sheet.getRange((i + 2), 19, 1);
      theRange.setValue('Pending');
    }  
  } 
}

/** This function builds the email body that is sent to the service organization
  * contact person so that they can verify or deny that the student in fact
  * completed these hours. The email will contain service details and a Service
  * ID Number, which the service organization will use to identify this particular submission.
  */
function buildVerifEmailBody(currentRequest, index){
  var form = FormApp.openById('1l4wIM_pKCEwVGoKUaE4k1IqeUwqs6dHjCQF6xjZ7vp8');
  var url = form.getPublishedUrl();
  var htmlBody = '<head><body>'
    + '<img style="height:100px;width:100px;" src="http://salesian.schoolwires.net/cms/lib03/CA02001206/Centricity/Domain/2/SHS%20Seal%20New%20600x600.jpg"><br/>'
        + '<p>' + currentRequest.name + ' has requested verification for Christian Service Hours.</p>'
        + '<p>Here are the details of his service:</p>'
        + '<ul>'
        + '<li><b>Service ID Number: ' + (index) + '</b></li>'
        + '<li>Date of Service: ' + currentRequest.dateOfService + '</li>'
        + '<li>Hours of Service: ' + currentRequest.hoursOfService + '</li>'
        + '<li>Service Organization: ' + currentRequest.serviceOrganization + '</li>'
        + '<li>Contact Name: ' + currentRequest.contactPerson + '</li>'
        + '<li>Description of Service: ' + currentRequest.description + '</li></ul><br>'
        + '<p>Please verify that these details are correct in order for ' + currentRequest.name + ' to receive credit for his service hours.</p>'
        + "<p>Verify by visiting <a href='" + url + "'>this link.</a> Use the <b>Service ID Number (" + index + ")</b> to identify the student.</p>"
        + '</body></html>';
  
  return htmlBody;
}

/** This function pulls the verification data from the service organization (whether 
  * the contact person verified or denied the service request) and responds to these 
  * data by adding it to the active spreadsheet, then emailing students with the 
  * current status of their service request. The function takes a parameter so that
  * it can iterate through all of the items in the verification data spreadsheet.
  */
function getVerificationStatus(idNumber){
  // Get data of Service Org spreadsheet
  
  // ADD UNIQUE SPREADSHEET ID OF VERIFICATION DATA DOCUMENT.
  var serviceOrgSheet = SpreadsheetApp.openById('1o6cdcAuP9aZ5d-Ducg2b-1prFAnqx5G3qNhahUT95wc').getActiveSheet();
  
  var verificationData = getRowsData(serviceOrgSheet);
//  Logger.log(verificationData);
  // Use ID number to find instance of that ID number 
    // Get range of the column in which ID numbers are located
  for (var i = 0; i < verificationData.length; i++){
    var index = verificationData[i].serviceIdNumber;
    var verificationStatus = sheet.getRange((index), 19, 1);
    if (verificationData[i].emailSent == ''){
      Logger.log(verificationData[i].verification);
      
      var emailSent = serviceOrgSheet.getRange((i+2), 4, 1);
      var emailBody;
      var emailSubject;
      
      if (verificationData[i].verification == "Yes"){
        // Mark the approval request sheet with "Verified" or "Denied"
        
        verificationStatus.setValue('Verified'); 
        emailBody = '<head><body>'
          + '<img style="height:100px;width:100px;" src="http://salesian.schoolwires.net/cms/lib03/CA02001206/Centricity/Domain/2/SHS%20Seal%20New%20600x600.jpg"><br/>'
        + '<p>' + approvalRequest[index].name + ',</p>'
        + '<p>' + approvalRequest[index].contactPerson + ' has verified the details of your Christian Service Hour submission.</p>'
        + '<p>Here are the details of your submission:</p>'
        + '<ul>'
        + '<li><b>Service ID Number: ' + verificationData[i].serviceIdNumber + '</b></li>'
        + '<li>Date of Service: ' + approvalRequest[index].dateOfService + '</li>'
        + '<li>Hours of Service: ' + approvalRequest[index].hoursOfService + '</li>'
        + '<li>Service Organization: ' + approvalRequest[index].serviceOrganization + '</li>'
        + '<li>Contact Name: ' + approvalRequest[index].contactPerson + '</li>'
        + '<li>Description of Service: ' + approvalRequest[index].description + '</li></ul><br>'        
        + '<p>Your Theology teacher will review your Christian Service Hour submission for approval shortly.</p>'
        + '</body></html>';
        emailSubject = approvalRequest[index].contactPerson + ' has verified your Christian Service Hour submission.';
        
        MailApp.sendEmail({
          to: approvalRequest[idNumber].username,
          subject: emailSubject,
          htmlBody: emailBody,
        });
      }
      else {
        verificationStatus.setValue('Denied');
        emailBody = '<head><body>'
          + '<img style="height:100px;width:100px;" src="http://salesian.schoolwires.net/cms/lib03/CA02001206/Centricity/Domain/2/SHS%20Seal%20New%20600x600.jpg"><br/>'
        + '<p>' + approvalRequest[index].name + ',</p>'
        + '<p>Your recent Christian Service Hour submission has been denied by your service organization.</p>'
        + '<p>Here are the details of your submission:</p>'
        + '<ul>'
        + '<li><b>Service ID Number: ' + verificationData[i].serviceIdNumber + '</b></li>'
        + '<li>Date of Service: ' + approvalRequest[index].dateOfService + '</li>'
        + '<li>Hours of Service: ' + approvalRequest[index].hoursOfService + '</li>'
        + '<li>Service Organization: ' + approvalRequest[index].serviceOrganization + '</li>'
        + '<li>Contact Name: ' + approvalRequest[index].contactPerson + '</li>'
        + '<li>Description of Service: ' + approvalRequest[index].description + '</li></ul><br>'        
        + '<p>Please contact your service organization for further details and then resubmit a new Christian Service Hour request.</p>'
        + '</body></html>';
        emailSubject = 'Your Christian Service Hour submission has been denied.';
        
        MailApp.sendEmail({
          to: approvalRequest[idNumber].username,
          subject: emailSubject,
          htmlBody: emailBody,
        });
      }

      emailSent.setValue('Sent');
      
    }
    else {
      Logger.log("Try Again.");
     if (verificationData[i].verification == "Yes"){
        // Mark the approval request sheet with "Verified" or "Denied"
        
        verificationStatus.setValue('Verified');  
     }
    }
  }
}

/** This function iterates through all data in the verification spreadsheet
  * and calls the getVerificationStatus function to act on all verification 
  * data that has not been acted upon yet.
*/
function getAllVerifStatus(){
  var serviceOrgSheet = SpreadsheetApp.openById('1o6cdcAuP9aZ5d-Ducg2b-1prFAnqx5G3qNhahUT95wc').getActiveSheet();
  var verificationData = getRowsData(serviceOrgSheet);
  for (var i = 0; i < verificationData.length; i++){
    getVerificationStatus(i);
  }
}


/**********************************************************************/

