/** Function that alerts Theology teacher and/or student of service organization confirmation.
  * Theology teacher will have opportunity to approve service request with feedback.
  * This response will be reflected on the active sheet by adding "Pending" to the correct cell. 
  */
function sendServiceApproval(){
    // Iterate through data
  for (var i = 0; i < approvalRequest.length; i++){

    // Check to see if the approval request was verified
    var currentRequest = approvalRequest[i].verified;

    if (currentRequest == 'Verified' && approvalRequest[i].approved == ''){
      
      // If so, construct email (with service details, unique ID number, and simple form)
      var emailBody = buildApprovalEmailBody(approvalRequest[i], i);
      var emailSubject = "Service Hour Approval Request - " + approvalRequest[i].name + ' (' + approvalRequest[i].username + ')';
      var emailTo;
      
      switch(approvalRequest[i].theologyTeacher){
        case 'Almazan':
          emailTo = 'salmazan@mustangsla.org';
          break;
        case 'Bronzina':
          emailTo = 'abronzina@mustangsla.org';
          break;
        case 'Granados':
          emailTo = 'dgranados@mustangsla.org';
          break;
        case 'Ibarra':
          emailTo = 'ibarra@mustangsla.org';
          break;
        case 'Meraz':
          emailTo = 'meraz@mustangsla.org';
          break;
        case 'Sierra':
          emailTo = 'lsierra@mustangsla.org';
          break;
        case 'Sifuentes':
          emailTo = 'dsifuentes@mustangsla.org';
          break;
        default:
          emailTo = 'ibarra@mustangsla.org';
          break;
      
      }
      
      // Send email to theology teacher
      MailApp.sendEmail({
        to: emailTo,
        subject: emailSubject,
        htmlBody: emailBody,
      });
      // Mark that row as having already sent email (with "Pending" or something)
      var theRange = sheet.getRange((i + 2), 20, 1);
      theRange.setValue('Pending');
    }  
  }
}


function buildApprovalEmailBody(currentRequest, index){
  
  // ADD UNIQUE SPREADSHEET ID OF VERIFICATION DATA DOCUMENT.  
  var form = FormApp.openById('1zmNbL2wBgKt5hPnOYzfwmgrXgDYFgdQ18BId2FmLGT4');
  
  
  var url = form.getPublishedUrl();
  var htmlBody = '<head><body>'
        + '<p>' + currentRequest.name + ' has received verification from his service organization and requests approval for Christian Service Hours.</p>'
        + '<p>Here are the details of his service:</p>'
        + '<ul>'
        + '<li><b>Service ID Number: ' + (index) + '</b></li>'
        + '<li>Date of Service: ' + currentRequest.dateOfService + '</li>'
        + '<li>Hours of Service: ' + currentRequest.hoursOfService + '</li>'
        + '<li>Service Organization: ' + currentRequest.serviceOrganization + '</li>'
        + '<li>Contact Name: ' + currentRequest.contactPerson + '</li>'
        + '<li>Description of Service: ' + currentRequest.description + '</li></ul><br>'
        + '<li>Description of Service: ' + currentRequest.beneficiaries + '</li></ul><br>'
        + '<li>Description of Service: ' + currentRequest.improvement + '</li></ul><br>'
        + '<li>Description of Service: ' + currentRequest.impact + '</li></ul><br>'

        + "<p>Please submit approval status by visiting <a href='" + url + "'>this link.</a> Use the <b>Service ID Number (" + index + ")</b> to identify the student.</p>"
        + '</body></html>';
  
  return htmlBody;
}

/** This function pulls the approval data from the Theology Teacher (whether 
  * the teacher approved or did not approve the service request) and responds to these 
  * data by adding it to the active spreadsheet, then emailing students with the 
  * current status of their service request. The function takes a parameter so that
  * it can iterate through all of the items in the verification data spreadsheet.
  */
function getApprovalStatus(idNumber){
  // Get data of Theology Teacher spreadsheet
  var theologyApprovalSheet = SpreadsheetApp.openById('1CIJNsuMw-8XzWrgCf0Mxi52jHlsXnZFDpbIuz0gTlWs').getActiveSheet();
  var approvalData = getRowsData(theologyApprovalSheet);

  // Use ID number to find instance of that ID number 
    // Get range of the column in which ID numbers are located
  for (var i = 0; i < approvalData.length; i++){
    var index = approvalData[i].serviceIdNumber;
    if (approvalData[i].emailSent == ''){
      Logger.log(approvalData[i].approvalStatus);
      var approvalStatus = sheet.getRange((index + 2), 20, 1);
      var explanation = sheet.getRange((index + 2), 21, 1);
      var emailSent = theologyApprovalSheet.getRange((i+2), 5, 1);
      
      
      var emailBody;
      var emailSubject;
      if (approvalData[i].approvalStatus == "Approved"){
        // Mark the approval request sheet with "Approved" or "Not Approved"
        var runningTotal = approvalRequest[index].hoursOfService + approvalRequest[index].runningTotal;
        approvalStatus.setValue('Approved'); 
        emailBody = '<head><body>'
        + '<p>' + approvalRequest[index].name + ',</p>'
        + '<p>Your Theology teacher, has <b>APPROVED</b> your Christian Service Hour submission, below:</p>'
        + '<ul>'
        + '<li>Date of Service: ' + approvalRequest[index].dateOfService + '</li>'
        + '<li>Hours of Service: ' + approvalRequest[index].hoursOfService + '</li>'
        + '<li>Service Organization: ' + approvalRequest[index].serviceOrganization + '</li>'
        + '<li>Contact Name: ' + approvalRequest[index].contactPerson + '</li>'
        + '<li>Description of Service: ' + approvalRequest[index].description + '</li></ul><br>'
        + '<li>Description of Service: ' + approvalRequest[index].beneficiaries + '</li></ul><br>'
        + '<li>Description of Service: ' + approvalRequest[index].improvement + '</li></ul><br>'
        + '<li>Description of Service: ' + approvalRequest[index].impact + '</li></ul><br>'        
        + '<p>At this time, you have performed <b>' + runningTotal + ' hours </b> of approved Christian Service.</p>'
        + '</body></html>';
        emailSubject = 'Your Theology Teacher has APPROVED your Christian Service Hour submission.';
        
        MailApp.sendEmail({
          to: approvalRequest[index].username,
          subject: emailSubject,
          htmlBody: emailBody,
        });
        Logger.log(approvalData[i].serviceIdNumber + ': Approval Sent');

      }
      else {
        approvalStatus.setValue('Not Approved');
        explanation.setValue(approvalData[i].explanation);
        emailBody = '<head><body>'
        + '<p>' + approvalRequest[index].name + ',</p>'
        + '<p>Your Theology teacher, has <b>NOT APPROVED</b> your Christian Service Hour submission, below:</p>'
        + '<ul>'
        + '<li>Date of Service: ' + approvalRequest[index].dateOfService + '</li>'
        + '<li>Hours of Service: ' + approvalRequest[index].hoursOfService + '</li>'
        + '<li>Service Organization: ' + approvalRequest[index].serviceOrganization + '</li>'
        + '<li>Contact Name: ' + approvalRequest[index].contactPerson + '</li>'
        + '<li>Description of Service: ' + approvalRequest[index].description + '</li></ul><br>'
        + '<li>Description of Service: ' + approvalRequest[index].beneficiaries + '</li></ul><br>'
        + '<li>Description of Service: ' + approvalRequest[index].improvement + '</li></ul><br>'
        + '<li>Description of Service: ' + approvalRequest[index].impact + '</li></ul><br>'         
        + '<p>Your teacher gave these reasons for not approving this submission:</p>'
        + '<p>' + approvalRequest[index].explanation + '</p><br>'
        + '<p>At this time, you have performed <b>' + approvalRequest[index].runningTotal + ' hours </b> of approved Christian Service.</p>'
        
        + '</body></html>';
        emailSubject = 'Your Theology Teacher has NOT APPROVED your Christian Service Hour submission.';
        
        MailApp.sendEmail({
          to: approvalRequest[index].username,
          subject: emailSubject,
          htmlBody: emailBody,
        });
        Logger.log(approvalData[i].serviceIdNumber + ': Not Approval Sent');
        Logger.log(approvalRequest[index].explanation);
        
      }

      emailSent.setValue('Sent');
      
    }
    else {
      Logger.log(approvalData[i].serviceIdNumber + ': Does not meet reqs');
    }
  }
}

/** This function iterates through all data in the approval spreadsheet
  * and calls the getApprovalStatus function to act on all approval 
  * data that has not been acted upon yet.
*/
function getAllApprovalStatus(){
  // Get data of Theology Teacher spreadsheet
  var theologyApprovalSheet = SpreadsheetApp.openById('1CIJNsuMw-8XzWrgCf0Mxi52jHlsXnZFDpbIuz0gTlWs').getActiveSheet();
  var approvalData = getRowsData(theologyApprovalSheet);
  
  for (var i = 0; i < approvalData.length; i++){
    getApprovalStatus(i);
  }
}
