// Keep standard variable naming ("{{Client Address}} vs {{Address}}, etc)
// Any fees that are filled in, remove the dollar sign from the contract template ("$$3,500")
// Boring Proposal: # of borings & cost per foot? Highlighted portions?

function propGen(e) {
  const sheet = e.source.getActiveSheet();
  if(sheet.getName() == "Proposals"){
    // changes detected in Proposals sheet
    Logger.log("fire: " + sheet.getName());
    mostRecentProposalObject = visquery("lastprop");

    // checks to ensure proposal it's about to create doesn't already exist
    if(propexists(mostRecentProposalObject.propNum + " : " + mostRecentProposalObject.propName)){
      throw new Error("Folder with matching proposal  name, number already exists!");
    }

    // create new folder + subfolder
    const propsdir = DriveApp.getFolderById("/*proposals dir ID*/");
    const newdir = propsdir.createFolder(mostRecentProposalObject.propNum + " : " + mostRecentProposalObject.propName);
    const contractdir = newdir.createFolder("Contract");

    // determine type to copy and generate proper contract
    var contractcopy;
    switch(mostRecentProposalObject.propType){
      case "Authorization to Proceed":
        Logger.log("Authorization to Proceed");
        const atpdoc = DriveApp.getFileById("/*atp temp doc ID*/"); 
        contractcopy = atpdoc.makeCopy(mostRecentProposalObject.propName + " Authorization to Proceed", contractdir);        
        break;
      case "Lump Sum Proposal":
        Logger.log("Lump Sum Proposal");
        const lspdoc = DriveApp.getFileById("/*lsp temp doc ID*/"); 
        contractcopy = lspdoc.makeCopy(mostRecentProposalObject.propName + " Lump Sum Proposal", contractdir);
        break;
      case "Hourly Proposal":
        Logger.log("Hourly Proposal");
        const hourlydoc = DriveApp.getFileById("/*hourly temp doc ID*/"); 
        contractcopy = hourlydoc.makeCopy(mostRecentProposalObject.propName + " Hourly Proposal", contractdir);
        break;
      case "Test Boring Proposal":
        Logger.log("Test Boring Proposal");
        const tbpdoc = DriveApp.getFileById("/*tbp temp doc ID*/");
        contractcopy = tbpdoc.makeCopy(mostRecentProposalObject.propName + " Test Boring Proposal", contractdir);
        break;
      case "Test Pit Proposal":
        Logger.log("Test Pit Proposal");
        const tppdoc = DriveApp.getFileById("/*tpp temp doc ID*/"); 
        contractcopy = tppdoc.makeCopy(mostRecentProposalObject.propName + " Test Pit Proposal", contractdir);
        break;
      /*case "Site Design Proposal":
        Logger.log("Site Design Proposal");
        const sdpdoc = DriveApp.getFileById("/*sdp temp doc ID*/"); 
        contractcopy = sdpdoc.makeCopy(mostRecentProposalObject.propName + " Site Design Proposal", contractdir);
        break;
      case "Construction Proposal":
        Logger.log("Construction Proposal");
        const cpdoc = DriveApp.getFileById("/*const prop temp doc ID*/"); 
        contractcopy = cpdoc.makeCopy(mostRecentProposalObject.propName + " Construction Proposal", contractdir);
        break;*/
      default:
        // error: needs a prop type to create form
        contractdir.setTrashed(true);
        newdir.setTrashed(true);
        throw new Error("Missing or invalid proposal type!")
    }
    
    // open contract copy, replace {{}} keywords
    contractcopy = DocumentApp.openById(contractcopy.getId());
    var body = contractcopy.getBody();
    body.replaceText("{{Date}}",Utilities.formatDate(mostRecentProposalObject.propDate, 'America/New_York', 'MMMM dd, yyyy'));
    body.replaceText("{{Proposal Name}}",mostRecentProposalObject.propName);
    body.replaceText("{{Project}}",mostRecentProposalObject.propName);
    body.replaceText("{{Client Address}}", mostRecentProposalObject.propAddr);
    body.replaceText("{{Address}}", mostRecentProposalObject.propAddr);
    body.replaceText("{{Fee or estimated fee}}", new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(mostRecentProposalObject.propEstFee));
    body.replaceText("{{Fee or Estimated Fee}}", new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(mostRecentProposalObject.propEstFee));
    body.replaceText("{{Project Information}}", mostRecentProposalObject.propInfo);
    body.replaceText("{{Scope}}", mostRecentProposalObject.propScope);
    body.replaceText("{{Proposal Number}}",mostRecentProposalObject.propNum);
    // email, phone from client contact sheet
    Logger.log("getting contact from client ID");
    clientContact = visquery("clientcontactbyid",mostRecentProposalObject.propClientContactID);
    Logger.log("contact retrieved, replacing email and phone in doc for ID: "+clientContact.contactID);
    body.replaceText("{{Email}}", clientContact.contactEmail);
    body.replaceText("{{Phone}}",clientContact.contactPhone);
    body.replaceText("{{Client Contact}}",clientContact.contactName);
    body.replaceText("{{Title}}",clientContact.contactTitle);
    // company name from clients sheet
    Logger.log("getting company name from clients sheet");
    client = visquery("clientbyid",mostRecentProposalObject.propClient)
    Logger.log("client retrieved, replacing company/client name in doc for ID: "+client.clientID);
    body.replaceText("{{Company}}",client.clientName);
    body.replaceText("{{Client}}",client.clientName);
    Logger.log("all info replaced, saving and closing");
    // save and close google doc copy, create pdf copy
    contractcopy.saveAndClose();
    Logger.log("generating PDF");
    gdocToPDF(contractcopy.getId(), contractdir.getId());
  }
  else{
    // changes detected but not in proposals sheet
    Logger.log("not a proposal, ignoring");
  }
}

// search spreadsheet for different types of data:
// - "numprops" returns the number of proposals currently on the "Proposals" sheet
// - "lastprop" returns latest/last proposal on "Proposals" sheet
// - "clientbyid" returns a client object for the given client/company id
// - "clientcontactbyid" returns a contact object for the given client contact id
// any other use will throw "I don't know what to do here"
function visquery(type, query=null)
{
  Logger.log("visquery, type "+type);
  const spreadsheetID = "/*spreadsheet ID*/"
  spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  switch(type){
    case "numprops":
      // return number of proposals 
      sheetName = "Proposals";
      query = "SELECT COUNT(A)";
      queryColumnLetterStart = "A";
      queryColumnLetterEnd = "AD";
      var qvizURL = 'https://docs.google.com/spreadsheets/d/' + spreadsheetID + '/gviz/tq?tqx=out:json&headers=1&sheet=' + sheetName + '&range=' + queryColumnLetterStart + ":" + queryColumnLetterEnd + '&tq=' + encodeURIComponent(query);
      var ret = UrlFetchApp.fetch(qvizURL, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}}).getContentText();
      //Logger.log(ret);
      ret = ret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2);
      ret = JSON.parse(ret);
      Logger.log("returning number of proposals in sheet: "+ret.table.rows[0].c[0].v);
      return ret.table.rows[0].c[0].v;
    case "lastprop":
      // return last proposal
      sheet = spreadsheet.getSheetByName("Proposals");
      lastrow = visquery("numprops")+1
      const lastRange = sheet.getRange(lastrow, 1, lastrow, 30);
      const vals = lastRange.getValues()[0]
      mostRecentProposalObject = {
        propNum: vals[0],
        propName: vals[1],
        propCode: vals[2],
        propDate: vals[3],               
        propProjManager: vals[4],
        propEstFee: vals[5],
        propProb: vals[6],
        propInfo: vals[7],
        propScope: vals[8],
        propAddr: vals[9],
        propSubconsultant: vals[10],
        propSubconsultantFee: vals[11],
        propClient: vals[12],
        propClientContactID: vals[13],
        propStatus: vals[14],
        propProjNum: vals[15],
        propIsSubconsultant: vals[16],
        propExpProfit: vals[17],
        propIsOtherExpenses: vals[18],
        propCost: vals[19],
        propIsDead: vals[20],
        propIsBorings: vals[21],
        propTags: vals[22],
        propIsPotentialConstruction: vals[23],
        propType: vals[24], 
        propDocLink: vals[25],
        propNotes: vals[26],
        propID: vals[27],
        propProjID: vals[28],
        propType: vals[29]
      }
      Logger.log("returning most recent project in proposals sheet: "+mostRecentProposalObject.propName);
      return mostRecentProposalObject;
    case "clientbyid":
      // return client object from given id
      sheetName = "Clients";
      id = query;
      queryColumnLetterStart = "A";
      queryColumnLetterEnd = "E";
      queryColumnLetterSearch= "B"
      query = "SELECT * WHERE "+queryColumnLetterSearch+ " = '" + id + "'";
      var qvizURL = 'https://docs.google.com/spreadsheets/d/' + spreadsheetID + '/gviz/tq?tqx=out:json&headers=1&sheet=' + sheetName + '&range=' + queryColumnLetterStart + ":" + queryColumnLetterEnd + '&tq=' + encodeURIComponent(query);
      var ret = UrlFetchApp.fetch(qvizURL, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}}).getContentText();
      ret = ret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2);
      ret = JSON.parse(ret);
      var cvals = [];
      for(let i=0; i<5; i++){
        if(ret.table.rows[0].c[i] == null){
          cvals[i] = "";
          //Logger.log(i+" is null");
        }
        else{
          cvals[i] = ret.table.rows[0].c[i].v;
          //Logger.log(i+" is not null");
        }
      }
      // create client object to return
      client = {
          clientName: cvals[0],
          clientID: cvals[1],
          clientAddress: cvals[2],
          clientBilling: cvals[3],
          clientType: cvals[4]
        }
        Logger.log("returning client: "+client.clientName+" from given id: "+id);
      return client;
    case "clientcontactbyid":
      // return client contact object from given id
      sheetName = "Client Contacts";
      id = query;
      queryColumnLetterStart = "A";
      queryColumnLetterEnd = "F";
      queryColumnLetterSearch= "A"
      query = "SELECT * WHERE "+queryColumnLetterSearch+ " = '" + id + "'";
      var qvizURL = 'https://docs.google.com/spreadsheets/d/' + spreadsheetID + '/gviz/tq?tqx=out:json&headers=1&sheet=' + sheetName + '&range=' + queryColumnLetterStart + ":" + queryColumnLetterEnd + '&tq=' + encodeURIComponent(query);
      var ret = UrlFetchApp.fetch(qvizURL, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}}).getContentText();
      ret = ret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2);
      ret = JSON.parse(ret);
      var cvals = [];
      for(let i=0; i<6; i++){
        if(ret.table.rows[0].c[i] == null){
          cvals[i] = "";
          //Logger.log(i+" is null");
        }
        else{
          cvals[i] = ret.table.rows[0].c[i].v;
          //Logger.log(i+" is not null");
        }
      }
      // create contact object to return
      contact = {
          contactID: cvals[0],
          contactName: cvals[1],
          contactCompany: cvals[2],
          contactTitle: cvals[3],
          contactPhone: cvals[4],
          contactEmail: cvals[5]
        }
      Logger.log("returning client contact: "+contact.contactName+" from given id: "+id);
      return contact;
    default:
      // default case, should never be called
      throw new Error("Invalid search type, don't know what to do here");
  }
}

// create pdf helper function
function createPDF(fileID, folderID, callback) {
    var templateFile = DriveApp.getFileById(fileID);
    var templateName = templateFile.getName();
    
    var existingPDFs = DriveApp.getFolderById(folderID).getFiles();

    //in case no files exist
    if (!existingPDFs.hasNext()) {
        return callback(fileID, folderID);
    }

    for (; existingPDFs.hasNext();) {

        var existingPDFfile = existingPDFs.next();
        var existingPDFfileName = existingPDFfile.getName();
        if (existingPDFfileName == templateName + ".pdf") {
            Logger.log("PDF exists already. No PDF created")
            return callback();
        }
        if (!existingPDFs.hasNext()) {
            Logger.log("PDF is created")
            return callback(fileID, folderID)
        }
    }
}

// other create pdf helper function
function createPDFfile(fileID, folderID) {
    var templateFile = DriveApp.getFileById(fileID);
    var folder = DriveApp.getFolderById(folderID);
    var theBlob = templateFile.getBlob().getAs('application/pdf');
    var newPDFFile = folder.createFile(theBlob);

    var fileName = templateFile.getName().replace(".", ""); 
    newPDFFile.setName(fileName + ".pdf");
}

// primary gdoc -> pdf controller function
function gdocToPDF(fileID, folder) { 
  var pdfFolder = DriveApp.getFolderById(folder); 
  
  var docFile = DriveApp.getFileById(fileID)
  
  createPDF(docFile.getId(), pdfFolder.getId(), function (fileID, folderID) {
    if (fileID) createPDFfile(fileID, folderID);
  }
 )
}

// checks proposals folder to see if proposal already exists
function propexists(propcode) {
  // list subfolders in props folder
  var tpropsdir = DriveApp.getFolderById("/*proposals dir ID*/");
  var folders = tpropsdir.getFolders();
  while(folders.hasNext()){
    folder = folders.next();
    if(folder.getName() == propcode){
      return true;
    }
  }
  return false;
}

