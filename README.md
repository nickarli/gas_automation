# Automation

Taking advantage of Google Apps Script to automate a few workflows for a local civil engineering firm.
* [propgen](propgen.gs): Proposal generator. Assume a Google spreadsheet with sheets for Proposals, Clients, Client Contacts as well as template documents for each of the contracts listed (Authorization to Proceed, etc). This script will detect a new proposal entry on the "Proposals" sheet, gather information about the proposal itself, the client, and the contact, and will automagically fill in the proper contract template, copy it to a PDF, and place these files into a newly-created folder for this particular proposal.
  * [Google Query Language](https://developers.google.com/chart/interactive/docs/querylanguage) is pretty sick, cut runtime from ~6 minutes for linear searching three sheets down to ~12 seconds 
  
