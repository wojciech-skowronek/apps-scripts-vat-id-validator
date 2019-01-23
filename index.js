function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{name: "Validate", functionName: "validate"}]
    ss.addMenu("VAT ID Validator", menuEntries);
}

function validate() {
    // Take current selection
    var range = SpreadsheetApp.getActiveSpreadsheet().getSelection().getActiveRange();

    // Iterate through range
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
    for (var i = 1; i <= numRows; i++) {
        for (var j = 1; j <= numCols; j++) {
            var cell = range.getCell(i, j);
            var currentValue = cell.getValue();
            // check if the cell value has VAT ID format
            var parsed = currentValue.match(/([A-Z]{2})(\w+)/);
            // if it's not VAT ID, skip it
            if(!parsed){
                continue
            }

            // make call to VIES
            var validated = makeCall(parsed[1], parsed[2])

            // Add colors
            if(validated.isValid){
                // as addition we can add a note with company name and address retrieved from VIES
                cell.setNote(validated.companyName + "\n\n" + validated.companyAddress)
                cell.setBackgroundColor("#dfffdb");
            } else {
                cell.setBackgroundColor("#e6b8af");
            }
        }
    }
}

function makeCall(countryCode, vatNumber) {

    // Create SOAP message for WDSL: http://ec.europa.eu/taxation_customs/vies/checkVatService.wsdl
    var message = '<?xml version="1.0" encoding="UTF-8"?>' +
        '<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns1="urn:ec.europa.eu:taxud:vies:services:checkVat:types">' +
        '  <SOAP-ENV:Body>' +
        '    <ns1:checkVat>' +
        '      <ns1:countryCode>' + countryCode + '</ns1:countryCode>' +
        '      <ns1:vatNumber>' + vatNumber + '</ns1:vatNumber>' +
        '    </ns1:checkVat>' +
        '  </SOAP-ENV:Body>' +
        '</SOAP-ENV:Envelope>'

    // Use UrlFetchApp (https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app) to send POST request
    var xml = UrlFetchApp.fetch("http://ec.europa.eu/taxation_customs/vies/services/checkVatService", {
        method: "POST",
        contentType: 'text/xml',
        payload: message
    }).getContentText()

    // the response is XML, which can be parsed with XmlService (https://developers.google.com/apps-script/reference/xml-service/)
    var document = XmlService.parse(xml);
    var mainNs = XmlService.getNamespace('http://schemas.xmlsoap.org/soap/envelope/');
    var checkVatResponseNs = XmlService.getNamespace('urn:ec.europa.eu:taxud:vies:services:checkVat:types');

    var root = document.getRootElement().getChild("Body", mainNs).getChild("checkVatResponse", checkVatResponseNs);

    // Extract interesting information
    var isValid = root.getChild("valid", checkVatResponseNs).getText()
    var companyName = root.getChild("name", checkVatResponseNs).getText()
    var companyAddress = root.getChild("address", checkVatResponseNs).getText()

    return {
        isValid: "true" === isValid,
        companyName: companyName,
        companyAddress: companyAddress
    }
}


