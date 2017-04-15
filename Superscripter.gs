var userProps = PropertiesService.getUserProperties();

function onOpen() {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Format', 'format')
        .addItem('Options', 'showSidebar')
        .addToUi();
}

function onInstall() {
    //if it's their first time using the addon, set the default prefs
    newPrefs();
    onOpen();
}

function showSidebar() {
    var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
        .setTitle('Superscripter Options');
    DocumentApp.getUi().showSidebar(ui);
}

function newPrefs() {
    userProps.setProperty('super', '^');
    userProps.setProperty('sub', '>');
}

function savePref(id, val) {
    userProps.setProperty(id, val);
}

function getPrefs() {
    var charPrefs = {
        super: userProps.getProperty('super'),
        sub: userProps.getProperty('sub')
    };
    return charPrefs;
}

function format() {
    //function to change numbers that are lead by certain symbols to super/subscript.
    var prefs = getPrefs();
    //construct our regexes out of our characters
    var superRegex = new RegExp('[\\' + prefs['super'] + '][0-9]+', 'ig');
    var subRegex = new RegExp('[\\' + prefs['sub'] + '][0-9]+', 'ig');
    var checks = [superRegex, subRegex];

    var doc = DocumentApp.getActiveDocument();
    var paras = doc.getParagraphs();
    //for each of our regexes run the checks
    for (var j = 0; j < checks.length; j++) {
        for (var i = 0; i < paras.length; i++) {
            while (match = checks[j].exec(paras[i].getText())) {
                var find = paras[i].findText('\\' + match[0]);
                if (find) {
                    var elementText = find.getElement().asText();
                    if (j == 0) {
                        elementText.setTextAlignment(find.getStartOffset(), find.getEndOffsetInclusive(), DocumentApp.TextAlignment.SUPERSCRIPT);
                    } else {
                        elementText.setTextAlignment(find.getStartOffset(), find.getEndOffsetInclusive(), DocumentApp.TextAlignment.SUBSCRIPT);
                    }
                    //remove the super/subscript symbol
                    elementText.deleteText(find.getStartOffset(), find.getStartOffset());
                }
            }
        }
    }
}