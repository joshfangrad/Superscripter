

function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Format', 'format')
        .addItem('Options', 'showSidebar')
        .addToUi();
}

function onInstall(e) {
    //if it's their first time using the addon, set the default prefs
    newPrefs();
    onOpen(e);
}

function showSidebar() {
    var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
        .setTitle('Superscripter Options');
    DocumentApp.getUi().showSidebar(ui);
}

function newPrefs() {
    var userProps = PropertiesService.getUserProperties();
    userProps.setProperty('super', '^');
    userProps.setProperty('sub', '>');
}

function savePref(id, val) {
    var userProps = PropertiesService.getUserProperties();
    userProps.setProperty(id, val);
}

function getPrefs() {
    var userProps = PropertiesService.getUserProperties();
    var charPrefs = {
        super: userProps.getProperty('super'),
        sub: userProps.getProperty('sub')
    };
    return charPrefs;
}

function format() {
    //function to change numbers that are lead by certain symbols to super/subscript.
    var prefs = getPrefs();
    //there are some letters we can't escape with a backslash otherwise they would do other regex things
    var letterBlacklist = new RegExp('^[wsdcbn]','i');
    var blackFind = [prefs['super'].match(letterBlacklist), prefs['sub'].match(letterBlacklist)];
    //construct our regexes out of our characters
    var superRegex;
    var subRegex;
    if (blackFind[0]) {
        superRegex = new RegExp('(' + prefs['super'] + ')[0-9.]+', 'g');
    } else {
        superRegex = new RegExp('(\\' + prefs['super'] + ')[0-9.]+', 'g');
    }
    if (blackFind[1]) {
        subRegex = new RegExp('(' + prefs['sub'] + ')[0-9.]+', 'g');
    } else {
        subRegex = new RegExp('(\\' + prefs['sub'] + ')[0-9.]+', 'g');
    }
    Logger.log(superRegex);
    var checks = [superRegex, subRegex];

    var doc = DocumentApp.getActiveDocument();
    var paras = doc.getParagraphs();
    //for each of our regexes run the checks
    for (var j = 0; j < checks.length; j++) {
        for (var i = 0; i < paras.length; i++) {
            while (match = checks[j].exec(paras[i].getText())) {
                var find;
                if (blackFind[j]) {
                    find = paras[i].findText(match[0]);
                } else {
                    find = paras[i].findText('\\' + match[0]);
                }
                if (find) {
                    var elementText = find.getElement().asText();
                    if (j == 0) {
                        elementText.setTextAlignment(find.getStartOffset(), find.getEndOffsetInclusive(), DocumentApp.TextAlignment.SUPERSCRIPT);
                    } else {
                        elementText.setTextAlignment(find.getStartOffset(), find.getEndOffsetInclusive(), DocumentApp.TextAlignment.SUBSCRIPT);
                    }
                    elementText.deleteText(find.getStartOffset(), find.getStartOffset());
                }
            }
        }
    }
}