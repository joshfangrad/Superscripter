

function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Format', 'format')
        .addItem('Options', 'showSidebar')
        .addToUi();

    //we have this for a while so curent users get the new preferences
    var prefs = getPrefs();
    if (!prefs['letter']) {
        savePref('letter', 'false');
    } 
    if (!prefs['plus']) {
        savePref('plus', 'false');
    }
    if (!prefs['minus']) {
        savePref('minus', 'false');
    }
    if (!prefs['slash']) {
        savePref('slash', 'false');
    }
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
    var defaultProps = {super: '^', sub: '>', letter: 'false', plus: 'false', minus: 'false', slash: 'false'};
    userProps.setProperties(defaultProps, true);
}

function savePref(id, val) {
    var userProps = PropertiesService.getUserProperties();
    userProps.setProperty(id, val);
}

function getPrefs() {
    var userProps = PropertiesService.getUserProperties();
    var charPrefs = {
        super: userProps.getProperty('super'),
        sub: userProps.getProperty('sub'),
        letter: userProps.getProperty('letter'),
        plus: userProps.getProperty('plus'),
        minus: userProps.getProperty('minus'),
        slash: userProps.getProperty('slash')
    };
    return charPrefs;
}

function format() {
    //function to change numbers that are lead by certain symbols to super/subscript.
    var prefs = getPrefs();
    //there are some letters we can't escape with a backslash otherwise they would do other regex things
    var letterBlacklist = new RegExp('^[wsdcbn]','i');
    var blackFind = [prefs['super'].match(letterBlacklist), prefs['sub'].match(letterBlacklist)];
    //construct our regexes out of strings
    var superString = '(';
    var subString = '(';
    //we always escape the first regex character unless it's a blacklisted char
    if (!blackFind[0]) {
        superString += '\\';
    }
    if (!blackFind[1]) {
        subString += '\\';
    }
    superString += prefs['super'] + ')[';
    subString += prefs['sub'] + ')[';
    //add extra regex options based on user settings
    if (prefs['letter'] == 'true') {
        superString += 'a-z';
        subString += 'a-z';
    }
    if (prefs['minus'] == 'true') {
        superString += '-';
        subString += '-';
    }
    if (prefs['plus'] == 'true') {
        superString += '+';
        subString += '+';
    }
    if (prefs['slash'] == 'true') {
        superString += '/';
        subString += '/';
    }
    superString += '0-9.]+';
    subString += '0-9.]+';
    //end regex will look something like:
    //   /(\^)[a-z-+0-9.]+/g
    superRegex = new RegExp(superString, 'g');
    subRegex = new RegExp(subString, 'g');
    var checks = [superRegex, subRegex];

    var doc = DocumentApp.getActiveDocument();
    var paras = doc.getParagraphs();
    //for each of our regexes run the checks
    for (var j = 0; j < checks.length; j++) {
        for (var i = 0; i < paras.length; i++) {
            while (match = checks[j].exec(paras[i].getText())) {
                //plusses need to be escaped
                match[0] = match[0].replace(/\+/g, '\\+');
                //we also escape the first character to avoid regex problems
                var find = paras[i].findText('\\' + match[0]);
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