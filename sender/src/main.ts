const messageSheets = SpreadsheetApp.openById('12Lduw8zGSu45YByeQy0fTLx2PKzaJ6NyOoZFkYtlbqA')

const doPost = (e: RequestParameter) => {
    Logger.log(e.parameter)
    const timestamp = parseFloat(e.parameter.timestamp)
    if (isNaN(timestamp) || new Date().getTime() > timestamp) {
        ContentService.createTextOutput('')
    }
    const date = new Date(timestamp)
    Logger.log(date)
    const trigger: GoogleAppsScript.Script.Trigger = ScriptApp.newTrigger('sendMessage').timeBased().at(date).create()
    return ContentService.createTextOutput(JSON.stringify({ triggerUid: trigger.getUniqueId() }))
}

const sendMessage = (e: object) => {
    Logger.log(e)
    if (e == null) {
        Logger.log('no args')
        return
    }
    const triggerUid: string = e['triggerUid']
    const sheet = messageSheets.getSheetByName('prepared')
    const lastRowNumber = sheet.getLastRow()
    if (lastRowNumber == 0) return
    const messages: string[][] = sheet.getRange(1, 1, lastRowNumber, sheet.getLastColumn()).getDisplayValues()
    const message: string[] = messages.find(message => message[0] == triggerUid)
    const channelId: string = message[1]
    const botName: string = message[2]
    const text: string = message[3]
    if (channelId.length == 0 || text.length == 0) return
    const token = PropertiesService.getScriptProperties().getProperty(`slack_${botName}_token`)
    const formData = {
        token: token,
        channel: channelId,
        text: text
    }
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: 'post',
        payload: formData,
        muteHttpExceptions: true
    }
    Logger.log(UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', options).getContentText())

    const index = messages.findIndex(message => message[0] == triggerUid)
    if (index >= 0) {
        sheet.deleteRow(index + 1)
        sheet.insertRowAfter(Math.max(sheet.getLastRow(), 1))
    }
    ScriptApp.getProjectTriggers().forEach(trigger => {
        if (trigger.getUniqueId() == triggerUid) {
            ScriptApp.deleteTrigger(trigger)
        }
    });
}

interface RequestParameter {
    queryString: string,
    parameter: { [key: string]: string },
    parameters: { [key: string]: string[] },
    contextPath: string,
    contextLength: number,
    postData: PostData,

}

interface PostData {
    length: number,
    type: string,
    contents: any,
    name: string // Always the value "postData"
}