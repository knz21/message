const messageSheets = SpreadsheetApp.openById('12Lduw8zGSu45YByeQy0fTLx2PKzaJ6NyOoZFkYtlbqA')

const prepareMessages = () => {
    const holidaysSheet = SpreadsheetApp.openById('1dD14Up7ialGPRQmTb74zgxbctB9CaUsYa41aVIHxfJ4').getSheetByName('public holiday')
    const holidays = holidaysSheet.getRange(1, 1, holidaysSheet.getLastRow()).getValues().map(row => row[0])
    const channelMaster: string[][] = getAllValues(messageSheets, 'master_channel')
    const userMaster: string[][] = getAllValues(messageSheets, 'master_user')
    const groupMaster: string[][] = getAllValues(messageSheets, 'master_group')
    const preparedSheet = messageSheets.getSheetByName('prepared')
    const preparingMessages: any[][] = []

    const createMessage = (channel, user, group, message): { [key: string]: string } => {
        const channelId = channelMaster.find(channelRow => channelRow[0] == channel)?.[1]
        const userId = userMaster.find(userRow => userRow[0] == user)?.[1]
        const groupId = userMaster.find(groupRow => groupRow[0] == group)?.[1]
        const text = `${userId ? `<@${userId}> ` : ''}${groupId ? `<!subteam^${groupId}> ` : ''}${message}`
        return {
            channelId: channelId,
            text: text,
        }
    }

    const prepareRegularMessages = () => {
        const isValidRow = (day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, user?, group?, message?): boolean => {
            if (day == '' || day == 'なし') {
                Logger.log('day: empty or なし')
                return false
            }
            if (times != '' && times == 0) {
                Logger.log('times: 0')
                return false
            }
            return isValidMessage(hour, minute, channel, message)
        }
        const isHoliday = (): boolean => {
            const holidays = holidaysSheet.getRange(1, 1, holidaysSheet.getLastRow()).getValues().map(row => row[0])
            const today = new Date()
            return holidays.some(holiday => holiday.getFullYear() == today.getFullYear() && holiday.getMonth() == today.getMonth() && holiday.getDate() == today.getDate())
        }
        const isTheDayOfTheWeek = (day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, user?, group?, message?): boolean => {
            const dayOfTheWeek = new Date().getDay()
            const days = ['日', '月', '火', '水', '木', '金', '土']
            if (day == days[dayOfTheWeek]) return true
            if (day == '平日' && dayOfTheWeek != 0 && dayOfTheWeek != 6) return true
            Logger.log('today is not the day of the week')
            return false
        }
        const isSkip = (day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, user?, group?, message?): boolean => {
            if (skip != '' && skip > 0) {
                Logger.log('skip')
                return true
            }
            return false
        }
        const prepare = (day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, user?, group?, message?): any[] => {
            const { channelId, text } = createMessage(channel, user, group, message)
            if (!channelId) {
                Logger.log('found no channel id')
                return null
            }
            if (preparedMessages.some(messageRow => messageRow[1] == channelId && messageRow[2] == text)) {
                Logger.log('already prepared')
                return null
            }
            const date = new Date()
            date.setHours(hour)
            date.setMinutes(minute)
            const trigger = ScriptApp.newTrigger('sendMessage').timeBased().at(date).create()
            Logger.log('regular message prepared')
            return [trigger.getUniqueId(), channelId, text, date]
        }

        const updateTimes = (index: number, day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, user?, group?, message?) => {
            if (times != '' && times > 0) {
                const newRow = [day, hour, minute, times - 1, skip, skipPeriod, channel, user, group, message]
                regularSheet.getRange(index + 2, 1, 1, newRow.length).setValues([newRow])
            }
        }

        const updateSkips = (index: number, day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, user?, group?, message?) => {
            const newSkip = skip - 1
            const newRow = [day, hour, minute, times, newSkip == 0 ? skipPeriod : newSkip, skipPeriod, channel, user, group, message]
            regularSheet.getRange(index + 2, 1, 1, newRow.length).setValues([newRow])
        }

        const regularSheet = messageSheets.getSheetByName('regular')
        const regularMessages = regularSheet.getRange(2, 1, regularSheet.getLastRow() - 1, regularSheet.getLastColumn()).getValues()
        const lastPreparedRowNumber = preparedSheet.getLastRow()
        const preparedMessages: any[][] = lastPreparedRowNumber > 0 ? preparedSheet.getRange(1, 1, lastPreparedRowNumber, preparedSheet.getLastColumn()).getValues() : []
        regularMessages.forEach((row: any[], index: number) => {
            if (isValidRow(...row) && isTheDayOfTheWeek(...row) && !isHoliday()) {
                if (!isSkip(...row)) {
                    const preparingMessage: any[] = prepare(...row)
                    if (preparingMessage != null) {
                        preparingMessages.push(preparingMessage)
                        updateTimes(index, ...row)
                    }
                } else {
                    updateSkips(index, ...row)
                }
            }
        })
    }

    const prepareSingleMessage = () => {
        const isToday = (date?, hour?, minute?, channel?, user?, group?, message?): boolean => {
            const today = new Date()
            const isToday = typeOf(date) == 'date' && date.getFullYear() == today.getFullYear() && date.getMonth() == today.getMonth() && date.getDate() == today.getDate()
            if (!isToday) {
                Logger.log('is not today')
            }
            return isToday
        }
        const isValidRow = (date?, hour?, minute?, channel?, user?, group?, message?): boolean => {
            return isValidMessage(hour, minute, channel, message)
        }
        const prepare = (date?, hour?, minute?, channel?, user?, group?, message?): any[] => {
            const { channelId, text } = createMessage(channel, user, group, message)
            if (!channelId) {
                Logger.log('found no channel id')
                return null
            }
            date.setHours(hour)
            date.setMinutes(minute)
            const trigger = ScriptApp.newTrigger('sendMessage').timeBased().at(date).create()
            Logger.log('single message prepared')
            return [trigger.getUniqueId(), channelId, text, date]
        }
        const deleteRow = (index: number) => {
            singleSheet.deleteRow(index + 2)
            singleSheet.insertRowAfter(Math.max(singleSheet.getLastRow(), 1))
        }

        const singleSheet = messageSheets.getSheetByName('single')
        const singleMessages = singleSheet.getRange(2, 1, singleSheet.getLastRow() - 1, singleSheet.getLastColumn()).getValues()
        singleMessages.forEach((row: any[], index: number) => {
            if (isToday(...row) && isValidRow(...row)) {
                const preparingMessage: any[] = prepare(...row)
                if (preparingMessages != null) {
                    preparingMessages.push(preparingMessage)
                    deleteRow(index)
                }
            }
        })
    }

    try {
        prepareRegularMessages()
        prepareSingleMessage()
    } finally {
        if (preparingMessages.length == 0) return
        preparedSheet.getRange(preparedSheet.getLastRow() + 1, 1, preparingMessages.length, preparingMessages[0].length).setValues(preparingMessages)
    }
}

const getAllValues = (sheets, sheetName: string): string[][] => {
    const sheet = sheets.getSheetByName(sheetName)
    return sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues()
}

const isValidMessage = (hour, minute, channel, message): boolean => {
    if (typeOf(hour) != 'number' || hour < 0 || hour > 23) {
        Logger.log('hour: invalid')
        return false
    }
    if (typeOf(minute) != 'number' || minute < 0 || minute > 59) {
        Logger.log('minute: invalid')
        return false
    }
    if (channel == '') {
        Logger.log('empty channel')
        return false
    }
    if (message == '') {
        Logger.log('empty message')
        return false
    }

    return true
}

const sendMessage = (e: object) => {
    const triggerUid: string = e['triggerUid']
    const sheet = messageSheets.getSheetByName('prepared')
    const lastRowNumber = sheet.getLastRow()
    if (lastRowNumber == 0) return
    const messages: string[][] = sheet.getRange(1, 1, lastRowNumber, sheet.getLastColumn()).getDisplayValues()
    const message: string[] = messages.find(message => message[0] == triggerUid)
    const channelId: string = message[1]
    const text: string = message[2]
    if (channelId.length == 0 || text.length == 0) return
    const token = PropertiesService.getScriptProperties().getProperty('slack_user_token')
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

const setToken = () => {
    PropertiesService.getScriptProperties().setProperty('', '')
}

const typeOf = (obj: any): string => Object.prototype.toString.call(obj).slice(8, -1).toLowerCase()