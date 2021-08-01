const messageSheets = SpreadsheetApp.openById('12Lduw8zGSu45YByeQy0fTLx2PKzaJ6NyOoZFkYtlbqA')
const senderEndpoint: string = PropertiesService.getScriptProperties().getProperty('sender_endpoint')

const createEverydayTrigger = () => {
    ScriptApp.newTrigger('prepareMessages').timeBased().nearMinute(1).everyDays(1).atHour(0).create()
}

const prepareMessages = () => {
    const holidaysSheet = SpreadsheetApp.openById('1dD14Up7ialGPRQmTb74zgxbctB9CaUsYa41aVIHxfJ4').getSheetByName('public holiday')
    const holidays = holidaysSheet.getRange(1, 1, holidaysSheet.getLastRow()).getValues().map(row => row[0])
    const channelMaster: string[][] = getAllValues(messageSheets, 'master_channel')
    const botMaster: string[][] = getAllValues(messageSheets, 'master_bot')
    const userMaster: string[][] = getAllValues(messageSheets, 'master_user')
    const groupMaster: string[][] = getAllValues(messageSheets, 'master_group')
    const preparedSheet = messageSheets.getSheetByName('prepared')
    const preparingMessages: any[][] = []

    const createMessage = (channel, bot, user, group, message): { [key: string]: string } => {
        const channelId = channelMaster.find(channelRow => channelRow[0] == channel)?.[1]
        const botName = botMaster.find(botRow => botRow[0] == bot)?.[0]
        const userId = userMaster.find(userRow => userRow[0] == user)?.[1]
        const groupId = groupMaster.find(groupRow => groupRow[0] == group)?.[1]
        const text = `${userId ? `<@${userId}> ` : ''}${groupId ? `<!subteam^${groupId}> ` : ''}${message}`
        return {
            channelId: channelId,
            botName: botName,
            text: text,
        }
    }

    const createSenderTrigger = (date: Date): string => {
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            headers: {
                Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
            },
            payload: {
                timestamp: date.getTime()
            },
            muteHttpExceptions: true
        }
        Logger.log(options)
        const contentText = UrlFetchApp.fetch(senderEndpoint, options).getContentText()
        try {
            const response = JSON.parse(contentText)
            return response.triggerUid
        } catch {
            Logger.log(contentText)
            return null
        }
    }

    const prepareRegularMessages = () => {
        const isValidRow = (day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, bot?, user?, group?, message?): boolean => {
            if (day == '' || day == 'なし') {
                Logger.log('day: empty or なし')
                return false
            }
            if (typeOf(times) == 'number' && times < 1) {
                Logger.log('times: 0')
                return false
            }
            return isValidMessage(hour, minute, channel, bot, message)
        }
        const isHoliday = (): boolean => {
            const today = new Date()
            const isHoliday: boolean = holidays.some(holiday => holiday.getFullYear() == today.getFullYear() && holiday.getMonth() == today.getMonth() && holiday.getDate() == today.getDate())
            if (isHoliday) {
                Logger.log('is holiday')
            }
            return isHoliday
        }
        const isTheDayOfTheWeek = (day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, bot?, user?, group?, message?): boolean => {
            const dayOfTheWeek = new Date().getDay()
            const days = ['日', '月', '火', '水', '木', '金', '土']
            if (day == days[dayOfTheWeek]) return true
            if (day == '平日' && dayOfTheWeek != 0 && dayOfTheWeek != 6) return true
            Logger.log('today is not the day of the week')
            return false
        }
        const isSkip = (day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, bot?, user?, group?, message?): boolean => {
            if (typeOf(skip) == 'number' && skip > 0) {
                Logger.log('skip')
                return true
            }
            return false
        }
        const prepare = (day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, bot?, user?, group?, message?): any[] => {
            const { channelId, botName, text } = createMessage(channel, bot, user, group, message)
            if (!channelId) {
                Logger.log('found no channel id')
                return null
            }
            if (!botName) {
                Logger.log('found no bot')
                return null
            }
            if (preparedMessages.some(messageRow => messageRow[1] == channelId && messageRow[2] == botName && messageRow[3] == text)) {
                Logger.log('already prepared')
                return null
            }
            const date = new Date()
            date.setHours(hour)
            date.setMinutes(minute)
            const triggerUid = createSenderTrigger(date)
            if (!triggerUid) {
                Logger.log('create sender trigger failed')
                return null
            }
            Logger.log('regular message prepared')
            return [triggerUid, channelId, botName, text, date]
        }
        const updateIfNeeded = (index: number, day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, bot?, user?, group?, message?) => {
            const newTimes = times != '' && times > 0 ? times - 1 : times
            const newSkip = typeOf(skipPeriod) == 'number' && skipPeriod > 0 ? skipPeriod : skip
            if (newTimes != times || newSkip != skip) {
                const newRow = [day, hour, minute, newTimes, newSkip, skipPeriod, channel, bot, user, group, message]
                regularSheet.getRange(index + 2, 1, 1, newRow.length).setValues([newRow])
            }
        }
        const updateSkip = (index: number, day?, hour?, minute?, times?, skip?, skipPeriod?, channel?, bot?, user?, group?, message?) => {
            const newSkip = skip - 1
            const newRow = [day, hour, minute, times, newSkip, skipPeriod, channel, bot, user, group, message]
            regularSheet.getRange(index + 2, 1, 1, newRow.length).setValues([newRow])
        }

        const regularSheet = messageSheets.getSheetByName('regular')
        const lastRowNumber = regularSheet.getLastRow()
        if (lastRowNumber <= 1) {
            Logger.log('no regular message')
            return
        }
        const regularMessages = regularSheet.getRange(2, 1, lastRowNumber - 1, regularSheet.getLastColumn()).getValues()
        const lastPreparedRowNumber = preparedSheet.getLastRow()
        const preparedMessages: any[][] = lastPreparedRowNumber > 0 ? preparedSheet.getRange(1, 1, lastPreparedRowNumber, preparedSheet.getLastColumn()).getValues() : []
        regularMessages.forEach((row: any[], index: number) => {
            if (isValidRow(...row) && isTheDayOfTheWeek(...row) && !isHoliday()) {
                if (!isSkip(...row)) {
                    const preparingMessage: any[] = prepare(...row)
                    if (preparingMessage != null) {
                        preparingMessages.push(preparingMessage)
                        updateIfNeeded(index, ...row)
                    }
                } else {
                    updateSkip(index, ...row)
                }
            }
        })
    }

    const prepareSingleMessage = () => {
        const isToday = (date?, hour?, minute?, channel?, bot?, user?, group?, message?): boolean => {
            const today = new Date()
            const isToday = typeOf(date) == 'date' && date.getFullYear() == today.getFullYear() && date.getMonth() == today.getMonth() && date.getDate() == today.getDate()
            if (!isToday) {
                Logger.log('is not today')
            }
            return isToday
        }
        const isValidRow = (date?, hour?, minute?, channel?, bot?, user?, group?, message?): boolean => {
            return isValidMessage(hour, minute, channel, bot, message)
        }
        const prepare = (date?, hour?, minute?, channel?, bot?, user?, group?, message?): any[] => {
            const { channelId, botName, text } = createMessage(channel, bot, user, group, message)
            if (!channelId) {
                Logger.log('found no channel id')
                return null
            }
            if (!botName) {
                Logger.log('found no bot')
                return null
            }
            date.setHours(hour)
            date.setMinutes(minute)
            const triggerUid = createSenderTrigger(date)
            if (!triggerUid) {
                Logger.log('create sender trigger failed')
                return null
            }
            Logger.log('single message prepared')
            return [triggerUid, channelId, botName, text, date]
        }
        const deleteRow = (index: number) => {
            singleSheet.deleteRow(index + 2)
            singleSheet.insertRowAfter(Math.max(singleSheet.getLastRow(), 1))
        }

        const singleSheet = messageSheets.getSheetByName('single')
        const lastRowNumber = singleSheet.getLastRow()
        if (lastRowNumber <= 1) {
            Logger.log('no single message')
            return
        }
        const singleMessages = singleSheet.getRange(2, 1, lastRowNumber - 1, singleSheet.getLastColumn()).getValues()
        const preparedIndices = []
        singleMessages.forEach((row: any[], index: number) => {
            if (isToday(...row) && isValidRow(...row)) {
                const preparingMessage: any[] = prepare(...row)
                if (preparingMessage != null) {
                    preparingMessages.push(preparingMessage)
                    preparedIndices.push(index)
                }
            }
        })
        preparedIndices.reverse().forEach(index => deleteRow(index))
    }

    try {
        Logger.log('start prepare messages')
        prepareRegularMessages()
        prepareSingleMessage()
    } finally {
        if (preparingMessages.length == 0) return
        Logger.log('write prepare messages')
        preparedSheet.getRange(preparedSheet.getLastRow() + 1, 1, preparingMessages.length, preparingMessages[0].length).setValues(preparingMessages)
    }
}

const getAllValues = (sheets, sheetName: string): string[][] => {
    const sheet = sheets.getSheetByName(sheetName)
    return sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues()
}

const isValidMessage = (hour, minute, channel, bot, message): boolean => {
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
    if (bot == '') {
        Logger.log('empty bot')
        return false
    }
    if (message == '') {
        Logger.log('empty message')
        return false
    }
    return true
}

const typeOf = (obj: any): string => Object.prototype.toString.call(obj).slice(8, -1).toLowerCase()