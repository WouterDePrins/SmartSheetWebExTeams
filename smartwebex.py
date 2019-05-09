# Integration between SmartSheet and WebEx teams
# This script can run as a FaaS at GCP and serves as a webhook for both WebEx Teams and SmartSheet.
# When a row has been added to SmartSheet, it will push a message to this webhook and from here, we can parse the message and forward it to a WebEx Teams Room
# On the other hand, we can ask the WebEx Teams Bot to list all the rows in the sheet.

# Created by Wouter De Prins - System Engineer Data Center - Cisco Belgium

import requests
import json

## variables - fields need to be populated

WebExBotName = ''
WebExBearer = ''
WebExRoom = ''

SmartSheetBearer = ''
SmartSheetSheet = ''

## URLS

WebExUrl = 'https://api.ciscospark.com/v1/messages/'
SmartSheetUrl = 'https://api.smartsheet.com/2.0/sheets/'

# Function to return header
def getHeader(bearer)
    return { 'Authorization': "Bearer " + bearer } 

# Funcion to get values from a row in SmartSheet
def getSmartText(id):
    values = {}
    url = SmartSheetUrl + SmartSheetSheet + "/rows/" + str(id)

    response = requests.request("GET", url, headers=getHeader(SmartSheetBearer))
    json_data = json.loads(response.text)

    smart_values = {'voornaam': 5057618067646340, 'achternaam': 2805818253961092, 'bedrijf': 7309417881331588, 'email': 1679918347118468, 'telefoon': 6183517974488964, 'functie': 3931718160803716, 'sessie1': 1507243481950084, 'sessie2': 6010843109320580, 'sessie3': 3759043295635332, 'comments': 8262642923005828 }

    for i in json_data['cells']:
        for key, value in smart_values.items():
            if i["columnId"] == value:
                if "value" in i:
                    values[key] = i["value"]
                else:
                    values[key] = ''
    return values
    
# Simple POST call for post message in WebEx Teams Room
def postWebEx(message):
    payload = {
        "roomId": WebExRoom,
        "markdown": message
    }   
    
    response = requests.request("POST", WebExUrl, data=payload, headers=getHeader(WebExBearer))    

# Custom function to parse some text for an event
def createWebExReg(data):
    
    date = []
    if data["sessie1"] is True:
        date.append('04/07')
    if data["sessie2"] is True:
        date.append('25/07')
    if data["sessie3"] is True:
        date.append('Visit me')

    payload = "Name: **%s %s** | Company:  **%s** | Email: %s | Function: **%s** | Date: **%s**" % (data["voornaam"], data["achternaam"], data["bedrijf"], data["email"], data["functie"], ', '.join(date))

    return payload

# Get message if question have been asked to the bot
def getMessage(messageId):
    url = WebExUrl + str(messageId)
    response = requests.request("GET", url, headers=getHeader(WebExBearer))

    json_data = json.loads(response.text)   
    cleartext = json_data['text']
    cleartext = cleartext.replace(WebExBotName, '')

    if 'reg' in cleartext:
        checkregistrations()
    else: 
        postWebEx("Hmm, I'm not able to do that :( Please ask my master (Wouter) to add this. You can ask me 'registrations' and I'll list the current registrations.")

# Checks the registrations in the SmartSheet Sheet. 
def checkregistrations():
    url = SmartSheetUrl + SmartSheetSheet
    response = requests.request("GET", url, headers=getHeader(SmartSheetBearer))
    json_data = json.loads(response.text)
    
    values = []
    allrows = []
    for i in json_data['rows']:
        allrows.append(i['id'])

    number = len(allrows)
    if number > 0:
        for z in allrows:
            smartsheet_row_text = getSmartText(z)
            values.append(str(createWebExReg(smartsheet_row_text)))

        postWebEx('Here you have the current [registrations](https://app.smartsheet.com/sheets/JjP6mC7rW7468RJF8fVxGRRXmVRfG8WJqQwC9Jg1?view=grid) ('+ str(number) +'): \n- %s' % '\n- '.join(values))

    else: 
        postWebEx('No [registrations](https://app.smartsheet.com/sheets/JjP6mC7rW7468RJF8fVxGRRXmVRfG8WJqQwC9Jg1?view=grid) yet.')

## listens to incoming packets from SmartSheet WebHook and WebEx Teams webhook
def webhookListener(request):
    
    if 'Smartsheet-Hook-Challenge' in request.headers:
        header = request.headers
        challenge = header['Smartsheet-Hook-Challenge']
        a = {'smartsheetHookResponse': challenge }
        return a
    
    else: 
        body = json.loads(request.data)
        if 'data' in body:
            messageId = body['data']['id']
            user_id = body["data"]["personId"]
            if user_id != 'Y2lzY29zcGFyazovL3VzL1BFT1BMRS8wOTZkNjljMC00N2E2LTQwM2YtOTk1Ni03OGJjNzQyYTYwOGE':
                getMessage(messageId)
        else:
            for i in body['events']:
                if i["objectType"] == 'row' and i["eventType"] == 'created':
                    row_id = i["id"] 

            smartsheet_row_text = getSmartText(row_id)
            postWebEx('Niewe persoon heeft zich aangemeld: ' + str(createWebExReg(smartsheet_row_text)))              
        return json.dumps({'success':True}), 200, {'ContentType':'application/json'}             
