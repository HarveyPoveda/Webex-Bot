import requests.packages.urllib3
requests.packages.urllib3.disable_warnings()
from flask import Flask
from flask import request
app = Flask(__name__)
import requests
import json
import smartsheet
from datetime import datetime
import pprint
import os


#Access token can be generated using the CISCO BOT account and Bot Email also can find in there

botEmail = "hpovedaf-local-bot@webex.bot"
accessToken = os.environ['TOKEN'] 
host = "https://api.ciscospark.com/v1"
headers = {"Authorization": "Bearer %s" % accessToken,"Content-Type": "application/json"}
authetication_sheet=os.environ['autheticationsheet']



@app.route('/', methods=['POST'])
def getMessage():
    message_json= request.json
    #pprint.pprint(json.dumps(message))
    messageId=message_json["data"]["id"]
    room_id=message_json["data"]["roomId"]
    message = requests.get(host + "/messages/" + messageId, headers=headers)
    Answer(message,room_id)
    return ""


#
def sendMessage(message, room_id):
    cisco = """
.:|:.:|:. CISCO"""
    payload = {"roomId": room_id, "text": message + "\n\n" + cisco}
    requests.post(host+"/messages/", data=json.dumps(payload), headers=headers)


#
def Answer(message_request, room_id):
    global bandera
    message_json = json.loads(message_request.text)
    #pprint.pprint(json.dumps(message_json))
    message=message_json["text"]
    WhoisEmail = message_json["personEmail"]
    if WhoisEmail != botEmail:
        if bandera==0:
            whois = requests.get(host + "/people?email=" + WhoisEmail, headers=headers)
            WhoisName_json = json.loads(whois.text)
            print(WhoisName_json)
            WhoisName = WhoisName_json["items"][0]["firstName"]
            if 'hola' in message.lower():
                ans = "Hola "+WhoisName+"!\n" + menu()
                sendMessage(ans, room_id)
            elif "1" in message.lower():
                ss_client = smartsheet.Smartsheet(authetication_sheet)
                sheet_id = 7812041256789892
                columnId = 8202839215368068
                row_id = 7005051021485956
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId, include_all=True)
                # print(json.loads(str(response)))
                messageString = "El total de requerimientos atendidos es de: " + json.loads(str(response))["data"][0][
                    'displayValue'] + "\nal día de hoy: " + datetime.now().strftime('%A %d %B %Y')

                # messageString = "it is " + datetime.now().strftime('%A %d %B %Y') + " today."
                sendMessage(messageString, room_id)
            elif "3" in message.lower():
                ss_client = smartsheet.Smartsheet(authetication_sheet)
                #row_id = 7005051021485956
                #response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId, include_all=True)
                # print(json.loads(str(response)))
                messageString = "Ingrese el requerimiento a buscar"
                bandera=3

                # messageString = "it is " + datetime.now().strftime('%A %d %B %Y') + " today."
                sendMessage(messageString, room_id)
            else:
                ans="Lo siento, no entiendo tu petición\n"+menu()
                sendMessage(ans, room_id)
        elif bandera==3:
            bandera=0
            ss_client = smartsheet.Smartsheet(authetication_sheet)
            sheet_id = 5067660233860996
            response = ss_client.Search.search(message)
            row_id=json.loads(str(response))["results"][0]["objectId"]
            #columnId_requerimiento = 8796249076852612 # id de requerimiento
            columnId_Asunto = 4292649449482116
            columnId_Estado = 4714861914548100
            response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId_Asunto, include_all=True)
            messageString = "asunto: " + json.loads(str(response))["data"][0]['displayValue'] + "\n"
            response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId_Estado, include_all=True)
            messageString+="Estado: " + json.loads(str(response))["data"][0]['displayValue']+ "\n"
            messageString+="\nal día de hoy: " +datetime.now().strftime('%A %d %B %Y')
            sendMessage(messageString, room_id)
            #print(result)


def menu():
    messageString = """¿En que te puedo ayudar?
1. Total de requerimientos atendidos 
2. Total de requerimientos de planeación
3. Buscar Caso"""

    return messageString

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int("5000"))

