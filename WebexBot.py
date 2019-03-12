import requests.packages.urllib3
requests.packages.urllib3.disable_warnings()
from flask import Flask
from flask import request
app = Flask(__name__)
import requests
import json
import smartsheet
from datetime import datetime,timedelta
import os


#Access token can be generated using the CISCO BOT account and Bot Email also can find in there

botEmail = "hpovedaf-local-bot@webex.bot"
accessToken = os.environ['TOKEN'] 
host = "https://api.ciscospark.com/v1"
headers = {"Authorization": "Bearer %s" % accessToken,"Content-Type": "application/json"}
authetication_sheet=os.environ['autheticationsheet']
bandera=0


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
    cisco = "\n\nInformación al día de hoy: " + (datetime.utcnow() +timedelta(hours=-5)).strftime("%c") + "\n" + ".:|:.:|:. CISCO"
    payload = {"roomId": room_id, "text": message  + cisco}
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
            print("quien esta enviando la info:\n"+str(WhoisName_json))
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
                    'displayValue']

                # messageString = "it is " + datetime.now().strftime('%A %d %B %Y') + " today."
                sendMessage(messageString, room_id)
            elif "2" in message.lower():
                ss_client = smartsheet.Smartsheet(authetication_sheet)
                sheet_id = 7812041256789892
                columnId = 6206190523836292
                row_id = 7005051021485956
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId, include_all=True)
                # print(json.loads(str(response)))
                messageString = "El total de requerimientos en ejecución son: " + json.loads(str(response))["data"][0][
                    'displayValue']

                sendMessage(messageString, room_id)
            elif "3" in message.lower():
                ss_client = smartsheet.Smartsheet(authetication_sheet)
                sheet_id = 7812041256789892
                columnId = 1522017586440068
                row_id = 7005051021485956
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId, include_all=True)
                # print(json.loads(str(response)))
                messageString = "El total de requerimientos en planeación son: " + json.loads(str(response))["data"][0][
                    'displayValue']
                sendMessage(messageString, room_id)
            elif "4" in message.lower():
                messageString = "Ingrese el requerimiento a buscar"
                bandera=4
                sendMessage(messageString, room_id)
            elif "5" in message.lower():
                week= int((datetime.utcnow() +timedelta(hours=-5)).strftime("%W"))+1
                ss_client = smartsheet.Smartsheet(authetication_sheet)
                ##### for para buscar semana
                sheet_id = 6685130556237700
                columnSemana = 3092366245554052
                columnDias = 3065488340215684
                response = ss_client.Sheets.get_sheet(sheet_id)
                resjson=json.loads(str(response))
                for rowssheet in resjson["rows"]:
                    flag=0
                    for cell_rows in rowssheet["cells"]:
                        if cell_rows["columnId"]==columnSemana and cell_rows["value"]==week:
                            flag=1
                        if flag==1 and cell_rows["columnId"]==columnDias and 0<cell_rows["value"]<7:
                            row_id = rowssheet["id"]
                            print("semana encontrada: "+str(cell_rows["value"]),"dias :"+str(cell_rows["value"]),sep="\n")

                # response = ss_client.Search.search_sheet(sheet_id, week)
                # row_id = json.loads(str(response))["results"][0]["objectId"]
                columnRemoto = 8132037921007492
                columnNocturno= 813688526530436
                messageString = "Semana del año número " + str(week) + "\n"
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnRemoto, include_all=True)
                messageString += "El ingeniero en turno remoto es: " + json.loads(str(response))["data"][0][
                    'displayValue']+"\n"
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnNocturno, include_all=True)
                messageString += "El ingeniero en turno nocturno es: " + json.loads(str(response))["data"][0][
                    'displayValue']
                sendMessage(messageString, room_id)
            elif message.lower()=="prueba-secret":
                week = int((datetime.utcnow() + timedelta(hours=-5)).strftime("%W")) + 1
                ss_client = smartsheet.Smartsheet(authetication_sheet)
                sheet_id = 6685130556237700
                pruebaId = 3092366245554052
                response = ss_client.Search.search_sheet(sheet_id, week)
                row_id = json.loads(str(response))["results"][0]["objectId"]
                messageString = "semana " + str(week) + "\n"
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, pruebaId, include_all=True)
                messageString += "Valor encontrado " + json.loads(str(response))["data"][0][
                    'displayValue'] + "\n"

                sendMessage(messageString, room_id)
            else:
                ans="Lo siento, no entiendo tu petición\n"+menu()
                sendMessage(ans, room_id)
        elif bandera==4:
            bandera=0
            ss_client = smartsheet.Smartsheet(authetication_sheet)
            sheet_id = 5067660233860996
            try:
                response = ss_client.Search.search_sheet(sheet_id,message)
                print(json.loads(str(response)))
                row_id=json.loads(str(response))["results"][0]["objectId"]
                #columnId_requerimiento = 8796249076852612 # id de requerimiento
                columnId_id = 8796249076852612
                columnId_Asunto = 4292649449482116
                columnId_Solicitud = 1724276186343300
                columnId_Estado = 4714861914548100
                columnId_esfuerzo = 3993496387381124
                columnId_propietario = 4547598708172676

                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId_id, include_all=True)
                messageString = "Id: " + json.loads(str(response))["data"][0]['displayValue'] + "\n"
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId_Asunto, include_all=True)
                messageString += "Asunto: " + json.loads(str(response))["data"][0]['displayValue'] + "\n"
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId_Solicitud, include_all=True)
                messageString += "Tipo de solicitud: " + json.loads(str(response))["data"][0]['displayValue'] + "\n"
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId_Estado, include_all=True)
                messageString+="Estado: " + json.loads(str(response))["data"][0]['displayValue']+ "\n"
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId_esfuerzo, include_all=True)
                messageString += "Tiempo de Esfuerzo (min): " + json.loads(str(response))["data"][0]['displayValue'] + "\n"
                response = ss_client.Cells.get_cell_history(sheet_id, row_id, columnId_propietario, include_all=True)
                messageString += "Propietario: " + json.loads(str(response))["data"][0]['displayValue'] + "\n"
                sendMessage(messageString, room_id)
            except:
                messageString = "No se encontró ID\n"
                sendMessage(messageString, room_id)


def menu():
    messageString = """¿En que te puedo ayudar?
1. Total de requerimientos atendidos. 
2. Total de requerimientos de ejecución.
3. Total de requerimientos de planeación.
4. Buscar Caso.
5. Rotación BO"""


    return messageString

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=os.environ.get("PORT", 5000))

