import requests.packages.urllib3
requests.packages.urllib3.disable_warnings()
from flask import Flask
from flask import request
app = Flask(__name__)
import requests
import json
import smartsheet
from datetime import datetime
import os


#Access token can be generated using the CISCO BOT account and Bot Email also can find in there
botEmail = "hpovedaf-local-bot@webex.bot"#bot's email address
accessToken = os.environ['TOKEN'] #Bot's access token
host = "https://api.ciscospark.com/v1/"#end point provided by the CISCO Spark to communicate between their services
headers = {"Authorization": "Bearer %s" % accessToken,"Content-Type": "application/json"}
room_id =  os.environ['room_id'] ## Room id where bot is added@app.route('/webhook', methods=['POST'])
authetication_sheet=os.environ['autheticationsheet']


@app.route('/', methods=['POST'])
def get_tasks():
    messageId = request.json.get('data').get('id')
    messageDetails = requests.get(host+"messages/"+messageId, headers=headers)
    replyForMessages(messageDetails)
    return ""

#A function to send the message by particular email of the receiver
def sendMessage(message,  toPersonEmail):
    payload = {"roomId": room_id, "text": message}
    response = requests.post("https://api.ciscospark.com/v1/messages/", data=json.dumps(payload),  headers=headers)
    return response.status_code
    
#A function to get the reply and generate the response of from the bot's side
## Add Bot logic for complicated bots 
def replyForMessages(response):

    responseMessage = response.json().get('text')
    toPersonEmail = response.json().get('personEmail')
    print (responseMessage)
    if toPersonEmail != botEmail:
        if 'hola' in responseMessage:
            messageString = """
Hola, En que te puedo ayudar?
1. Total de requerimientos atendidos
2. Total de requerimientos de planeación
"""
            sendMessage(messageString,  toPersonEmail)
        elif "1".lower() in responseMessage.lower():
            ss_client = smartsheet.Smartsheet(authetication_sheet)
            """
            sheet_id = 9005334300780420
            columnId= 4255539602450308
            row_id=4600688007243652
            """
            sheet_id = 7812041256789892
            columnId= 8202839215368068
            row_id=7005051021485956
            response=ss_client.Cells.get_cell_history(sheet_id,row_id,columnId,include_all=True)
            #print(json.loads(str(response)))
            messageString = "El total de requerimientos atendidos es de: "+json.loads(str(response))["data"][0]['displayValue']+"\nal día de hoy: " + datetime.now().strftime('%A %d %B %Y')


            #messageString = "it is " + datetime.now().strftime('%A %d %B %Y') + " today."
            sendMessage(messageString,  toPersonEmail)
        else:
            messageString = 'Sorry! I was not programmed to answer this question!'
            sendMessage(messageString, toPersonEmail)

if __name__ == "__main__":
    app.run()

