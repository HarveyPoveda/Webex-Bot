import requests.packages.urllib3
requests.packages.urllib3.disable_warnings()
from flask import Flask
from flask import request
app = Flask(__name__)
import requests
import json
import smartsheet
from datetime import datetime

############# How to set a webhook for Cisco Spark
# Use ngrok for making localhost:port to public safely (Install ngrok free )
# use command 'ngrok http <port>'
# Now Register a webhook at cisco spark using https://developer.ciscospark.com/endpoint-webhooks-post.html#
# in this link use 
# authorization: Bot access token
# targetUrl: <ngrok link>/
# resource: messages
# event : create
# filter: roomId=room_id (as below)
# This will register ngrok link to webhook 
# now run this file using python starterbotCiscoSpark.py
#############
## Supporting links 
#1 - http://madumal.piyasundara.org/2017/02/python-chat-bot-integration-with-cisco.html
#2 - https://communities.cisco.com/community/developer/spark/blog/2016/07/18/spark-bot-demo
#3 - https://developer.ciscospark.com/endpoint-rooms-get.html to get room ids
#4 - https://developer.ciscospark.com  to create a new bot 
#5 - https://learninglabs.cisco.com/lab/collab-spark-chatops-bot-itp/step/1  Lab for many bots ay cisco spark
#6 - https://learninglabs.cisco.com/lab/collab-sparkwebhook/step/1  webhook understanding good one

#Access token can be generated using the CISCO BOT account and Bot Email also can find in there
botEmail = "hpovedaf-local-bot@webex.bot"#bot's email address
accessToken = "Y2Y5YjY2N2QtYTFiNS00Y2MzLWE5MmYtNjQyYmI0MWQ4ODQwMzNjYmE3NmEtNWYx_PF84_1eb65fdf-9643-417f-9974-ad72cae0e10f" #Bot's access token
host = "https://api.ciscospark.com/v1/"#end point provided by the CISCO Spark to communicate between their services
server = "localhost" #Web hook won't work until the server sets up
port = 4000
headers = {"Authorization": "Bearer %s" % accessToken,"Content-Type": "application/json"}
room_id = "Y2lzY29zcGFyazovL3VzL1JPT00vNzFjODc4YjAtMTM2Yy0xMWU5LTlhZWQtYTcwOTRkNjYyOTA2"  ## Room id where bot is added@app.route('/webhook', methods=['POST'])
authetication_sheet="zx2ag6oscmzx5nw8r31c5i85xy"


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
    app.run(host=server,port=port,debug=True)

