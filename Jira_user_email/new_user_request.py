'''
    New User Request Script
    Created 4.26.22
    Ver 1.0
    By luismtz 
'''

import json
import win32com.client as cl
import requests
from requests.auth import HTTPBasicAuth

#importing config.py to access API Variables.
import config


# Connection setup for JIRA 
auth = HTTPBasicAuth(config.user_email, config.api_key)
headers = { 
    "Accept": "application/json"
}

def send_to_outlook(msg): 
    outlook = cl.Dispatch('outlook.application')
    
    #Create mail object
    mail = outlook.CreateItem(0)
    mail.Display(False)
    mail.To = config.support_email
    mail.BCC = config.net_email
    mail.Subject = msg['subject']
    mail.Body = msg['body']

    # TODO add a way to include the signature into to message. 
    
    

def create_msg(fields): 

    list_name = fields["full_name"].split(" ")
    list_name.reverse() 
    dname = ", ".join(list_name)
    l_name = list_name[0]
    f_name = list_name[-1]

    body = """
            Hello EM, 

            Please create the following a new user account for {name} - {dept} 

            Department: {dept}
            Group: {dept}
            First Name: {fname}
            Last Name: {lname}
            Full Name: {name}
            Display Name: {dname}
            Desired E-Mail: {email}
            UCLA Logon ID: {ucla}
            UID: {uid}
            Affiliation: {affiliation}
            Account Type: O365

            Also please add to the following group: 
            Ex2k3 HumAll


            Note: 
            """.format(name = fields["full_name"], dept = fields["department"], lname = l_name, fname = f_name, dname = dname, email = fields["desired_email"], ucla = fields["ucla_logon"], affiliation = fields["affiliation"], uid = fields["uid"])

    subject = "Humnet Account Request for {} - {}".format(fields['full_name'], fields['department'])

    msg = {
        'body': body, 
        'subject': subject
    }
    
    return msg
def get_ticket(num): 
    #TODO -- Connect to the API 
    url = config.api_url+"HUM-{}".format(num)
    #print(url)
    response = requests.request(
            "GET", 
            url, 
            headers=headers, 
            auth=auth
        )
    #TODO -- Store Object in dictionary
    if response.status_code == 200: 
        issue = json.loads(response.text)
        issue_fields = issue["fields"]

        msg_fields = { 
            "desired_email": issue_fields["customfield_10100"], 
            "department": issue_fields["customfield_10078"]["value"],
            "ucla_logon": issue_fields["customfield_10149"],
            "affiliation": issue_fields["customfield_10163"]["value"],
            "full_name": issue_fields["customfield_10164"],
            "uid": issue_fields["customfield_10252"]
            #"note": issue_fields["description"]["content"][0]["content"][0]["text"]
        }
    return msg_fields

def main(): 
    while True: 
        try: 
            ticket_num = int(input("Enter Ticket Number: "))
            break 
        except: 
            print("Your entry was not valid. Please try again. You do not need to enter the project code.")
    
    fields = get_ticket(ticket_num)
    msg = create_msg(fields)
    send_to_outlook(msg)
    

if __name__ == "__main__": 
    main()

