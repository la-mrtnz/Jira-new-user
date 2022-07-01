'''
    New User Request Script
    Created 4.26.22
    Ver 1.0
    By luismtz 
'''

import json
from tkinter import messagebox
import win32com.client as cl
import requests
from requests.auth import HTTPBasicAuth
from tkinter import *
from tkinter import ttk

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

    body = """Hello EM, 

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

Thank you, 

            """.format(name = fields["full_name"], dept = fields["department"], lname = l_name, fname = f_name, dname = dname, email = fields["desired_email"], ucla = fields["ucla_logon"], affiliation = fields["affiliation"], uid = fields["uid"])

    subject = "Humnet Account Request for {} - {}".format(fields['full_name'], fields['department'])

    msg = {
        'body': body, 
        'subject': subject
    }
    
    return msg
def get_ticket(num): 
    #-- Connect to the API 
    url = config.api_url+"HUM-{}".format(num)
    response = requests.request(
            "GET", 
            url, 
            headers=headers, 
            auth=auth
        )
    #-- Store Object in dictionary
    if response.status_code == 200: 
        issue = json.loads(response.text)
        issue_fields = issue["fields"]
        
        #TODO: Validate that the issue is the correct type. 
        msg_fields = { 
            "desired_email": issue_fields["customfield_10100"], 
            "department": issue_fields["customfield_10078"]["value"],
            "ucla_logon": issue_fields["customfield_10149"],
            "affiliation": issue_fields["customfield_10163"]["value"],
            "full_name": issue_fields["customfield_10164"],
            "uid": issue_fields["customfield_10252"]
        }
    else:
        messagebox.showerror("Ticket Number Error", "Invalid Ticket Number. Issue not found. Please try again.")
        return 0
    return msg_fields



def main(): 
    def retrieve():
        fields = get_ticket(ticket.get())
        if fields == 0: 
            return
        msg = create_msg(fields)
        send_to_outlook(msg)
        messagebox.showinfo("Success", "Email Successfully Created in Outlook!")
    root = Tk()
    root.title("New Email Request")
    root.geometry("300x150")
    root.columnconfigure(1, weight=1)
    root.columnconfigure(2, weight=1)
    root.rowconfigure(0, weight=1)
    

    ticket_label = ttk.Label(root, text="Enter a ticket number below to generate a request email in outlook.", wraplength=200, justify=CENTER)
    ticket_label.grid(row=0, column=0, columnspan=3,padx=5, pady=5)

    number_label = ttk.Label(root, text="Hum-", )
    number_label.grid(row=2, column=1, sticky=E)

    ticket = IntVar()
    ticket_entry = ttk.Entry(root, textvariable=ticket, width=7)
    ticket_entry.grid(row=2, column=2,sticky=(W),padx=5)
    
    ticket_button = ttk.Button(root, text="Submit", command=retrieve)
    ticket_button.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

    root.mainloop()
    
    
    

if __name__ == "__main__": 
    main()

