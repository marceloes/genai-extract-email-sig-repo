import os
import re
import win32com.client
import json
from typing import List
from openai import AzureOpenAI
import sys

# Azure OpenAI connection
AIclient = AzureOpenAI(
    azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT"), 
    api_key = os.getenv("AZURE_OPENAI_KEY"),  
    api_version = os.getenv("AZURE_OPENAI_API_VERSION")
)


def extract_email(s: str) -> List[str]:
    email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    return re.findall(email_regex, s)

def extract_email_signature (body: str) -> str:
    # Extract email signature using OpenAI

    message_text = [{"role":"system",
                     "content":"Extract Name, Job Title and Company from email signature in the email body. \
                      Signatures are typically found at the end of the email body. \
                      In any JSON response, always use double quotes for properties and values. \
                       ---- EXAMPLE SIGNATURE ----- \
                       --- \
                       John Doe \
                       Senior Software Engineer \
                       Acme Corp \
                       johndoe@email.com \
                       --- \
                       --- EXAMPLE RESPONSE --- \
                        {\"Status\": \"Success\" \
                         \"Name\": \"John Doe\", \
                         \"Job Title\": \"Senior Software Engineer\", \
                         \"Company\": \"Acme Corp\"} \
                      --- \
                       ---- EXAMPLE SIGNATURE ----- \
                       --- \
                       John Doe \
                       Senior Software Engineer \
                       1234 Main Street, Suite 456, Anytown, FL \
                       Acme Corp \
                       johndoe@email.com \
                       --- \
                       --- EXAMPLE RESPONSE --- \
                        {\"Status\": \"Success\" \
                         \"Name\": \"John Doe\", \
                         \"Job Title\": \"Senior Software Engineer\", \
                         \"Company\": \"Acme Corp\"} \
                      --- \
                       ---- EXAMPLE SIGNATURE ----- \
                       --- \
                       Mary Jane - Accountant - Contoso\
                       --- \
                       --- EXAMPLE RESPONSE --- \
                        {\"Status\": \"Success\" \
                         \"Name\": \"Mary Jane\", \
                         \"Job Title\": \"Accountant\", \
                         \"Company\": \"Contoso\"} \
                      --- \
                      If you cannot find a signature, respond with the following JSON: \
                        {\"Status\": \"No signature found\" \
                         \"Name\": \" \", \
                         \"Job Title\": \" \", \
                         \"Company\": \" \"} \
                      If you cannot find a name, job title or company, respond with the following JSON format, putting 'N/A' on any field that was not found, like in the example below where Job Title was not found but Name and Company were found: \
                        {\"Status\": \"Partial Success\" \
                         \"Name\": \"Marc Smith\", \
                         \"Job Title': \"N/A\", \
                         \"Company': \"Contoso\"} \
                      --- EMAIL BODY ---- " + body + "--- END OF EMAIL BODY ---"}]
    
    completion = AIclient.chat.completions.create(
        model=os.getenv("AZURE_DEPLOYMENT_NAME"), # model = "deployment_name"
        messages = message_text,
        temperature=0.1,
        max_tokens=400,
        top_p=0.95,
        frequency_penalty=0,
        presence_penalty=0,
        stop=None
    )

    return completion.choices[0].message.content

def process_data_extraction(domain_name: str):   
    # Load the Outlook COM object
    Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Get the Inbox folder
    inbox = Outlook.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

    # Initialize a list to store email addresses
    emails_and_sigs_dict = {}

    # Get count of all inbox items
    item_count = inbox.Items.Count

    current_step = 0

    # Loop through each email in the Inbox
    for mail_item in inbox.Items:
        # Extract sender and recipient addresses
        if hasattr(mail_item, "SenderEmailAddress"):
            email_address = extract_email(mail_item.SenderEmailAddress)
        else:
            if hasattr(mail_item, "sender_address"):
                email_address = extract_email(mail_item.sender_address)
            else:
                email_address = None

        # Combine sender and recipient addresses
        all_addresses = email_address

        if all_addresses:
            # Filter addresses based on specified domains
            filtered_addresses = [address for address in all_addresses if domain_name.lower() in address.lower()]

            # Add filtered addresses to the list only if filtered addresses are not empty
            if filtered_addresses:
                email = filtered_addresses[0].lower()
                if (email not in emails_and_sigs_dict) or (email in emails_and_sigs_dict and emails_and_sigs_dict[email] == None):
                    # Extract email signature using OpenAI
                    trimmed_email_body = mail_item.Body.split("\r\nFrom: ")[0]
                    if len(trimmed_email_body) > 10:
                        email_sig_json = extract_email_signature(trimmed_email_body)
                        email_sig = json.loads(email_sig_json)

                        # Check if returned name is N/A in the name field
                        if (email_sig["Status"] == "Success" or email_sig["Status"] == "Partial Success"): # and email_sig["Name"] != "YOUR NAME":     <-- Add this in case you get a lot of your own signatures in the email
                            email_sig["Email"] = email
                            emails_and_sigs_dict[email] = email_sig
                            print(f"\r\nAdded email with signature: {email_sig['Email']}: {email_sig['Name']} - {email_sig['Job Title']} - {email_sig['Company']}")
                        else:
                            emails_and_sigs_dict[email] = None
                            print(f"\r\nAdded email without signature: {email}")                       

        # Calculate the percentage completed
        percent_complete = (current_step / item_count) * 100
        print(f"Extracting email addresses - Processing email {current_step} of {item_count} - {int(percent_complete)}% complete", end='\r')

        current_step += 1

    return emails_and_sigs_dict

# Main

if len(sys.argv) > 1:
    domain_name = sys.argv[1]
    with open(f"{domain_name}.json", "w") as outfile:
        result = process_data_extraction(domain_name)
        json.dump(result, outfile)
    print()
    print(result)
else:
    print("Usage: python extract-emails.py <domain_name>")
    print("Example: python extract-emails.py @contoso.com")
    print("This will extract all email addresses from the Outlook inbox and save the results to a file called <domain_name>.json")
    print("Note: You must have Outlook installed and configured with an email account to use this script.")
    print("Note: You must have the OpenAI Python SDK installed to use this script.")
    print("Note: You must have an OpenAI API key to use this script.")
