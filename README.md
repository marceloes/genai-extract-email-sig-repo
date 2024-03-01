# Extract Emails and Signatures Python Script

This code came along one afternoon where I was looking at how to get the email addresses, names and job titles from my customers so I can create a reference document to use when talking to them about GenAI.  
Then, it hit me: Why not use GenAI to do this?  
Here's how I solved this problem:  
1. Opened Visual Studio Code and made sure to have GitHub Copilot in place (both Copilot and Copilot Chat extensions)
1. Asked Copilot Chat to write me Python code to go thru emails in Outlook app (I'm in Windows) and check the senders email addresses
1. Asked Copilot to write sample code to connect to Azure OpenAI
1. Asked Copilot to create a prompt to extract signatures from emails
1. Tweaked that prompt a fair good amount... a few iterations and added some constraints to it
1. Merged the code to call Azure OpenAI with the code to go thru the emails
1. Ran a few tests, fixed some bugs... tweaked the prompt some more
1. Asked Copilot to remove any hardcoded keys, credentials
1. Asked Copilot to write me code to show usage and take the domain name from the command-line arguments
1. Asked Copilot to write a markdown README that I could use when publishing this to GitHub

All of this took me 2.5 hours. And I came with 0 knowledge of Outlook handling win32 COM objects in Python, the data structure, etc. That's the beauty of Gen AI!  
Also, I started this as a PowerShell script, thought it would be neat to do it in Python instead and asked Copilot to convert the code to Python. It even let me know some limitations of the win32com package and some tweaks it had to do.  
Enjoy!

---

This Python script is designed to extract email addresses and signatures from an Outlook inbox and save the results to a JSON file. It uses the Azure OpenAI service to extract email signatures.

**NOTE: This code was mostly generated using GitHub Copilot and my own few tweaks. It's not intended to be the best solution for this problem, but more of a proof-of-concept to showcase the capabilities with GenAI**
**NOTE 2: Be mindful that this code is collecting and storing email address, people's names and job titles which is PII. Please follow your companie's privacy and security procedures**

## Code Explanation

The script starts by setting up a connection to the Azure OpenAI service. It then defines several functions:

- `extract_email(s: str) -> List[str]`: This function uses a regular expression to find and return all email addresses in a given string.

- `extract_email_signature(body: str) -> str`: This function uses the Azure OpenAI service to extract an email signature from a given email body. It sends a message to the OpenAI service instructing it to extract the name, job title, and company from the email signature. The OpenAI service returns a JSON string containing this information.

- `process_data_extraction(domain_name: str)`: This function uses the `win32com.client` module to interact with the Outlook application. It opens the Inbox folder and loops through each email. For each email, it extracts the sender's email address and, if the email address contains the specified domain name, it uses the `extract_email_signature` function to extract the email signature. It stores the email address and signature in a dictionary. The function returns this dictionary.

The script then checks if a domain name was provided as a command-line argument. If so, it calls the `process_data_extraction` function with the provided domain name and saves the results to a JSON file. If no domain name was provided, it prints usage instructions.

## Requirements

To run this script, you need:

- Python 3.6 or later.
- The `os`, `re`, `win32com.client`, `json`, `typing`, and `sys` modules, which are part of the Python Standard Library and thus should be available in any standard Python installation.
- The `openai` package, which can be installed with `pip install openai`.
- The environment variables `AZURE_OPENAI_KEY`, `AZURE_OPENAI_API_VERSION`, `AZURE_OPENAI_ENDPOINT` and `AZURE_DEPLOYMENT_NAME` set to your environment. Create those resources before you run this. Check out this [documentation](https://learn.microsoft.com/en-us/azure/ai-services/openai/how-to/create-resource?pivots=web-portal) on how to do that. 
- Microsoft Outlook installed and configured with an email account.
- A command-line argument specifying the domain name to filter email addresses by.

## Running the Script

To run the script, navigate to the directory containing the script in your command line or terminal, and then type:

```bash
python extract-emails.py <domain_name>
```

Replace ```<domain_name>``` with the domain name to filter email addresses by. For example:

```bash
python extract-emails.py contoso.com
```
 This will extract all email addresses from the Outlook inbox that contain ```contoso.com``` and save the results to a file called ```contoso.com.json```.