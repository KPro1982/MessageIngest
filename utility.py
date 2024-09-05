# Utility.PY
import pandas as pd
import win32com.client


from anvil.tables import app_tables
from marshmallow import Schema, fields, post_load
from pprint import pprint
from datetime import date
from langchain_core.messages import HumanMessage
from langchain_core.prompts import ChatPromptTemplate
from langchain.prompts import PromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_community.document_loaders import OutlookMessageLoader
from langchain_community.chat_models import ChatOpenAI
# from langchain_community.llms import openai
from langchain.chains.summarize import load_summarize_chain
from langchain.docstore.document import Document
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import ChatOpenAI
from langchain_text_splitters import TokenTextSplitter
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.pydantic_v1 import BaseModel, Field
from typing import Optional
from langchain_core.pydantic_v1 import BaseModel, Field
from langchain_mistralai import ChatMistralAI

class Alias_Schema(BaseModel):
    """Information about a legal case."""

    # ^ Doc-string for the entity Person.
    # This doc-string is sent to the LLM as the description of the schema Person,
    # and it can help to improve extraction results.

    # Note that:
    # 1. Each field is an `optional` -- this allows the model to decline to extract it!
    # 2. Each field has a `description` -- this description is used by the LLM.
    # Having a good description can help improve extraction results.
    plaintiff: Optional[str] = Field(default=None, description="The Plaintiff's last name ONLY in a legal case. Do not include the first name. This usually preceeds the v.")
    defendant: Optional[str] = Field(default=None, description="The Defendant's last name ONLY in a legal case. Do not include the first name. This usually follows the v.")
    casename: Optional[str] = Field(default=None, description="The case name in the form: Plaintiff's last name v. Defendant's last name")
    


clientDict = ""
aliasesList = ""
matterList = ""




def get_smtp_address(entry):
    try:
        if entry.AddressEntry.Type == "EX":  # If the address is from an Exchange user
            exchange_user = entry.AddressEntry.GetExchangeUser()
            if exchange_user:
                return exchange_user.PrimarySmtpAddress
            else:
                # Try using PropertyAccessor to get the SMTP address
                return entry.AddressEntry.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
        else:
            return entry.Address  # For non-Exchange (SMTP) addresses
    except Exception as e:
        print(f"Failed to get SMTP address: {e}")
        return None
    

def get_all_email_addresses(mail_item):
    email_addresses = []

    # Add sender's SMTP address
    if mail_item.Sender is not None:
        sender_email = get_sender_smtp_address(mail_item)
        if sender_email:
            email_addresses.append(sender_email)

    # Add all recipients (To and CC)
    for recipient in mail_item.Recipients:
        email = get_smtp_address(recipient)
        if email:
            email_addresses.append(email)
    
    return email_addresses

def get_sender_smtp_address(mail_item):
    # # Get the PropertyAccessor for the mail item
    # prop_accessor = mail_item.PropertyAccessor
    
    # # Get the PR_SMTP_ADDRESS property using the MAPI property tag
    # smtp_address = prop_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
    
    # return smtp_address

    # Attempt to get the SMTP address using the AddressEntry object
    if mail_item.SenderEmailType == "EX":
        # For Exchange type emails, the address is more complex and requires resolving
        sender = mail_item.Sender
        address_entry = sender
        if address_entry.Type == "EX":
            # Get the actual SMTP address
            smtp_address = address_entry.GetExchangeUser().PrimarySmtpAddress
        else:
            smtp_address = mail_item.SenderEmailAddress
    else:
        # For regular SMTP emails, the SenderEmailAddress should suffice
        smtp_address = mail_item.SenderEmailAddress
    
    return smtp_address

def remove_image_files(file_list):
    # Define image file extensions
    image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.webp', '.svg'}

    # Filter out files that contain any of the image extensions

    filtered_list = [file for file in file_list if not any(ext in file for ext in image_extensions)]
    
    return filtered_list
    


def IsExternal(email_list):
    
    for email in email_list:
        parts = email.split('@')
        if "@ohaganmeyer.com" not in email:
            return "external"
    return "internal"

def check_cravens_role(mail_item):
    # Get the sender's address
    sender_addy = get_sender_smtp_address(mail_item)
    # Get the list of recipients (To and CC)
    recipients_to = mail_item.To
    recipients_cc = mail_item.CC
    
    # Check if Cravens is the sender
    if "cravens" in sender_addy.lower():
        return "sender"
    
    # Check if Cravens is a direct recipient (To field)
    if "cravens" in recipients_to.lower():
        # If Cravens is the only recipient in To field
        if recipients_to.lower().count("cravens") == 1 and recipients_to.lower().count(";") == 0:
            return "exclusive"
        return "recipient"
    
    # Check if Cravens is in the CC field
    if "cravens" in recipients_cc.lower():
        return "cc"
    
    # If Cravens is neither the sender nor a recipient
    return None
    
def get_attachments(msg_path):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    msg = outlook.OpenSharedItem(msg_path)
    return msg.Attachments

def get_msgdata(msg_path, api_key):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    msg = outlook.OpenSharedItem(msg_path)

    

        

    addys = get_all_email_addresses(msg)

    external = IsExternal(addys)
    
    role = check_cravens_role(msg)
    result_dict = {
        "role": role,
        "domain": external
    }
            

    return result_dict


def Simplify_Attachment(apikey, attachment_name, Examples):
    llm = ChatOpenAI(temperature=0.3, model_name="gpt-3.5-turbo-16k", api_key=apikey)
 
    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                """You are an attorney in the State of California.
                You will given the following a file name
                Your task is to simply the file name into a plain english format that describes the contents without extaneous information.
                You may consider the following examples for word choice and formatting.
                "EXAMPLE SIMPLIFIED FILE NAMES:""" + Examples,
            ),
            # Please see the how-to about improving performance with
            # reference examples.
            # MessagesPlaceholder('examples'),
            ("human", "Simplify the following attachment file name: {attachment_name}"),
        ]
    )

    runnable = prompt  | llm

    
    
    result =  runnable.invoke({"attachment_name": attachment_name}).content
    
    return result
