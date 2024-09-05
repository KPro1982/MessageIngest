
import pandas as pd



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


def ClientMatter(subject, apiKey, aliasList):
    clientDict = ""
    matterList = ""

    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                "You are an expert extraction algorithm. "
                "Only extract relevant information from the text. "
                "Only include last names of people. Do not include first name of people. "
                "If you cannot validate your answer by finding an exact match in the list return 'None'",
            ),
            # Please see the how-to about improving performance with
            # reference examples.
            # MessagesPlaceholder('examples'),
            ("human", "{subject}"),
        ]
    )




    
    llm = ChatOpenAI(temperature=0.3, model_name="gpt-3.5-turbo-16k", api_key=apiKey)
    runnable = prompt | llm.with_structured_output(schema=Alias_Schema) 

    
    
    result =  runnable.invoke({"subject": subject}).casename
    



    
    return result

def Narrative_Internal_Attachments(apiKey, msg_recipient, msg_from, msg_body, msg_subject, msg_attachments, NarrativeExamples):
    llm = ChatOpenAI(temperature=0.3, model_name="gpt-3.5-turbo-16k", api_key=apiKey)
    msg_body = reduce(msg_body)
    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                """You are a secretary working for attorney Daniel Cravens. Your job is to create a billing entry that succinctly summarizes the work that Daniel Cravens performed based on the email provided. You must begin your billing entry with a verb. 
                Your billing entry should be based on top most email in the thread.
                Do NOT use the following verbs: file, schedule, transmit, code
                Do NOT inlcude the case name in the time entry
                In drafting the billing entry, you may consider the following examples for word choice and formatting.
                EXAMPLE BILLING ENTRIES:""" + NarrativeExamples,
            ),
            # Please see the how-to about improving performance with
            # reference examples.
            # MessagesPlaceholder('examples'),
            ("human", "Generate a detailed time entry from the following email that includes 'Draft and update' each attachment: {attachments} and any task that is mentioned in the topmost email in the thread {body}."),
        ]
    )

    runnable = prompt  | llm

    
    result =  runnable.invoke({"attachments": msg_attachments, "body": msg_body}).content
    
    return result    

def Narrative_Internal_Zero_Attachments(apiKey, msg_recipient, msg_from, msg_body, msg_subject, NarrativeExamples):
    llm = ChatOpenAI(temperature=0.3, model_name="gpt-3.5-turbo-16k", api_key=apiKey)

    msg_body = reduce(msg_body)
 
    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                """You are a secretary working for attorney Daniel Cravens. Your job is to create a billing entry that succinctly summarizes the work that Daniel Cravens performed based on the email provided. You must begin your billing entry with a verb. 
                Your billing entry should be based on top most email in the thread.
                Do NOT use the following verbs: file, schedule, transmit, code
                Do NOT inlcude the case name in the time entry
                In drafting the billing entry, you may consider the following examples for word choice and formatting.
                EXAMPLE BILLING ENTRIES:""" + NarrativeExamples,
            ),
            # Please see the how-to about improving performance with
            # reference examples.
            # MessagesPlaceholder('examples'),
            ("human", "Generate a detailed time entry from the following email: {body}."),
        ]
    )

    runnable = prompt  | llm

    
    
    result =  runnable.invoke({"body": msg_body}).content
    
    return result


# Sender_Attachments_External
def Narrative_SAE(apiKey, msg_recipient, msg_from, msg_body, msg_subject, msg_attachments, NarrativeExamples):
    llm = ChatOpenAI(temperature=0.3, model_name="gpt-3.5-turbo-16k", api_key=apiKey)
    msg_body = reduce(msg_body)
    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                """You are a secretary working for attorney Daniel Cravens. Your job is to create a billing entry that succinctly summarizes the work that Daniel Cravens performed based on the email provided. You must begin your billing entry with a verb. 
                Your billing entry should be based on top most email in the thread.
                Do NOT use the following verbs: file, schedule, transmit, code
                Do NOT inlcude the case name in the time entry
                In drafting the billing entry, you may consider the following examples for word choice and formatting.
                EXAMPLE BILLING ENTRIES:""" + NarrativeExamples,
            ),
            # Please see the how-to about improving performance with
            # reference examples.
            # MessagesPlaceholder('examples'),
            ("human", "Generate a detailed time entry from the following email that includes 'Draft and update' each attachment: {attachments} and any task that is mentioned in the topmost email in the thread {body}."),
        ]
    )

    runnable = prompt  | llm

    
    
    result =  runnable.invoke({"attachments": msg_attachments, "body": msg_body}).content
    
    return result

# Recipient_Attachments_External
def Narrative_RAE(apiKey, msg_recipient, msg_from, msg_body, msg_subject, msg_attachments, NarrativeExamples):
    llm = ChatOpenAI(temperature=0.3, model_name="gpt-3.5-turbo-16k", api_key=apiKey)
    msg_body = reduce(msg_body)
    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                """You are a secretary working for attorney Daniel Cravens. Your job is to create a billing entry that succinctly summarizes the work that Daniel Cravens performed based on the email provided. You must begin your billing entry with a verb. 
                Your billing entry should be based on top most email in the thread.
                Do NOT use the following verbs: file, schedule, transmit, code
                Do NOT inlcude the case name in the time entry
                An email with another attorney should begin with "Meet and confer with" then state the name of the attorney and then summarize the substance
                In drafting the billing entry, you may consider the following examples for word choice and formatting.
                EXAMPLE BILLING ENTRIES:""" + NarrativeExamples,
            ),
            # Please see the how-to about improving performance with
            # reference examples.
            # MessagesPlaceholder('examples'),
            ("human", "Generate a detailed time entry from the following email that includes 'Review and respond to [attachment name]' for each attachment: {attachments} and any task that is mentioned in the topmost email in the thread {body}."),
        ]
    )

    runnable = prompt  | llm

    
    
    result =  runnable.invoke({"attachments": msg_attachments, "body": msg_body}).content
    
    return result

def Narrative_Default_External(apiKey, msg_recipient, msg_from, msg_body, msg_subject, NarrativeExamples):
    llm = ChatOpenAI(temperature=0.3, model_name="gpt-3.5-turbo-16k", api_key=apiKey)

    msg_body = reduce(msg_body)
 
    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                """You are a secretary working for attorney Daniel Cravens. Your job is to create a billing entry that succinctly summarizes the work that Daniel Cravens performed based on the email provided. You must begin your billing entry with a verb. 
                Your billing entry should be based on top most email in the thread.
                Do NOT use the following verbs: file, schedule, transmit, code
                Do NOT inlcude the case name in the time entry
                An email with another attorney should begin with "Meet and confer with" then state the name of the attorney and then summarize the substance
                In drafting the billing entry, you may consider the following examples for word choice and formatting.
                EXAMPLE BILLING ENTRIES:""" + NarrativeExamples,
            ),
            # Please see the how-to about improving performance with
            # reference examples.
            # MessagesPlaceholder('examples'),
            ("human", "Generate a detailed time entry from the following email: {body}."),
        ]
    )

    runnable = prompt  | llm    
    result =  runnable.invoke({"body": msg_body}).content
 
    return result


def reduce(input_string):

    # Ensure the string is no longer than 12,000 characters
    return input_string[:10000]