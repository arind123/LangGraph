# ReAct Agent,  Reasoning and Acting Agent
# Create Tools in Langgraph
# Create a ReAct Graph
# Different Types of Messeges such as ToolMessages, SystemMessages

from typing import Annotated, Sequence, TypedDict
import os
from dotenv import load_dotenv

from langchain_core.messages import BaseMessage # The foundational class for all message types in LangGraph
from langchain_core.messages import ToolMessage # Passes data back to LLM after it calls a tool such as the content and the tool_call_id
from langchain_core.messages import SystemMessage # Message for providing instructions to the LLM
# from langchain_ollama import ChatOllama
from langchain_openai import ChatOpenAI
from langchain_core.tools import tool
from langgraph.graph.message import add_messages #Reducer function
from langgraph.graph import StateGraph, END
from langgraph.prebuilt import ToolNode
from pydantic import BaseModel, Field
import math


load_dotenv()

class AgentState(BaseModel):
    messages: Annotated[Sequence[BaseMessage], add_messages]

@tool
def add(numbers: list[float|int]) -> float|int:
    '''This is a tool to perform addition on a list of numbers.'''
    return sum(numbers)

@tool
def multiply(numbers: list[float|int]) -> float|int:
    '''This is a tool to multiply a list of numbers.'''
    result = 1
    for num in numbers:
        result *= num
    return result

@tool
def subtract(a: float|int , b: float|int) -> float|int:
    '''Function to subtract two inputs.'''
    return a-b

@tool
def division(a: float|int, b: float|int) -> float|int:
    '''Function to perform divison'''
    return a/b

tools = [add, multiply, subtract, division]
# model = ChatOllama(model="mistral").bind_tools(tools)
model = ChatOpenAI(model = "gpt-4o").bind_tools(tools)



def model_call(state: AgentState) -> AgentState:
    '''Call the llm model with tools and system prompt.'''
    system_prompt = SystemMessage(content = "You are my AI assistant, please answer my query to the best of your ability.")                          
    response = model.invoke([system_prompt] + state.messages)
    return {'messages': [response]}

def should_continue(state: AgentState):
    messages = state.messages
    last_message = messages[-1]
    if not last_message.tool_calls:
        return "end"
    else:
        return "continue"

graph = StateGraph(AgentState)
graph.add_node("our_agent", model_call)

tool_node = ToolNode(tools = tools)
graph.add_node("tools", tool_node)

graph.set_entry_point("our_agent")
graph.add_conditional_edges(
    "our_agent",
    should_continue,
    {
        "continue" : "tools",
        "end": END
    }
)

graph.add_edge("tools", "our_agent")

app = graph.compile()

def print_stream(stream):
    for s in stream:
        message = s["messages"][-1]
        if isinstance(message, tuple):
            print(message)
        else:
            message.pretty_print()


inputs = {"messages": [("user", "Add 1 to 10 and then divid by 100.")]}
print_stream(app.stream(inputs, stream_mode="values"))