from pathlib import Path

from langchain_core.messages import SystemMessage, HumanMessage
from langchain_ollama import ChatOllama

import excelipy as ep

if __name__ == "__main__":
    base_model = ChatOllama(model="qwen2.5:7b")
    sys_prompt = SystemMessage(ep.AI_GUIDE)
    human_prompt = HumanMessage("Create a mocked social platform interaction table and style it")
    ai_table = ep.Table.model_validate((
        base_model
        .with_structured_output(schema=ep.Table.model_json_schema())
        .invoke([sys_prompt, human_prompt])
    ))
    ep.save(ep.Excel(
        path=Path("ai_output.xlsx"),
        sheets=[ep.Sheet(
            name="Sales",
            components=[ai_table],
        )],
    ))
