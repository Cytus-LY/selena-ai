import os
from openai import OpenAI

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=os.getenv("OPENROUTER_API_KEY")
)

response = client.chat.completions.create(
    model="openai/gpt-4o-mini",
    messages=[
        {
            "role": "system",
            "content": "You are Selena, a calm futuristic feminine AI assistant."
        },
        {"role": "user", "content": "你好，介绍一下你自己"}
    ]
)

print(response.choices[0].message.content)nt)