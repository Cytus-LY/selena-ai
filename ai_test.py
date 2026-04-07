import os
from openai import OpenAI

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=os.getenv("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")
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