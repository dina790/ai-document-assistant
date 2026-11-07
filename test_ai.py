from pathlib import Path

from dotenv import dotenv_values
from openai import OpenAI, APIConnectionError, APIStatusError

settings = dotenv_values(
    Path(__file__).with_name(".env"),
    encoding="utf-8-sig",
)

key = settings.get("OPENAI_API_KEY")

if not key:
    print("API key missing from .env")
    raise SystemExit(1)

try:
    with OpenAI(
        api_key=key,
        timeout=30.0,
        max_retries=0,
    ) as client:
        response = client.responses.create(
            model="gpt-4.1-mini",
            input="Reply with only: AI connection successful",
            max_output_tokens=30,
            store=False,
        )

    print(response.output_text)

except APIConnectionError:
    print("Connection failed. Check your internet connection.")

except APIStatusError as error:
    print("API request failed.")
    print("HTTP status:", error.status_code)
    print("Error code:", error.code)