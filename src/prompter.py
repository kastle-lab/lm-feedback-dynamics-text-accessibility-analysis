import os
import sys
import time


if "scripts" not in sys.path:
    sys.path.insert(0, "scripts")

from ollama_interface import chat_with_model


def _get_model_name():
    model_name = os.environ.get("OLLAMA_MODEL")
    if not model_name:
        raise RuntimeError("Set OLLAMA_MODEL before running the pipeline.")
    return model_name


def _get_temperature():
    return float(os.environ.get("OLLAMA_TEMPERATURE", "0"))


def call_model(model_name, prompt, temperature):
    # Use one user message to match the original OpenAI prompt-only call.
    response = chat_with_model(
        model_name=model_name,
        messages=[{"role": "user", "content": prompt}],
        options={"temperature": temperature},
    )
    return response, response["message"]["content"]


def get_model_response(prompt, retries=5, delay=1):
    model_name = _get_model_name()
    temperature = _get_temperature()

    for attempt in range(retries):
        try:
            response, content = call_model(
                model_name=model_name,
                prompt=prompt,
                temperature=temperature,
            )

            return {
                "text": content,
                "attempts": attempt + 1,
                "model": response.get("model", model_name),
            }

        except Exception as e:
            print(f"Model error on attempt {attempt + 1}/{retries}: {e}")

            if attempt == retries - 1:
                break

            time.sleep(delay)
            delay *= 2  # exponential backoff

    print("Exceeded maximum retry attempts.")
    return None


def prompter(chapter, instructions):
    prompt = (
        f"Make the chapter more readable and dyslexia friendly. \n"
        f"{instructions}\n"
        f" Here is the text: \n"
        f"{chapter}"
    )

    print("Sending prompt")
    response = get_model_response(prompt)
    print(type(response), " response type")
    print(response["text"], ' response ')
    return response["text"]
