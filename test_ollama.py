import json
import urllib.request

PROMPT_PATH = "out/prompt_mini.txt"
OUTPUT_PATH = "out/ollama_mini_exaone2.txt"
URL = "http://localhost:11434/api/generate"
MODEL = "exaone3.5:7.8b"


def main():
    with open(PROMPT_PATH, "r", encoding="utf-8") as f:
        prompt = f.read()

    payload = json.dumps({
        "model": MODEL,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0,
        },
    }).encode("utf-8")

    req = urllib.request.Request(
        URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")

    result = json.loads(body)
    response_text = result["response"]

    print(response_text[:300])

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(response_text)


if __name__ == "__main__":
    main()
