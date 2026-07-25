#!/usr/bin/env python3
"""OpenRouter image generation harness for the webinar deck.

Usage:
  python3 genimg.py "prompt text" out.png [--ref path.png ...] [--aspect 16:9]
                    [--model MODEL]

Reads OPENROUTER_API_KEY from the environment or from the scratchpad env file.
Uses the chat/completions endpoint with image output modality; reference
images (brand seeds) are attached as input image parts.
"""
import base64
import json
import os
import sys
import time
import urllib.request

ENV_FILE = ("/tmp/claude-0/-home-user/9d79adee-a1c4-531c-b0d7-1889983ce386/"
            "scratchpad/openrouter.env")
DEFAULT_MODEL = "openai/gpt-5.4-image-2"


def api_key():
    k = os.environ.get("OPENROUTER_API_KEY")
    if k:
        return k
    with open(ENV_FILE) as f:
        for line in f:
            if line.startswith("OPENROUTER_API_KEY="):
                return line.split("=", 1)[1].strip()
    raise SystemExit("no OPENROUTER_API_KEY found")


def data_url(path):
    ext = os.path.splitext(path)[1].lstrip(".").lower() or "png"
    if ext == "jpg":
        ext = "jpeg"
    with open(path, "rb") as f:
        return f"data:image/{ext};base64,{base64.b64encode(f.read()).decode()}"


def generate(prompt, out_path, refs=(), aspect=None, model=DEFAULT_MODEL,
             retries=3):
    content = [{"type": "text", "text": prompt}]
    for r in refs:
        content.append({"type": "image_url",
                        "image_url": {"url": data_url(r)}})
    body = {
        "model": model,
        "messages": [{"role": "user", "content": content}],
        "modalities": ["image", "text"],
    }
    if aspect:
        body["image_config"] = {"aspect_ratio": aspect}
    req = urllib.request.Request(
        "https://openrouter.ai/api/v1/chat/completions",
        data=json.dumps(body).encode(),
        headers={"Authorization": f"Bearer {api_key()}",
                 "Content-Type": "application/json"})
    last_err = None
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=300) as resp:
                data = json.loads(resp.read())
            msg = data["choices"][0]["message"]
            images = msg.get("images") or []
            if not images:
                last_err = f"no image in response: {str(msg)[:400]}"
                time.sleep(3)
                continue
            url = images[0]["image_url"]["url"]
            b64 = url.split(",", 1)[1]
            os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
            with open(out_path, "wb") as f:
                f.write(base64.b64decode(b64))
            return out_path
        except Exception as e:  # noqa: BLE001 - report the API error verbatim
            last_err = repr(e)
            time.sleep(5)
    raise SystemExit(f"generation failed after {retries} tries: {last_err}")


def main():
    args = sys.argv[1:]
    if len(args) < 2:
        raise SystemExit(__doc__)
    prompt, out = args[0], args[1]
    refs, aspect, model = [], None, DEFAULT_MODEL
    i = 2
    while i < len(args):
        if args[i] == "--ref":
            refs.append(args[i + 1]); i += 2
        elif args[i] == "--aspect":
            aspect = args[i + 1]; i += 2
        elif args[i] == "--model":
            model = args[i + 1]; i += 2
        else:
            raise SystemExit(f"unknown arg {args[i]}")
    path = generate(prompt, out, refs=refs, aspect=aspect, model=model)
    print("wrote", path)


if __name__ == "__main__":
    main()
