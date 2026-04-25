import argparse
import json
import re
import traceback


def normalize_markdown_output(text):
    if not text:
        return ""

    text = text.replace("\u00a0", " ")
    bullet_chars = "\u2022\u25cf\u25aa\u25ab\u25e6\u2023\u2043\u2219\u00b7\uf0b7\uf0a7\uf0d8\uf0fc"
    bullet_re = re.compile(r"^(\s*)[" + re.escape(bullet_chars) + r"]\s*(\S.*)$")
    lines = []
    for line in text.splitlines():
        match = bullet_re.match(line)
        if match:
            lines.append("%s- %s" % (match.group(1), match.group(2)))
        else:
            lines.append(line.rstrip())
    return "\n".join(lines).strip() + "\n"


def main():
    parser = argparse.ArgumentParser(description="MarkItDown backend bridge")
    parser.add_argument("--source", required=True, help="File path or URL to convert")
    parser.add_argument("--api-key", help="OpenAI API key for LLM descriptions")
    parser.add_argument("--model", default="gpt-4o-mini", help="OpenAI model name")
    args = parser.parse_args()

    try:
        from markitdown import MarkItDown

        kwargs = {}
        if args.api_key:
            from openai import OpenAI

            kwargs["llm_client"] = OpenAI(api_key=args.api_key)
            kwargs["llm_model"] = args.model

        converter = MarkItDown(**kwargs)
        result = converter.convert(args.source)
        payload = {
            "success": True,
            "text": normalize_markdown_output(result.text_content or ""),
            "error": "",
            "detail": "",
        }
    except Exception as exc:
        payload = {
            "success": False,
            "text": "",
            "error": str(exc),
            "detail": traceback.format_exc(),
        }

    print(json.dumps(payload, ensure_ascii=False))


if __name__ == "__main__":
    main()
