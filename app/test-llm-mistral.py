from utils import _call_mistralai, parse_segments

user_prompt = (
        f"Worksheet context:\n```json\n[]\n```\n\n"
        f"User request: I want to create a tour plan to Washington DC museums in this sheet."
    )

if __name__ == "__main__":
    raw_out = _call_mistralai(user_prompt)
    print("RAW:")
    print(raw_out)
    parse_out = parse_segments(raw_out)
    print("Parsed:")
    print(parse_out)
