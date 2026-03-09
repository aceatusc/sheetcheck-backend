from utils import parse_segments
from utils import _call_mistralai, _call_google

user_prompt = (
        f"Worksheet context:\n```json\n[]\n```\n\n"
        f"User request: I want to create a tour plan to Washington DC museums in this sheet."
    )

if __name__ == "__main__":
    for call in [_call_google, _call_mistralai]:
        print("MODEL:", call.__name__)
        raw_out = call(user_prompt)
        print("RAW:")
        print(raw_out)
        parse_out = parse_segments(raw_out)
        print("Parsed:")
        print(parse_out)
