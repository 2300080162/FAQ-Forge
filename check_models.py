from google import genai

# 🔑 API Key (hardcoded as default)
API_KEY = "AIzaSyCXnd1czbPj2pS2M1fws67j55BjtkMK16g"

client = genai.Client(api_key=API_KEY)

print("=" * 60)
print("  GEMINI AVAILABLE MODELS")
print("=" * 60)

models = list(client.models.list())

for i, model in enumerate(models, 1):
    print(f"\n[{i}] {model.name}")

    # Show display name if available
    if hasattr(model, "display_name") and model.display_name:
        print(f"    Display Name  : {model.display_name}")

    # Show description if available
    if hasattr(model, "description") and model.description:
        print(f"    Description   : {model.description}")

    # Show supported actions/methods
    if hasattr(model, "supported_actions") and model.supported_actions:
        print(f"    Actions       : {', '.join(model.supported_actions)}")

    # Show input/output token limits
    if hasattr(model, "input_token_limit") and model.input_token_limit:
        print(f"    Input Tokens  : {model.input_token_limit:,}")
    if hasattr(model, "output_token_limit") and model.output_token_limit:
        print(f"    Output Tokens : {model.output_token_limit:,}")

print("\n" + "=" * 60)
print(f"  Total models found: {len(models)}")
print("=" * 60)