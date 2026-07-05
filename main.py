"""Entry point, CLI loop."""
from agent import send_message


def main() -> None:
    print("Excel AI Agent — type a message (or 'exit' to quit)")
    while True:
        try:
            user_input = input("\n> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nGoodbye.")
            break

        if not user_input:
            continue
        if user_input.lower() in ("exit", "quit"):
            print("Goodbye.")
            break

        try:
            reply = send_message(user_input)
            print(reply)
        except Exception as e:
            print(f"Error talking to Claude: {e}")


if __name__ == "__main__":
    main()
