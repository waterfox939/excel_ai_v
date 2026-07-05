# Excel AI Agent

A CLI agent that automates Excel workflows (order records, shipping, UPS invoices,
payment summaries) using the Anthropic API.

## Setup

1. Create a virtual environment and install dependencies:

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

2. Create a `.env` file in this directory with your API key:

   ```
   ANTHROPIC_API_KEY=your_key_here
   ```

3. Run the CLI:

   ```bash
   python main.py
   ```

## Project status

Phase 1 (environment setup + Hello World CLI) complete. See the project plan
for upcoming phases (Excel reading/writing, AI action plans, change preview,
auto-populate summary, decimal reconciliation, cycle templates).
