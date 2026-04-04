# Adding Rules

Editorial rules are defined **server-side** in the Mastra workflow, not in the frontend. The frontend add-in has no local analysis logic — it sends text to the backend and receives suggestions.

## How It Works

1. The Word add-in reads the document text.
2. The text is chunked and sent to the Mastra `stylistic-workflow`.
3. The workflow (AI-powered) analyzes the text and returns suggestions with severity levels.
4. The add-in applies suggestions as Track Changes with justification comments.

## Adding New Rules

To add or modify editorial rules, update the **backend Mastra workflow**:

1. Modify the AI agent's system instructions in the workflow definition.
2. Add new pattern categories (e.g., "Passive Voice", "Gender-inclusive language").
3. Adjust the model's prompt to detect new patterns.
4. Assign appropriate severity levels (`"high"`, `"medium"`, `"low"`) for each type of suggestion.

The frontend does **not** need to change when rules are added or modified. The only requirement is that the workflow output conforms to the [API contract](api-contract.md).
