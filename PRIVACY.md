# Privacy

- The plugin reads and writes only the Pages, Numbers, Keynote, Excel and CSV files the user selects, locally (on the user's machine or in Claude's workspace).
- It sends no document content or personal data to any server and keeps no copy. Output files are written only where the user asks.
- Passwords given with `--password` are used only to open that one Numbers file in memory. They are never printed, stored or sent anywhere.
- The only network use is the one-time `pip install` of the open-source packages listed in the README.
