# mbox2pst

Convert mbox mailboxes to PST format — zero dependencies, GUI + CLI, Linux.

## Usage

### GUI mode (requires python3-tk)

```bash
sudo apt install python3-tk
python3 mbox2pst.py
```

### CLI mode (no installation required)

```bash
python3 mbox2pst.py source.mbox output.pst
```

## Output

- `output.pst` — PST file (Outlook compatible)
- `output_emails/` — Individual `.eml` files importable into Outlook or Thunderbird

## Features

- Zero third-party dependencies — uses Python standard library only
- Graphical interface with real-time progress bar and log output
- Falls back to CLI mode automatically when Tkinter is unavailable
- Handles MIME-encoded headers and multipart messages
- Encoding detection with fallback chain: UTF-8 → Latin-1 → CP1252 → ISO-8859-9
