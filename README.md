# shibo

Python tools for teaching evaluation and oscilloscope screen capture.

## Files

- `teaching_eval_app.py`: main desktop application
- `ai_scoring.py`: AI scoring logic
- `gds1000e.py`: GDS-1000E oscilloscope communication helpers
- `assets/xjtu_logo.png`: UI asset

## Quick start

Clone the repository, then run the main app:

```bash
python teaching_eval_app.py
```

Useful test scripts:

```bash
python test_screen.py
python test_visa.py
```

## Notes

- `.local_secrets.json` is ignored by Git and will not be uploaded.
- `experiments/` and `captures/` are also ignored by Git.
