# Vercel deployment

This project is now set up as a Vercel-ready Flask web app.

## What Vercel uses

- `app.py` as the Flask entrypoint
- `requirements.txt` for Python dependencies
- `.python-version` to pin Python 3.12
- `vercel.json` to keep unnecessary local build files out of the deployment bundle

## Deploy

1. Put this folder in a Git repository.
2. Push it to GitHub, GitLab, or Bitbucket.
3. Import the repository into Vercel.
4. Deploy with the default settings.

Vercel's Flask documentation says a Flask app can deploy with zero configuration as long as it exposes a top-level `app`, and that static assets should be placed in `public/**` when needed:

- https://vercel.com/docs/frameworks/backend/flask
- https://vercel.com/docs/functions/runtimes/python

## Local run

```bash
python -m venv .venv
. .venv/Scripts/activate
pip install -r requirements.txt
python app.py
```

Then open:

```text
http://127.0.0.1:5000
```

## Routes

- `/` frontend
- `/api/generate` upload endpoint
- `/health` health check
