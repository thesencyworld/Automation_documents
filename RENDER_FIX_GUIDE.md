# Render PDF conversion fix

## 1. Main Flask service
Deploy this project to the existing `Automaton_Documents_replace` web service.
The included `Procfile` sets Gunicorn's worker timeout to 300 seconds so it does
not kill a request while DOCX files are being converted.

If the Render dashboard has a custom **Start Command**, replace it with:

```bash
gunicorn app:app --workers 1 --threads 4 --timeout 300 --graceful-timeout 30 --keep-alive 5
```

## 2. Separate Gotenberg PDF service (required)
The Flask app does not convert DOCX to PDF locally. It sends each DOCX to a
separate Gotenberg service at:

`GOTENBERG_URL/forms/libreoffice/convert`

In Render, open the **main Flask service** → **Environment** and set:
- Key: `GOTENBERG_URL`
- Value: the base public URL of the separate Gotenberg Docker service, e.g.
  `https://your-actual-gotenberg-service.onrender.com`
  Do not add `/forms/libreoffice/convert` to the value and do not use the example
  URL literally. Do not include a trailing slash.

The Gotenberg service must be deployed from this repository's `pdf-service`
folder using its Dockerfile and must be running. Open its base URL in a browser
or test `/health` if supported by that Gotenberg version. If that service returns
502 or is unavailable, the main Flask code alone cannot make conversion succeed.
Check the Gotenberg service's Render logs and make sure the URL is its current
public URL.

## 3. Redeploy and verify
1. Push these files to the Git repository connected to Render, or upload the
   updated project to the repository.
2. Trigger **Manual Deploy → Deploy latest commit** for the Flask service.
3. Confirm `https://automaton-documents-replace.onrender.com/health` returns
   JSON with `status: ok` and `pdf_converter_configured: true`.
4. Generate one document. If it still fails, inspect the *Gotenberg service*
   logs at the matching timestamp; the error is upstream from Flask.

## Important
Retries help only with temporary upstream errors. A persistent HTTP 502 means
the Gotenberg service, its proxy, or its configured URL still needs attention.
