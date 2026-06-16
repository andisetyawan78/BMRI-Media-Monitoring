"""
Jalankan script ini SEKALI di komputer lokal untuk generate token.json.
Setelah token.json ada, encode ke base64 dan simpan sebagai GitHub Secret.

Cara pakai:
  python generate_token.py
"""

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

flow  = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
creds = flow.run_local_server(port=0)

with open("token.json", "w") as f:
    f.write(creds.to_json())

print("✅ token.json berhasil dibuat!")
print()
print("Sekarang encode ke base64 dan simpan sebagai GitHub Secret:")
print()
print("  Mac/Linux:")
print("    base64 -i credentials.json | tr -d '\\n'")
print("    base64 -i token.json       | tr -d '\\n'")
print()
print("  Windows PowerShell:")
print("    [Convert]::ToBase64String([IO.File]::ReadAllBytes('credentials.json'))")
print("    [Convert]::ToBase64String([IO.File]::ReadAllBytes('token.json'))")
