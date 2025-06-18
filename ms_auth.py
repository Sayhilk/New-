from msal import PublicClientApplication

def get_token(client_id, tenant_id):
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["Files.Read.All", "Sites.Read.All"]

    app = PublicClientApplication(client_id, authority=authority)
    result = app.acquire_token_interactive(scopes=scopes)
    return result["access_token"]