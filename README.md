msgraph2
========

This is a small(-ish) Python library for wrapping interactions
with Microsoft's Graph API, specifically with respect to
populating and interrogating SharePoint sites safely.

Why `msgraph2`?  Because PyPi.org doesn't allow the package name
`msgraph` because it is too similar to other (unspecified)
packages also on the index.  It was okay with msgraph2, however...

Examples
--------

```python
import msgraph2

sp = msgraph2.SafeSharePoint(
  host='yours.sharepoint.com',
  site='SITE-NAME',
  library='A Document Library',
  token=msgraph2.file_token('path/to/token.json')
)

sp.mkdir("/Incoming/Uploaded Documents", make_parents=True)
```

Where `/path/to/token.json` looks something like this:

```json
{
  "access_token": "... your access token ..."
}
```

The `msgraph2.file_token` function causes the token to be re-read
from the file every time it is needed.  Other keys in the token
JSON file will be explicitly ignored, so if you have a system of
refreshing access tokens that rewrites the on-disk file every
refresh, everything Just Works(TM).

If you are running a copy of [Oauth-Taker][1], you can point
msgraph2 there with the `msgraph2.oauth_taker_token()` helper
instead:

[1]: https://github.com/jhunt/oauth-taker

```python
import msgraph2

sp = msgraph2.SafeSharePoint(
  host='yours.sharepoint.com',
  site='SITE-NAME',
  library='A Document Library',
  token=msgraph2.oauth_taker_token(
    endpoint='https://ot.example.com/t/handler/t0',
    shared_key='my-sekrit-key-for-getting-tokens'
  )
)

sp.mkdir("/Incoming/Uploaded Documents", make_parents=True)
```
