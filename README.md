msgraph
=======

This is a small(-ish) Python library for wrapping interactions
with Microsoft's Graph API, specifically with respect to
populating and interrogating SharePoint sites safely.

Examples
--------

```python
import msgraph

sp = msgraph.SafeSharePoint(
  host='yours.sharepoint.com',
  site='SITE-NAME',
  library='A Document Library',
  token=msgraph.file_token('path/to/token.json')
)

sp.mkdir("/Incoming/Uploaded Documents", make_parents=True)
```

Where `/path/to/token.json` looks something like this:

```json
{
  "access_token": "... your access token ..."
}
```

The `msgraph.file_token` function causes the token to be re-read
from the file every time it is needed.  Other keys in the token
JSON file will be explicitly ignored, so if you have a system of
refreshing access tokens that rewrites the on-disk file every
refresh, everything Just Works(TM).
