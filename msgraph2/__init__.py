from requests import request
import json, re, time
import collections.abc
import urllib.parse

def assign(*x):
  o = {}
  for u in x:
    for k, v in u.items():
      if isinstance(v, collections.abc.Mapping):
        o[k] = assign(o.get(k, {}), v)
      else:
        o[k] = v
  return o

class ProcessError(Exception):
  def __init__(self, message, params):
    self.message = message
    self.params  = params

  def to_dict(self):
    return self.params

GRAPH_API_ENDPOINT='https://graph.microsoft.com/v1.0'

def literal_token(token):
  return lambda: token

def file_token(file):
  def reader():
    with open(file, 'r') as io:
      r = json.load(io)
      return r['access_token'] if 'access_token' in r else None
  return reader

def oauth_taker_token(endpoint, shared_key):
  def reader():
    r = request('GET', endpoint, headers={
      'Accept': 'application/json',
      'Authorization': f'API-Key {shared_key}'
    }).json()
    return r['access_token'] if 'access_token' in r else None
  return reader

class API:
  def __init__(self, tokenfn):
    self.tokenfn = tokenfn

  def get(self, endpoint, json_input=None, data=None, content_type='application/json', retries=2):
    return self.call('GET', endpoint, json_input=json_input, data=data, content_type=content_type, retries=retries)

  def put(self, endpoint, json_input=None, data=None, content_type='application/json', retries=2):
    return self.call('PUT', endpoint, json_input=json_input, data=data, content_type=content_type, retries=retries)

  def post(self, endpoint, json_input=None, data=None, content_type='application/json', retries=2):
    return self.call('POST', endpoint, json_input=json_input, data=data, content_type=content_type, retries=retries)

  def patch(self, endpoint, json_input=None, data=None, content_type='application/json', retries=2):
    return self.call('PATCH', endpoint, json_input=json_input, data=data, content_type=content_type, retries=retries)

  def delete(self, endpoint, json_input=None, data=None, content_type='application/json', retries=2):
    return self.call('DELETE', endpoint, json_input=json_input, data=data, content_type=content_type, retries=retries)

  def get_download_url(self, drive_id, item_id, attr='@microsoft.graph.downloadUrl'):
    r = self.get(f'/drives/{drive_id}/items/{item_id}?select=id,{attr}')
    return r.json()[attr]

  def depaginate(self, url, key='value', andthen='@odata.nextLink'):
    r = self.get(url).json()
    values = r[key]
    while andthen in r.keys():
      r = self.get(r[andthen]).json()
      values.extend(r[key])
    return values

  def call(self, method, endpoint, json_input=None, data=None, content_type='application/json', retries=2):
    token = self.tokenfn()
    if not re.match('^' + re.escape(GRAPH_API_ENDPOINT), endpoint):
      endpoint = f'{GRAPH_API_ENDPOINT}{endpoint}'
    print(f'>> {method} {endpoint}')
    headers = {
      'Content-Type': content_type,
      'Authorization': f"Bearer {token}",
    }
    if json_input:
      r = request(method, endpoint, headers=headers, json=json_input)
    elif data:
      r = request(method, endpoint, headers=headers, data=data)
    else:
      r = request(method, endpoint, headers=headers)

    if r.status_code in (200, 201, 204):
      return r

    e = r.json()
    if retries > 0 and 'error' in e and 'code' in e['error']:
      if e['error']['code'] in ['InvalidAuthenticationToken', 'unauthenticated'] and retries > 1:
        return self.call(method, endpoint, json_input=json_input, data=data, content_type=content_type, retries=retries - 1)

    raise ProcessError('bad http response', {
      'status_code': r.status_code,
      'endpoint': endpoint,
      'json_input': json_input,
      'response': r.json()
    })

class SharePoint:
  def __init__(self, host, site, library, token):
    self.host = host
    self.site = site
    self.api = API(token)
    self.n = 0
    self.start = time.clock_gettime(time.CLOCK_MONOTONIC_RAW)

    self.paths = {} # full path name => id
    self.aliases = {} # aliases for object -> column resolution
    self.columns = None # column definitions found on SharePoint list
    self.loaders = {}

    # resolve library name => id
    self.library = None
    r = self.api.get(f'/sites/{self.host}:/sites/{self.site}:/lists')
    for lib in r.json()['value']:
      if lib['name'] == library:
        self.library = lib

    if self.library is None:
      raise Exception(f'library {library} not found')
    self.library_id = self.library['id']

    # determine root drive ID
    r = self.api.get(f'/sites/{self.host}:/sites/{self.site}:/lists/{self.library_id}/drive')
    self.drive = r.json()
    self.drive_id = self.drive['id']

  def clock_start(self):
    self.n = 0
    self.start = time.clock_gettime(time.CLOCK_MONOTONIC_RAW)

  def clock_next(self):
    self.n = self.n + 1
    return self.clock_check()

  def clock_check(self):
    now = time.clock_gettime(time.CLOCK_MONOTONIC_RAW)
    return (self.n, now - self.start)

  def sanitize_file_component(self, s):
    # per https://learn.microsoft.com/en-us/graph/onedrive-addressing-driveitems#onedrive-reserved-characters
    return re.sub(r'[/\\*<>?:|#%]', '_', s)

  def uri_encode(self, s):
    # per https://learn.microsoft.com/en-us/graph/onedrive-addressing-driveitems#uri-path-characters
    return urllib.parse.quote(s, safe='')

  def split_path(self, path):
    if path[:1] == '/':
      path = path[1:]
    return [self.sanitize_file_component(s) for s in path.split('/')]

  def join_path(self, parts):
    return '/' + '/'.join(parts)

  def mkdir(self, path, make_parents=False):
    print(f'>> creating directory "{path}" ...')
    parts = self.split_path(path)
    details = {
      'name': parts[-1],
      'folder': {},
      '@microsoft.graph.conflictBehavior': 'replace'
    }

    if len(parts) == 1: # top-level directory
      r = self.api.post(f'/sites/{self.host}/drives/{self.drive_id}/root/children', json_input=details)
      self.paths[path] = r.json()['id']

    else:
      parent = self.join_path(parts[:-1])
      if make_parents and parent not in self.paths:
        self.mkdir(parent, make_parents=True)

      r = self.api.post(f'/sites/{self.host}/drives/{self.drive_id}/items/{self.paths[parent]}/children', json_input=details)
      self.paths[path] = r.json()['id']

    return self.paths[path]

  def upload(self, local, remote, make_parents=False):
    parts = self.split_path(remote)
    filename = self.uri_encode(parts[-1])

    if len(parts) == 1: # root item
      r = self.api.put(f'/drives/{self.drive_id}/items/root:/{filename}:/content',
        data=open(local, 'rb'),
        content_type='application/octet-stream'
      )

    else: # child item
      parent = self.join_path(parts[:-1])
      if make_parents and parent not in self.paths:
        self.mkdir(parent, make_parents=True)

      r = self.api.put(f'/drives/{self.drive_id}/items/{self.paths[parent]}:/{filename}:/content',
        data=open(local, 'rb'),
        content_type='application/octet-stream'
      )
      return r

  def load(self, src, src_uri, dst_file, attrs={}, make_parents=True):
    if src not in self.loaders:
      raise Exception(f'unhandled {src} source for {src_uri} -> {dst_file}')

    (n, total) = self.clock_next()
    print(f'{n} | {round(total / n, 2)}) {dst_file} ...')
    self.loaders[src](self, src_uri, dst_file, make_parents=make_parents)
    self.annotate(dst_file, attrs)

  def loader(self, source, fn):
    self.loaders[source] = fn

  def list_columns(self, force_reload=False):
    if self.columns is None or force_reload:
      self.columns = {}
      r = self.api.get(f'/sites/{self.host}:/sites/{self.site}:/lists/{self.library_id}/columns')
      for column in r.json()['value']:
        self.columns[column['name']] = column
    return self.columns

  def delete_column(self, name):
    self.list_columns()
    if name in self.columns:
      self.api.delete(f'/sites/{self.host}/columns/{self.columns["name"]["id"]}')

  def create_column(self, name, details):
    self.aliases[name] = name
    column = assign({
      'columnGroup': 'Custom Columns',
      'description': '',
      'displayName': name,
      'name':        name,
      'enforceUniqueValues': False,
      'hidden':              False,
      'indexed':             False,
      'readOnly':            False,
      'required':            False,
    }, details)

    self.list_columns()
    if name in self.columns:
      self.api.patch(f'/sites/{self.host}:/sites/{self.site}:/lists/{self.library_id}/columns/{self.columns[name]["id"]}', json_input=column)
    else:
      self.api.post(f'/sites/{self.host}:/sites/{self.site}:/lists/{self.library_id}/columns', json_input=column)

  def alias(self, key, column):
    self.aliases[key] = column

  def de_alias(self, obj):
    new = {}
    for k, v in obj.items():
      if isinstance(v, list):
        v = ', '.join(v)

      v = v.strip()
      if len(v) > 255:
        v = v[0:252] + '...'

      if k in self.aliases:
        new[self.aliases[k]] = v

    return new

  def annotate(self, rel_path, attrs):
    if rel_path[0:1] != '/':
      rel_path = f'/{rel_path}'

    # NOTE: `:?expand=fields` is not _strictly_ necessary since we only need ['id']
    #r = self.api.get(f'/drives/{self.drive_id}/items/root:{rel_path}:?expand=fields')

    # NOTE: rel_path does NOT get uri_encode()'d, because we WANT the slashes
    item_id = self.api.get(f'/drives/{self.drive_id}/items/root:{rel_path}').json()['id']

    attrs = self.de_alias(attrs)
    # from https://sharepoint.stackexchange.com/questions/307891/how-do-i-get-a-driveitem-as-a-listitem-so-that-i-can-manipulate-the-fields-of-th
    r = self.api.patch(f'/sites/{self.host}/drives/{self.drive_id}/items/{item_id}/listItem/fields', json_input=attrs)
    return r

class SafeSharePoint(SharePoint):
  def mkdir(self, path, make_parents=False):
    try:
      super().mkdir(path, make_parents)
    except Exception as e:
      print(f'mkdir("{path}") failed:')
      print(e)

  def upload(self, local, remote, make_parents=False):
    try:
      super().upload(local, remote, make_parents)
    except Exception as e:
      print(f'upload("{local}", "{remote}") failed:')
      print(e)

  def load(self, src, src_uri, dst_file, attrs={}, make_parents=True):
    try:
      super().load(src, src_uri, dst_file, attrs, make_parents)
    except Exception as e:
      print(f'load(<src>, "{src_uri}", "{dst_file}") failed:')
      print(e)

  def list_columns(self, force_reload=False):
    try:
      super().list_columns(force_reload)
    except Exception as e:
      print(f'list_columns() failed:')
      print(e)

  def delete_column(self, name):
    try:
      super().delete_column(name)
    except Exception as e:
      print(f'delete_column("{name}") failed:')
      print(e)

  def create_column(self, name, details):
    try:
      super().create_column(name, details)
    except Exception as e:
      print(f'create_column("{name}", <details...>) failed:')
      print(e)

  def annotate(self, rel_path, attrs):
    try:
      super().annotate(rel_path, attrs)
    except Exception as e:
      print(f'annotate("{rel_path}", <attrs...>) failed:')
      print(e)
