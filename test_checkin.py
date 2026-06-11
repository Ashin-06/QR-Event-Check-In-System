import sys, urllib.request, json, time
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
time.sleep(2)

def get(path):
    r = urllib.request.urlopen(f'http://localhost:5001{path}', timeout=8)
    return json.loads(r.read())

def post(path, payload):
    data = json.dumps(payload).encode()
    req = urllib.request.Request(f'http://localhost:5001{path}', data=data,
        headers={'Content-Type': 'application/json'}, method='POST')
    try:
        r = urllib.request.urlopen(req, timeout=10)
        return json.loads(r.read()), r.status
    except urllib.error.HTTPError as e:
        return json.loads(e.read()), e.code

# Registry lookup
reg = get('/registry')
john = next((x for x in reg if x['name'] == 'John Doe'), None)
print(f"Registry: {len(reg)} records, John status: {john['status']}")

# Check-in John Doe
result, code = post('/manual_checkin', {'qr_data': john['qr']})
print(f"Check-in [{code}]: {result.get('message')}")
if code == 200:
    d = result.get('details', {})
    print(f"  Name: {d.get('Name')} | Status: {d.get('Status')} | Count: {d.get('ScanCount')}")
else:
    print(f"  ERROR: {result}")

# Stats after
stats = get('/stats')
print(f"Stats: total={stats['total']} unique={stats['unique']} dup={stats['duplicate']}")

# Registry status after
reg = get('/registry')
john = next((x for x in reg if x['name'] == 'John Doe'), None)
print(f"John status after check-in: {john['status']}")

# Test duplicate scan
result2, code2 = post('/manual_checkin', {'qr_data': john['qr']})
print(f"Duplicate [{code2}]: {result2.get('message')}")
stats2 = get('/stats')
print(f"Stats after dup: total={stats2['total']} unique={stats2['unique']} dup={stats2['duplicate']}")
