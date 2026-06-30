import requests

base_url = "http://localhost:8000"
test_cedula = "94041597"
endpoints = [
    "/",
    "/api/v2/estado-cuenta?cedula=" + test_cedula,
    "/api/v2/estado-cuenta/xlsx?cedula=" + test_cedula,
    "/estado-cuenta?cedula=" + test_cedula,
    "/docs",
    "/openapi.json"
]

for ep in endpoints:
    url = f"{base_url}{ep}"
    try:
        response = requests.get(url, timeout=2)
        print(f"GET {url} - Status: {response.status_code}")
        if response.status_code == 200:
            print(f"Content: {response.text[:100]}...")
    except Exception as e:
        print(f"GET {url} - Error: {e}")
