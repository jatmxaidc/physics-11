import json
import urllib.request
from http.cookiejar import CookieJar


BASE = "http://127.0.0.1:8000"


def client():
    return urllib.request.build_opener(urllib.request.HTTPCookieProcessor(CookieJar()))


def request(opener, path, method="GET", payload=None):
    body = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(
        BASE + path,
        data=body,
        method=method,
        headers={"Content-Type": "application/json"},
    )
    with opener.open(req) as response:
        return json.loads(response.read())


health = request(client(), "/api/health")
assert health["ok"]
assert health["counts"]["lessons"] == 9
assert health["counts"]["questions"] == 45
assert health["counts"]["schedule"] == 138

student = client()
login = request(
    student,
    "/api/login",
    "POST",
    {"email": "student@physics.local", "password": "student123"},
)
assert login["user"]["role"] == "student"
dashboard = request(student, "/api/student/dashboard")
assert len(dashboard["lessons"]) == 9
lesson_id = dashboard["lessons"][1]["id"]
lesson = request(student, f"/api/lessons/{lesson_id}")
assert lesson["title"]
quiz = request(student, f"/api/lessons/{lesson_id}/quiz")
answers = {str(question["id"]): 0 for question in quiz["questions"]}
result = request(
    student,
    f"/api/lessons/{lesson_id}/quiz",
    "POST",
    {"answers": answers, "durationSeconds": 120},
)
assert result["total"] == 5
assert len(result["results"]) == 5

teacher = client()
login = request(
    teacher,
    "/api/login",
    "POST",
    {"email": "teacher@physics.local", "password": "teacher123"},
)
assert login["user"]["role"] == "teacher"
teacher_dashboard = request(teacher, "/api/teacher/dashboard")
assert teacher_dashboard["students"]
student_detail = request(
    teacher,
    f"/api/teacher/student/{teacher_dashboard['students'][0]['id']}",
)
assert student_detail["progress"]

print("Smoke test passed")
print(json.dumps(health, indent=2))
