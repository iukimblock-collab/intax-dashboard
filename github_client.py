import os
from dotenv import load_dotenv

load_dotenv()

GITHUB_API_KEY = os.getenv("GITHUB_API_KEY")

if not GITHUB_API_KEY:
    raise ValueError(".env 파일에 GITHUB_API_KEY가 없습니다.")

# 사용 예시: GitHub API 요청 시 헤더에 포함
HEADERS = {
    "Authorization": f"token {GITHUB_API_KEY}",
    "Accept": "application/vnd.github.v3+json",
}
