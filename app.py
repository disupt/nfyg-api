from flask import Flask, jsonify
from flask_cors import CORS  # CORS를 위해 추가
import random

# Flask 애플리케이션을 생성합니다.
app = Flask(__name__)

# CORS(app)을 추가하여 모든 출처에서의 API 요청을 허용합니다.
# 이렇게 하면 Gemini Canvas와 같은 다른 도메인에서도 API를 호출할 수 있습니다.
CORS(app)

# '/api/random' URL 경로에 대한 함수를 정의합니다.
# 이 경로로 GET 요청이 오면 이 함수가 실행됩니다.
@app.route('/api/random')
def random_number():
    """1부터 99 사이의 랜덤한 정수를 생성하고 JSON 형태로 반환합니다."""
    number = random.randint(1, 99)
    # 생성된 숫자를 'number'라는 키와 함께 딕셔너리로 만듭니다.
    # jsonify는 딕셔너리를 JSON 응답으로 변환해줍니다.
    return jsonify({'number': number})

# 이 스크립트가 직접 실행될 때 웹 서버를 구동합니다.
# render.com과 같은 서비스는 이 부분을 사용하여 앱을 실행합니다.
if __name__ == '__main__':
    # host='0.0.0.0'은 외부에서 접속 가능하도록 설정합니다.
    app.run(host='0.0.0.0', port=5000)
