from flask import Flask, jsonify, request
from flask_cors import CORS
import random
import requests # requests 라이브러리를 임포트합니다.

app = Flask(__name__)
CORS(app) # 모든 경로에 대해 CORS를 허용합니다.

@app.route('/api/random')
def random_number():
    """Generates a random integer between 1 and 99 and returns it as JSON."""
    number = random.randint(1, 99)
    return jsonify({'number': number})

# --- 새로운 날씨 API 엔드포인트 ---
@app.route('/api/weather')
def get_weather():
    """
    Fetches weather data from the Open-Meteo API based on query parameters.
    Required query parameters: lat, lon, date
    Example: /api/weather?lat=37.5665&lon=126.9780&date=2025-09-27
    """
    # 1. 클라이언트로부터 위도(lat), 경도(lon), 날짜(date) 파라미터를 받습니다.
    latitude = request.args.get('lat')
    longitude = request.args.get('lon')
    target_date = request.args.get('date')

    # 2. 필수 파라미터가 모두 있는지 확인합니다.
    if not all([latitude, longitude, target_date]):
        return jsonify({'error': 'Missing required query parameters: lat, lon, date'}), 400

    # 3. Open-Meteo API URL을 구성합니다.
    # 골프 대시보드에서 사용하던 정보(최고기온, 강수확률, 풍속, 시간별 강수량)를 요청합니다.
    api_url = (
        f"https://api.open-meteo.com/v1/forecast"
        f"?latitude={latitude}&longitude={longitude}"
        f"&daily=weathercode,temperature_2m_max,precipitation_probability_max,windspeed_10m_max"
        f"&hourly=precipitation"
        f"&timezone=Asia/Seoul"
        f"&start_date={target_date}&end_date={target_date}"
    )

    try:
        # 4. requests 라이브러리를 사용해 Open-Meteo API를 호출합니다.
        response = requests.get(api_url)
        response.raise_for_status()  # 200번대 상태 코드가 아닐 경우 예외를 발생시킵니다.

        # 5. 성공적으로 받은 날씨 데이터를 JSON 형태로 클라이언트에 반환합니다.
        weather_data = response.json()
        return jsonify(weather_data)

    except requests.exceptions.RequestException as e:
        # 6. API 호출 중 네트워크 오류나 HTTP 오류가 발생하면 에러 메시지를 반환합니다.
        return jsonify({'error': f'Failed to fetch weather data: {e}'}), 500


if __name__ == '__main__':
    # Render.com 같은 서비스에서는 gunicorn을 사용하므로, 
    # 이 부분은 주로 로컬 테스트용으로 사용됩니다.
    app.run(host='0.0.0.0', port=5000)
