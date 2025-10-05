from flask import Flask

app = Flask(__name__)

@app.route('/')
def home():
    return '''
    <h1>Emisyon Saha Uygulaması</h1>
    <p>Sunucu çalışıyor!</p>
    <p>Port: 5001</p>
    '''

if __name__ == '__main__':
    print("Sunucu başlatılıyor...")
    print("http://localhost:5001 adresine gidin")
    app.run(host='0.0.0.0', port=5001, debug=True)





