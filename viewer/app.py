from flask import Flask, render_template, jsonify
import json

app = Flask(__name__)

# Load your JSON data
with open("data.json", "r") as f:
    data = json.load(f)

@app.route("/")
def index():
    return render_template("index.html", documents=data)

if __name__ == "__main__":
    app.run(debug=True)