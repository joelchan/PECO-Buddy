from flask import Flask, render_template, jsonify
import json

app = Flask(__name__)

# Load your JSON data
with open("data.json", "r") as f:
    data = json.load(f)

@app.route("/")
def index():
    # Assign unique IDs to Population elements
    for doc_id, document in enumerate(data):
        document['id'] = f"doc-{doc_id}"  # Unique ID for the document
        for key, elements in document['peco_elements'].items():
            for elem_id, element in enumerate(elements):
                element['id'] = f"{doc_id}-{key}-{elem_id}"  # Unique ID based on indices
    return render_template("index.html", documents=data)

if __name__ == "__main__":
    app.run(debug=True)
