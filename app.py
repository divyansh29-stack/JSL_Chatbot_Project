from flask import Flask, request, jsonify, render_template
import requests
import openpyxl

app = Flask(__name__)
user_data = {}

def search_gemini_api(query):
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent"
    api_key = "AIzaSyCDdqf40IMnE_PbNgl82Z0zWaZKlvhd8DM"

    headers = {
        'Content-Type': 'application/json',
    }

    data = {
        "contents": [
            {
                "parts": [
                    {"text": f"Give basic troubleshooting steps for the following problem: {query}"}
                ]
            }
        ]
    }

    response = requests.post(f"{url}?key={api_key}", headers=headers, json=data)
    response.raise_for_status()

    result = response.json()
    troubleshooting_steps = result.get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', "No troubleshooting steps found.")
    
    # Replace "Step X" with "→"
    formatted_steps = troubleshooting_steps.replace("Step", "→")
    
    return formatted_steps

def save_to_excel(data):
    file_path = 'user_data.xlsx'
    try:
        workbook = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        workbook.active.append(["Name", "Designation", "Department", "Problem", "Resolution"])

    sheet = workbook.active
    sheet.append(data)
    workbook.save(file_path)

@app.route('/')
def home():
    user_data.clear()
    user_data["step"] = 0
    return render_template('index.html')

@app.route('/chat', methods=['POST'])
def chat():
    user_message = request.json.get("message").strip().lower()
    step = user_data.get("step", 0)

    if step == 0:
        if user_message == "hi":
            user_data["step"] = 1
            return jsonify({"response": "Hello! What is your name?"})
        else:
            return jsonify({"response": "Please type 'Hi' to start the conversation."})
    elif step == 1:
        user_data["name"] = user_message
        user_data["step"] = 2
        return jsonify({"response": "Please enter your designation."})
    elif step == 2:
        user_data["designation"] = user_message
        user_data["step"] = 3
        return jsonify({"response": "Please enter your department."})
    elif step == 3:
        user_data["department"] = user_message
        user_data["step"] = 4
        return jsonify({"response": "Please describe your problem."})
    elif step == 4:
        user_data["problem"] = user_message
        user_data["step"] = 5
        response = search_gemini_api(user_message)
        user_data["last_resolution"] = response
        return jsonify({"response": f"Here are some troubleshooting steps:\n{response}\n\nDid this resolve your issue? (Yes/No)"})
    elif step == 5:
        if user_message == "yes":
            save_to_excel([user_data["name"], user_data["designation"], user_data["department"], user_data["problem"], user_data["last_resolution"]])
            user_data.clear()
            return jsonify({"response": "Great! Your issue has been resolved. Your details have been saved."})
        elif user_message == "no":
            response = search_gemini_api(user_data["problem"])
            user_data["last_resolution"] = response
            user_data["step"] = 6
            return jsonify({"response": f"Try this instead:\n{response}\n\nDid this resolve your issue? (Yes/No)"})
    elif step == 6:
        if user_message == "yes":
            save_to_excel([user_data["name"], user_data["designation"], user_data["department"], user_data["problem"], user_data["last_resolution"]])
            user_data.clear()
            return jsonify({"response": "Great! Your issue has been resolved. Your details have been saved."})
        else:
            save_to_excel([user_data["name"], user_data["designation"], user_data["department"], user_data["problem"], "Not Resolved"])
            user_data.clear()
            return jsonify({"response": "Please contact the IT helpdesk for further assistance."})

if __name__ == '__main__':
    app.run(debug=True)

