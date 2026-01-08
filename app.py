from flask import Flask, render_template, request
import pandas as pd
import os
from datetime import datetime
import re

app = Flask(__name__)

FILE_NAME = "form_data.xlsx"

EMAIL_PATTERN = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"

@app.route("/", methods=["GET", "POST"])
def form():
    error = None

    if request.method == "POST":
        name = request.form["name"]
        email = request.form["email"]
        age = request.form["age"]
        feedback = request.form["feedback"]

        # Email format validation
        if not re.match(EMAIL_PATTERN, email):
            error = "Please enter a valid email address."
            return render_template("form.html", error=error)

        # Age validation
        if not age.isdigit() or int(age) < 1 or int(age) > 120:
            error = "Age must be a number between 1 and 120."
            return render_template("form.html", error=error)

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Duplicate email check
        if os.path.exists(FILE_NAME):
            old_df = pd.read_excel(FILE_NAME)

            if email in old_df["Email"].values:
                error = "This email has already been submitted."
                return render_template("form.html", error=error)

        new_data = {
            "Name": [name],
            "Email": [email],
            "Age": [age],
            "Feedback": [feedback],
            "Timestamp": [timestamp]
        }

        new_df = pd.DataFrame(new_data)

        if os.path.exists(FILE_NAME):
            final_df = pd.concat([old_df, new_df], ignore_index=True)
        else:
            final_df = new_df

        final_df.to_excel(FILE_NAME, index=False)

        return render_template("success.html")

    return render_template("form.html", error=error)


if __name__ == "__main__":
    app.run(debug=True)
