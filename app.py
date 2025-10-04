from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_file
import os
import pandas as pd
import math
from io import BytesIO
import uuid

app = Flask(__name__)
app.secret_key = "super_secret_key"

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

uploaded_files = {}  # filename -> list of sheets

# 🔹 Global Master File + DataFrame
MASTER_FILE = os.path.join(app.config['UPLOAD_FOLDER'], "master_data.xlsx")
uploaded_data = pd.DataFrame()


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# ----- Login -----
@app.route("/", methods=["GET", "POST"])
def login():
    # If already logged in, redirect to dashboard
    if "email" in session:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")

        # Registered users with email and role
        users = {
            "admin@hinditsolution.com": {"password": "admin123", "role": "admin"},
            "user@site.com": {"password": "user123", "role": "user"},
            "syed@hinditsolution.com": {"password": "syed@542", "role": "user"},
            "salce@hinditsolution.com": {"password": "salce@463", "role": "user"}
        }

        # Authentication
        if email in users and password == users[email]["password"]:
            session["email"] = email
            session["role"] = users[email]["role"]
            return redirect(url_for("dashboard"))
        else:
            flash("Invalid email or password")

    return render_template("login.html")


# ----- Dashboard -----
@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    global uploaded_files, uploaded_data
    if "email" not in session:
        flash("Please login first")
        return redirect(url_for("login"))

    # Handle search, pagination, page_size
    search_query = request.args.get("search", "")
    page = int(request.args.get("page", 1))
    page_size = int(request.args.get("page_size", 10))
    selected_file = request.args.get("selected_file")

    # Refresh uploaded files
    uploaded_files.clear()
    for fname in os.listdir(app.config['UPLOAD_FOLDER']):
        if allowed_file(fname):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], fname)
            try:
                xls = pd.ExcelFile(filepath, engine="openpyxl")
                uploaded_files[fname] = xls.sheet_names
            except Exception as e:
                flash(f"Error reading {fname}: {e}")

    # Upload new file (Admin)
    if session.get("role") == "admin" and request.method == "POST" and "file" in request.files:
        file = request.files["file"]
        if file and allowed_file(file.filename):
            required_columns = ["Date", "Vendor", "CountryISO", "CostPrice", "MCC", "MNC", "Operator"]

            try:
                df_new = pd.read_excel(file, engine="openpyxl")
            except Exception as e:
                flash(f"Error reading file: {e}")
                return redirect(url_for("dashboard"))

            # Schema check
            if not all(col in df_new.columns for col in required_columns):
                missing = [col for col in required_columns if col not in df_new.columns]
                flash(f"Schema mismatch! Missing columns: {', '.join(missing)}")
                return redirect(url_for("dashboard"))

            # Ensure _id column exists
            if "_id" not in df_new.columns:
                df_new["_id"] = [uuid.uuid4().hex for _ in range(len(df_new))]

            # Append or create master file
            if os.path.exists(MASTER_FILE):
                try:
                    df_master = pd.read_excel(MASTER_FILE, engine="openpyxl")
                    if not all(col in df_master.columns for col in required_columns):
                        flash("Existing master file has invalid schema! Please fix manually.")
                        return redirect(url_for("dashboard"))

                    df_master = pd.concat([df_master, df_new], ignore_index=True)
                    df_master.to_excel(MASTER_FILE, index=False)
                    uploaded_data = df_master.copy()
                    flash("Data appended successfully to master file!")
                except Exception as e:
                    flash(f"Error appending to master file: {e}")
                    return redirect(url_for("dashboard"))
            else:
                df_new.to_excel(MASTER_FILE, index=False)
                uploaded_data = df_new.copy()
                flash("Master file created successfully!")

            return redirect(url_for("dashboard", selected_file="master_data.xlsx"))

        else:
            flash("Select a valid .xlsx file")

    # Default file selection
    if not selected_file and uploaded_files:
        selected_file = list(uploaded_files.keys())[0]

    # Load data
    if selected_file:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], selected_file)
        try:
            uploaded_data = pd.read_excel(filepath, engine="openpyxl")
        except Exception as e:
            flash(f"Error loading {selected_file}: {e}")
    total_count_all = len(uploaded_data) if not uploaded_data.empty else 0
    # Ensure `_id` exists and persists
    if not uploaded_data.empty and "_id" not in uploaded_data.columns:
        uploaded_data["_id"] = [uuid.uuid4().hex for _ in range(len(uploaded_data))]
        uploaded_data.to_excel(MASTER_FILE, index=False)

    # Apply search
    df_display = uploaded_data.copy()
    if search_query and not df_display.empty:
        keywords = [s.strip().lower() for s in search_query.split(",") if s.strip()]
        if keywords:
            df_display = df_display[df_display.astype(str).apply(
                lambda row: row.str.lower().str.contains("|".join(keywords)).any(), axis=1
            )]

    # Pagination
    total_pages = max(1, math.ceil(len(df_display) / page_size))
    start = (page - 1) * page_size
    end = start + page_size
    page_data = df_display.iloc[start:end]

    rows = page_data.to_dict(orient="records")
    columns = [c for c in page_data.columns if c != '_id']

    return render_template("dashboard.html",
                           username=session["email"],
                           role=session.get("role"),
                           rows=rows,
                           columns=columns,
                           page=page,
                           total_pages=total_pages,
                           search_query=search_query,
                           page_size=page_size,
                           files_list=list(uploaded_files.keys()),
                           selected_file=selected_file,
                           total_count_all=total_count_all)


# ----- Download Filtered Excel -----
@app.route("/download")
def download():
    if "email" not in session:
        flash("Please login first")
        return redirect(url_for("login"))

    search_query = request.args.get("search", "").strip()
    selected_file = request.args.get("selected_file")

    if not selected_file:
        flash("No file selected")
        return redirect(url_for("dashboard"))

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], selected_file)
    try:
        df_to_download = pd.read_excel(filepath, engine="openpyxl")
    except Exception as e:
        flash(f"Error loading data for download: {e}")
        return redirect(url_for("dashboard"))

    # Apply search filter
    if search_query:
        keywords = [s.strip().lower() for s in search_query.split(",") if s.strip()]
        if keywords:
            df_to_download = df_to_download[df_to_download.astype(str).apply(
                lambda row: row.str.lower().str.contains("|".join(keywords)).any(), axis=1
            )]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_download.to_excel(writer, index=False)
    output.seek(0)

    return send_file(output, download_name=f"filtered_{selected_file}.xlsx", as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ----- Delete Selected Rows -----
@app.route("/delete_rows", methods=["POST"])
def delete_rows():
    global uploaded_data
    if "email" not in session:
        flash("Please login first")
        return redirect(url_for("login"))

    selected_file = request.form.get("selected_file")
    selected_rows = request.form.getlist("delete_checkbox")

    if not selected_file:
        flash("No file selected")
        return redirect(url_for("dashboard"))

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], selected_file)
    try:
        df = pd.read_excel(filepath, engine="openpyxl")
        if '_id' not in df.columns:
            df['_id'] = [uuid.uuid4().hex for _ in range(len(df))]
        df = df[~df['_id'].isin(selected_rows)]
        df.to_excel(filepath, index=False)
        uploaded_data = df.copy()
        flash(f"Selected rows deleted successfully!")
    except Exception as e:
        flash(f"Error deleting rows: {e}")

    return redirect(url_for("dashboard", selected_file=selected_file))


# ----- Update Cell (AJAX) -----
@app.route("/update_cell", methods=["POST"])
def update_cell():
    global uploaded_data
    data = request.get_json()
    row_id = data.get("id")
    column = data.get("column")
    value = data.get("value")

    try:
        row_index = uploaded_data.index[uploaded_data["_id"] == row_id][0]
        uploaded_data.at[row_index, column] = value
        uploaded_data.to_excel(MASTER_FILE, index=False)
        return jsonify({"message": f"Updated row {row_id}, column {column} → {value}"})
    except Exception as e:
        return jsonify({"message": f"Error: {str(e)}"}), 400


# ----- Logout -----
@app.route("/logout")
def logout():
    session.clear()
    flash("Logged out successfully")
    return redirect(url_for("login"))

@app.route("/add_row", methods=["POST"])
def add_row():
    global uploaded_data
    if "email" not in session:
        flash("Please login first")
        return redirect(url_for("login"))

    selected_file = request.form.get("selected_file")
    if not selected_file:
        flash("No file selected")
        return redirect(url_for("dashboard"))

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], selected_file)

    try:
        df = pd.read_excel(filepath, engine="openpyxl")

        # Ensure _id column exists
        if "_id" not in df.columns:
            df["_id"] = [uuid.uuid4().hex for _ in range(len(df))]

        # New row
        new_row = {}
        for col in df.columns:
            if col == "_id":
                new_row[col] = uuid.uuid4().hex
            else:
                new_row[col] = request.form.get(col, "")

        # Append row
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        # Save back
        df.to_excel(filepath, index=False)
        uploaded_data = df.copy()

        flash("✅ Row added successfully!")
    except Exception as e:
        flash(f"❌ Error adding row: {e}")

    return redirect(url_for("dashboard", selected_file=selected_file))


if __name__ == "__main__":
    app.run(debug=True)
