import re
import requests
import pandas as pd
from flask import Flask, render_template, request, send_file, after_this_request, redirect, url_for, session
import tempfile, os

app = Flask(__name__)
app.secret_key = "your-secret-key"  # required for session
MUFG_BASE = "https://in.mpms.mufg.com/Initial_Offer/IPO.aspx"


# ----------------------
# Helper: Get Token
# ----------------------
def get_token():
    url = f"{MUFG_BASE}/generateToken"
    resp = requests.post(url, json={}, headers={
        "Content-Type": "application/json; charset=utf-8",
        "X-Requested-With": "XMLHttpRequest"
    }, verify=False)
    data = resp.json()
    return data.get("d")


# ----------------------
# Helper: Get IPO List (Cleaned)
# ----------------------
def fetch_ipos():
    token = get_token()
    if not token:
        return []

    url = f"{MUFG_BASE}/GetDetails"
    resp = requests.post(url, json={"token": token}, headers={
        "Content-Type": "application/json; charset=utf-8",
        "X-Requested-With": "XMLHttpRequest"
    }, verify=False)

    raw = resp.json().get("d")
    if not raw:
        return []

    ids = re.findall(r"<company_id>(\d+)</company_id>", raw)
    names = re.findall(r"<companyname>(.*?)</companyname>", raw, re.S)

    ipos = []
    for i in range(min(len(ids), len(names))):
        ipos.append({"id": ids[i], "name": names[i].strip()})

    return ipos


# ----------------------
# Helper: Check PAN Status
# ----------------------
def check_pan_status(ipo_id, pan, token):
    url = f"{MUFG_BASE}/SearchOnPan"
    payload = {
        "clientid": ipo_id,
        "PAN": pan,
        "IFSC": "",
        "CHKVAL": "1",
        "token": token
    }
    resp = requests.post(url, json=payload, headers={
        "Content-Type": "application/json; charset=utf-8",
        "X-Requested-With": "XMLHttpRequest"
    }, verify=False)

    try:
        raw = resp.json().get("d")
        if not raw:
            return {"PAN": pan, "Status": "No Response"}

        def find(tag):
            m = re.search(rf"<{tag}>(.*?)</{tag}>", raw, re.S)
            return m.group(1).strip() if m else ""

        return {
            "PAN": pan,
            "Name": find("NAME1"),
            "IPO": find("companyname"),
            "AllottedShares": find("ALLOT"),
            "AppliedShares": find("SHARES"),
            "AmountAdjusted": find("AMTADJ"),
            "RefundNo": find("RFNDNO"),
            "RefundAmt": find("RFNDAMT"),
            "Category": find("PEMNDG"),
            "CutOffPrice": find("offer_price")
        }

    except Exception as e:
        return {"PAN": pan, "Status": f"Error {e}"}


# ----------------------
# Routes
# ----------------------
@app.route("/", methods=["GET"])
def index():
    ipo_list = fetch_ipos()
    result = session.pop("single_result", None)
    return render_template("index.html", ipos=ipo_list, single_result=result)


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return "No file uploaded", 400

    file = request.files["file"]
    filename = file.filename.lower()

    # ✅ Support both CSV and Excel uploads
    if filename.endswith(".csv"):
        df = pd.read_csv(file)
    elif filename.endswith((".xls", ".xlsx")):
        df = pd.read_excel(file)
    else:
        return "Unsupported file format. Please upload CSV or Excel.", 400

    # ✅ Normalize column names
    df.columns = [c.strip().upper() for c in df.columns]

    # ✅ Detect PAN column (case/space variations)
    possible_pan_cols = ["PAN NO", "PAN", "PAN_NUMBER", "PANNO"]
    pan_col = None
    for col in possible_pan_cols:
        if col in df.columns:
            pan_col = col
            break

    if not pan_col:
        return f"Error: PAN column not found. Found: {df.columns.tolist()}", 400

    ipo_id = request.form.get("ipo_name")
    token = get_token()

    results = []
    for pan in df[pan_col]:
        res = check_pan_status(ipo_id, str(pan).strip(), token)
        results.append(res)

    result_df = pd.DataFrame(results)

    # ✅ Save as Excel
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    result_df.to_excel(tmp.name, index=False, engine="openpyxl")
    tmp.close()

    @after_this_request
    def cleanup(response):
        try:
            os.remove(tmp.name)
        except Exception as e:
            print("Cleanup error:", e)
        return response

    return send_file(tmp.name, as_attachment=True, download_name="ipo_results.xlsx")


@app.route("/single", methods=["POST"])
def single():
    ipo_id = request.form.get("ipo_name")
    pan = request.form.get("pan")
    if not ipo_id or not pan:
        return "IPO and PAN are required", 400

    token = get_token()
    res = check_pan_status(ipo_id, pan.strip(), token)

    session["single_result"] = res
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)

