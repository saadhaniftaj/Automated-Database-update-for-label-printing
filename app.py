import os
import re
import tempfile
from flask import Flask, request, send_file, jsonify, make_response
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Change for production
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB upload limit

DATABASE_FILENAME = 'DATABASE 4.23.25.xlsx'
LOAD_PATTERN = r'^Load # 0*(\d+) Nissi\.xlsx$'

# The HTML/CSS/JS interface as provided, adapted for Excel files
UPLOAD_PAGE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1.0">
    <title>Nissi Fulfillment | Database Update</title>
    <meta name="description" content="Nissi Fulfillment: Upload your database and load files to generate an updated label printing database.">
    <meta name="keywords" content="Nissi Fulfillment, database update, Excel, label printing, automation, operational support">
    <meta name="author" content="Nissi Fulfillment">
    <meta name="robots" content="index, follow">
    <link href="https://fonts.googleapis.com/css?family=Inter:400,700,900&display=swap" rel="stylesheet">
    <style>
        :root {
            --navy: #0a1a3c;
            --offwhite: #f8f9fa;
            --blue-particle: #3a7bd5;
            --navy-dark: #07122a;
            --navy-light: #223366;
            --bg-gradient: linear-gradient(135deg, #e3eefd 0%, #b3c6e7 100%);
        }
        html { scroll-behavior: smooth; }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0; background: var(--bg-gradient); color: var(--navy);
            min-height: 100vh;
            overflow-x: hidden;
            position: relative;
        }
        .bg-animation { position: fixed; z-index: 0; top: 0; left: 0; width:100vw; height:100vh; pointer-events:none; }
        .particle { position: absolute; width: 4px; height: 4px; border-radius: 50%; opacity: 0.85; background: var(--blue-particle); animation: float 16s linear infinite; }
        @keyframes float {
            0% { transform: translateY(0); opacity: 0; }
            10% { opacity: 1; }
            90% { opacity: 1; }
            100% { transform: translateY(-110vh); opacity: 0; }
        }
        nav {
            position: fixed; top: 0; left: 0; width: 100%; z-index: 100;
            height: 62px; background: var(--navy);
            display: flex; align-items: center; box-shadow: 0 1px 0 #1112;
            border-bottom: 1px solid #223366;
        }
        .nav-content {
            width: 100%; max-width: 1200px; margin: 0 auto; display: flex; justify-content: space-between; align-items: center; padding: 0 2rem;
        }
        .logo {
            display: flex; align-items: center; font-weight: 900; font-size: 1.3rem; letter-spacing: -.03em;
            color: var(--offwhite); text-decoration: none;
        }
        .logo img { height: 38px; width: 38px; margin-right: 10px; border-radius: 8px; box-shadow: 0 0 12px #0a1a3c11; }
        .nav-links { display: flex; gap: 2rem; }
        .nav-links a {
            color: var(--offwhite); text-decoration: none; font-weight: 600; font-size: 1rem;
            padding: 6px 3px; border-radius: 3px; transition: color 0.17s, background 0.18s;
        }
        .nav-links a:hover, .nav-links a.active {
            color: var(--navy-light); background: rgba(10,26,60,0.13);
        }
        .hero {
            min-height: 40vh; display: flex; flex-direction: column; align-items: center; justify-content: flex-start;
            text-align: center; z-index: 1; position: relative; padding-top: 90px; background: transparent;
        }
        .hero h1 {
            font-size: clamp(2.2rem, 5vw, 3.2rem);
            font-weight: 900;
            letter-spacing: -1.2px;
            line-height: 1.09;
            margin-bottom: 1.2rem;
            background: linear-gradient(90deg, #3a7bd5, #6a82fb, #b3c6e7, #3a7bd5);
            background-size: 150% auto;
            color: #fff;
            background-clip: text;
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            animation: rainbow-shine 6s linear infinite;
            text-shadow: 0 0 6px #3a7bd555, 0 0 2px #fff4;
        }
        @keyframes rainbow-shine {
            0% { background-position: 0% 50%; }
            100% { background-position: 150% 50%; }
        }
        .company {
            color: var(--navy-light); font-size: 2.1rem; font-weight: 900; margin-bottom: 0.7rem;
        }
        .upload-section {
            margin-top: 2.5rem; background: var(--navy-light); border-radius: 14px; padding: 2.2rem 2rem 2.5rem 2rem; box-shadow: 0 2px 18px #0a1a3c13; max-width: 480px; margin-left: auto; margin-right: auto; z-index: 2;
        }
        .upload-section label {
            font-size: 1.13rem; color: var(--offwhite); font-weight: 600; margin-bottom: 1.1rem; display: block;
        }
        .file-input {
            margin-bottom: 25px;
        }
        .file-input input[type="file"] {
            display: block; width: 100%; padding: 12px; border-radius: 7px; border: 1px solid #dbeafe; background: #fff; color: #0a1a3c; font-size: 1.08rem;
        }
        .upload-btn {
            background: linear-gradient(90deg, #3a7bd5, #6a82fb, #b3c6e7, #3a7bd5);
            background-size: 200% auto;
            color: #fff;
            font-weight: 800;
            font-size: 1.13rem;
            border-radius: 50px;
            padding: 1em 2.6em;
            border: none;
            cursor: pointer;
            box-shadow: 0 4px 18px #3a7bd555, 0 2px 8px #19141419;
            transition: background 0.19s, color 0.19s, transform 0.15s;
            text-decoration: none;
            margin-top: 1.1rem;
            animation: rainbow-shine-btn 4s linear infinite;
            text-shadow: 0 0 6px #3a7bd555, 0 0 2px #fff4;
        }
        .upload-btn:hover {
            background: linear-gradient(90deg, #6a82fb, #3a7bd5, #b3c6e7, #6a82fb);
            color: #fff;
            transform: scale(1.04) translateY(-2px);
        }
        @keyframes rainbow-shine-btn {
            0% { background-position: 0% 50%; }
            100% { background-position: 200% 50%; }
        }
        .status-section {
            display: none; margin-top: 20px; padding: 20px; border: 1px solid #ddd; border-radius: 5px; background: #fff; color: var(--navy);
        }
        .status-section.processing { background: #e3e8f7; color: #0a1a3c; }
        .status-section.success { background: #e6f9ed; color: #223366; }
        .status-section.error { background: #f8d7da; color: #721c24; }
        .mission {
            background: var(--navy-light); color: var(--offwhite); border-radius: 10px; padding: 1.2rem 1.5rem; margin: 2.5rem auto 0 auto; max-width: 600px; box-shadow: 0 2px 12px #0a1a3c11; text-align: center;
        }
        .mission-title {
            color: var(--offwhite); font-size: 1.2rem; font-weight: 700; margin-bottom: 0.3rem;
        }
        .mission-text {
            font-size: 1.08rem; line-height: 1.6;
        }
        @media (max-width: 700px) {
            .mission { padding: 1rem 0.7rem; }
            .upload-section { padding: 1.2rem 0.5rem 1.5rem 0.5rem; }
        }
    </style>
</head>
<body>
    <div class="bg-animation" id="bgParticles"></div>
    <nav>
        <div class="nav-content">
            <a class="logo" href="#home">
                <img src="https://placehold.co/40x40/0a1a3c/fff?text=N" alt="Nissi Logo"/>
                Nissi Fulfillment
            </a>
            <div class="nav-links">
                <a href="#home" class="active">Home</a>
            </div>
        </div>
    </nav>
    <section class="hero">
        <div class="company">Nissi Fulfillment</div>
        <h1>Database Update &amp; Label Preparation</h1>
        <div class="upload-section">
            <form id="uploadForm" enctype="multipart/form-data">
                <label for="dbInput">Upload your <b>DATABASE 4.23.25.xlsx</b> and a <b>Load # [number] Nissi.xlsx</b> file below. The system will process and merge them, then generate an updated database file for label printing.</label>
                <div class="file-input">
                    <input type="file" name="database_file" id="dbInput" accept=".xlsx" required>
                </div>
                <div class="file-input">
                    <input type="file" name="load_file" id="loadInput" accept=".xlsx" required>
                </div>
                <button type="submit" class="upload-btn">Upload &amp; Update Database</button>
            </form>
            <div id="statusSection" class="status-section">
                <div id="statusDetails"></div>
            </div>
        </div>
    </section>
    <script>
        // Blue moving particles
        const $bg=document.getElementById('bgParticles'); 
        for(let i=0;i<180;i++){
            const d = document.createElement('div');
            d.className='particle';
            d.style.left=Math.random()*100+'vw';
            d.style.top=Math.random()*100+'vh';
            d.style.animationDuration=(8+Math.random()*12)+'s';
            d.style.opacity = Math.random() * 0.3 + 0.7;
            $bg.appendChild(d);
        }
        document.getElementById('uploadForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const formData = new FormData(e.target);
            const statusSection = document.getElementById('statusSection');
            const statusDetails = document.getElementById('statusDetails');
            statusSection.style.display = 'block';
            statusSection.className = 'status-section processing';
            statusDetails.innerHTML = 'Processing your files...';
            try {
                const response = await fetch('/', {
                    method: 'POST',
                    body: formData
                });
                if (response.ok) {
                    const blob = await response.blob();
                    statusSection.className = 'status-section success';
                    statusDetails.innerHTML = 'Update complete! Your download will start automatically.';
                    // Download the file as DATABASE 4.23.25.xlsx
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'DATABASE 4.23.25.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                    setTimeout(() => window.URL.revokeObjectURL(url), 2000);
                } else {
                    let errorMsg = 'An unknown error occurred.';
                    try {
                        const errorData = await response.json();
                        errorMsg = errorData.detail || errorMsg;
                    } catch {}
                    statusSection.className = 'status-section error';
                    statusDetails.innerHTML = `Error: ${errorMsg}`;
                }
            } catch (error) {
                statusSection.className = 'status-section error';
                statusDetails.innerHTML = `An unexpected error occurred: ${error.message}`;
            }
        });
    </script>
</body>
</html>
'''

def validate_filenames(database_filename, load_filename):
    if database_filename != DATABASE_FILENAME:
        return False, 'Database file must be named "DATABASE 4.23.25.xlsx".'
    match = re.match(LOAD_PATTERN, load_filename)
    if not match:
        return False, 'Load file must match pattern "Load # [number] Nissi.xlsx".'
    return True, match.group(1)

def process_files(database_fp, load_fp, load_number):
    db_df = pd.read_excel(database_fp, dtype=str)
    load_df = pd.read_excel(load_fp, dtype=str)
    load_df = load_df.fillna('')
    expected_cols = list('ABCDEFGHIJ')
    if len(load_df.columns) < 10:
        raise ValueError('Load file must have at least 10 columns (A-J).')
    load_df = load_df.iloc[:, :10]
    load_df.columns = expected_cols
    load_df['E'] = load_df['E'].str[:13].str.replace(' ', '', regex=False)
    load_df['G'] = load_df['G'].apply(lambda x: x[:3] + ' ' + x[3:] if len(x) > 3 else x)
    load_number_stripped = str(int(load_number))
    load_df['I'] = load_df.apply(lambda row: f"{load_number_stripped}_{row['B']}", axis=1)
    def adjust_shipped(val):
        try:
            v = int(float(val))
            return str(v + 2) if v < 50 else str(v + 5)
        except:
            return val
    load_df['J'] = load_df['J'].apply(adjust_shipped)
    new_rows = []
    for _, row in load_df.iterrows():
        new_row = {
            'A': 'x',
            'B': row['B'],
            'C': 'x',
            'D': row['B'],
            'E': 'x',
            'F': 'x',
            'G': 'x',
            'H': 'x',
            'I': row['I'],
            'J': '1',
        }
        new_rows.append(new_row)
    processed_rows = []
    for i, row in load_df.iterrows():
        processed_rows.append(row)
        processed_rows.append(pd.Series(new_rows[i]))
    processed_df = pd.DataFrame(processed_rows).reset_index(drop=True)
    shifted = pd.DataFrame('', index=processed_df.index, columns=db_df.columns)
    for i, col in enumerate(processed_df.columns):
        if i + 2 < len(shifted.columns):
            shifted.iloc[:, i + 2] = processed_df[col]
    # --- Updated Column A assignment logic ---
    # x = number of original database rows
    # y = number of processed Load rows (after new row insertion)
    x = len(db_df)
    y = len(shifted) // 2  # since every original row gets a new row after it
    # Generate pattern: x+1, x+y+1, x+2, x+y+2, ...
    a_vals = []
    for i in range(y):
        a_vals.append(x + 1 + i)      # for original row from Load
        a_vals.append(x + y + 1 + i)  # for inserted new row
    # If odd number of rows, handle last one
    if len(a_vals) < len(shifted):
        a_vals.append(x + 1 + y)
    shifted.iloc[:, 0] = a_vals[:len(shifted)]
    # --- Column B: increment by 1, repeat once (existing logic) ---
    last_b = db_df.iloc[-2:, 1].tolist() if len(db_df) >= 2 else [0, 0]
    try:
        b_start = int(last_b[-1])
    except:
        b_start = len(db_df)
    b_vals = []
    b_val = b_start + 1
    for _ in range(len(shifted) // 2):
        b_vals.extend([b_val, b_val])
        b_val += 1
    if len(b_vals) < len(shifted):
        b_vals.append(b_val)
    shifted.iloc[:, 1] = b_vals[:len(shifted)]
    updated_db = pd.concat([db_df, shifted], ignore_index=True)
    return updated_db

@app.route('/', methods=['GET'])
def index():
    return UPLOAD_PAGE

@app.route('/', methods=['POST'])
def handle_upload():
    if 'database_file' not in request.files or 'load_file' not in request.files:
        return jsonify({'detail': 'Both files are required.'}), 400
    db_file = request.files['database_file']
    load_file = request.files['load_file']
    valid, result = validate_filenames(db_file.filename, load_file.filename)
    if not valid:
        return jsonify({'detail': result}), 400
    load_number = result
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as db_tmp, \
             tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as load_tmp:
            db_file.save(db_tmp.name)
            load_file.save(load_tmp.name)
            updated_db = process_files(db_tmp.name, load_tmp.name, load_number)
            output_fp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
            updated_db.to_excel(output_fp.name, index=False)
            output_fp.close()
        response = make_response(send_file(output_fp.name, as_attachment=True, download_name=DATABASE_FILENAME))
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
    except Exception as e:
        return jsonify({'detail': f'Error processing files: {e}'}), 500

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port) 