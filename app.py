from flask import Flask, render_template_string, request, send_file
import subprocess
import json
import os
from datetime import datetime
import threading
import time

app = Flask(__name__)
OUTPUT_FOLDER = "output/"

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Lesson Plan Generator</title>
    <style>
        * { box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial, sans-serif; max-width: 650px; margin: 30px auto; padding: 20px; background: #f5f5f5; }
        .container { background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; margin-top: 0; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        label { font-weight: 600; display: block; margin-top: 15px; color: #34495e; }
        input, select { width: 100%; padding: 10px 12px; margin: 5px 0 5px 0; border: 2px solid #ddd; border-radius: 6px; font-size: 14px; transition: border-color 0.3s; }
        input:focus { border-color: #3498db; outline: none; }
        .topic-input { background: #f8f9fa; padding: 10px; border-radius: 6px; margin: 8px 0; }
        button { background: #3498db; color: white; padding: 14px 30px; border: none; border-radius: 6px; cursor: pointer; font-size: 16px; font-weight: 600; width: 100%; margin-top: 20px; transition: background 0.3s; }
        button:hover { background: #2980b9; }
        button:disabled { background: #95a5a6; cursor: not-allowed; }
        .result { margin-top: 20px; padding: 20px; background: #f8f9fa; border-radius: 6px; border-left: 4px solid #3498db; }
        .error { border-left-color: #e74c3c; }
        .success { border-left-color: #27ae60; }
        .loading { text-align: center; padding: 30px; display: none; }
        .loading .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #3498db; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .download-btn { background: #27ae60; padding: 10px 20px; border-radius: 6px; color: white; text-decoration: none; display: inline-block; margin-top: 10px; }
        .download-btn:hover { background: #219a52; }
        .info { font-size: 12px; color: #7f8c8d; margin-top: 3px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📚 Lesson Plan Generator</h1>
        <p style="color: #7f8c8d;">Enter the details below to generate a CBSE lesson plan</p>
        
        <form method="POST" id="lpForm">
            <label>👨‍🏫 Teacher Name:</label>
            <input type="text" name="teacher" value="Teacher" required>
            
            <label>🏫 Class & Section:</label>
            <input type="text" name="class" value="X A" required>
            
            <label>📖 Subject:</label>
            <input type="text" name="subject" value="Biology" required>
            
            <label>📘 Chapter:</label>
            <input type="text" name="chapter" value="Nutrition" required>
            
            <label>📅 Start Date (DD-MM-YYYY):</label>
            <input type="text" name="start_date" value="{{ today }}" required>
            <div class="info">Format: 25-12-2024</div>
            
            <label>⏱️ Duration per period:</label>
            <input type="text" name="duration" value="40 min" required>
            
            <label>📊 Total Periods:</label>
            <input type="number" name="total_periods" value="3" min="1" max="5" id="totalPeriods" required>
            
            <div id="topicsContainer">
                <!-- Topics will be added here -->
            </div>
            
            <button type="submit" id="generateBtn">🚀 Generate Lesson Plan</button>
        </form>
        
        <div id="loading" class="loading">
            <div class="spinner"></div>
            <p style="margin-top: 15px; font-weight: 600;">⏳ Generating your lesson plan... This may take 1-2 minutes.</p>
        </div>
        
        <div id="result" style="display:none;"></div>
    </div>
    
    <script>
        const totalPeriodsInput = document.getElementById('totalPeriods');
        const topicsContainer = document.getElementById('topicsContainer');
        
        function updateTopics() {
            const count = parseInt(totalPeriodsInput.value) || 3;
            topicsContainer.innerHTML = '';
            for (let i = 1; i <= count; i++) {
                const div = document.createElement('div');
                div.className = 'topic-input';
                div.innerHTML = `
                    <label style="margin-top:0;">📌 Topic for Period ${i}:</label>
                    <input type="text" name="topic_p${i}" value="Topic ${i}" required>
                `;
                topicsContainer.appendChild(div);
            }
        }
        
        totalPeriodsInput.addEventListener('change', updateTopics);
        totalPeriodsInput.addEventListener('input', updateTopics);
        updateTopics();
        
        document.getElementById('lpForm').addEventListener('submit', function(e) {
            document.getElementById('loading').style.display = 'block';
            document.getElementById('result').style.display = 'none';
            document.getElementById('generateBtn').disabled = true;
        });
    </script>
</body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            # Get the current script directory
            script_dir = os.path.dirname(os.path.abspath(__file__))
            # Use the correct script name: Ipneww.py or lesson_plan_generator.py
            script_name = "Ipneww.py"  # Change this if you renamed it
            
            # Check if script exists
            script_path = os.path.join(script_dir, script_name)
            if not os.path.exists(script_path):
                # Try alternative name
                script_path = os.path.join(script_dir, "lesson_plan_generator.py")
                
            if not os.path.exists(script_path):
                return f'''
                <div class="result error">
                    <h3>❌ Error</h3>
                    <p>Could not find the script file. Make sure the Python script is in the same folder.</p>
                    <p>Tried looking for: <code>Ipneww.py</code> and <code>lesson_plan_generator.py</code></p>
                </div>
                '''
            
            # Build command
            args = [sys.executable, script_path]
            
            # Add inputs
            args.extend(['--teacher', request.form.get('teacher', 'Teacher')])
            args.extend(['--class', request.form.get('class', 'X A')])
            args.extend(['--subject', request.form.get('subject', 'Biology')])
            args.extend(['--chapter', request.form.get('chapter', 'Nutrition')])
            args.extend(['--start_date', request.form.get('start_date', datetime.now().strftime('%d-%m-%Y'))])
            args.extend(['--duration', request.form.get('duration', '40 min')])
            args.extend(['--total_periods', request.form.get('total_periods', '3')])
            
            # Add topics
            total = int(request.form.get('total_periods', 3))
            for i in range(1, total + 1):
                topic = request.form.get(f'topic_p{i}', f'Topic {i}')
                args.extend([f'--topic_p{i}', topic])
            
            # Run the generator
            result = subprocess.run(args, capture_output=True, text=True, cwd=script_dir)
            
            # Find the generated file
            output_path = os.path.join(script_dir, OUTPUT_FOLDER)
            if result.returncode == 0 and os.path.exists(output_path):
                files = os.listdir(output_path)
                if files:
                    latest = max(files, key=lambda f: os.path.getctime(os.path.join(output_path, f)))
                    file_path = os.path.join(output_path, latest)
                    
                    return f'''
                    <div class="result success">
                        <h3>✅ Lesson Plan Generated Successfully!</h3>
                        <p><strong>📄 File:</strong> {latest}</p>
                        <a href="/download/{latest}" class="download-btn">📥 Download Lesson Plan</a>
                        <br><br>
                        <details>
                            <summary>📋 View Output Details</summary>
                            <pre style="background:white;padding:10px;border-radius:4px;font-size:12px;max-height:200px;overflow:auto;">{result.stdout}</pre>
                        </details>
                    </div>
                    '''
            
            return f'''
            <div class="result error">
                <h3>❌ Error Generating Lesson Plan</h3>
                <details>
                    <summary>📋 View Error Details</summary>
                    <pre style="background:white;padding:10px;border-radius:4px;font-size:12px;max-height:300px;overflow:auto;">{result.stderr or result.stdout}</pre>
                </details>
            </div>
            '''
                
        except Exception as e:
            return f'''
            <div class="result error">
                <h3>❌ Error</h3>
                <pre style="background:white;padding:10px;border-radius:4px;font-size:12px;">{str(e)}</pre>
            </div>
            '''
    
    return render_template_string(HTML_TEMPLATE, today=datetime.now().strftime('%d-%m-%Y'))

@app.route('/download/<filename>')
def download(filename):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, OUTPUT_FOLDER, filename)
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    import sys
    # Check if required packages are installed
    try:
        import flask
    except ImportError:
        print("Flask not installed! Run: pip install flask")
        sys.exit(1)
    
    print("=" * 50)
    print("🚀 Lesson Plan Generator Web Interface")
    print("=" * 50)
    print("📁 Script location:", os.path.dirname(os.path.abspath(__file__)))
    print("🌐 Open http://127.0.0.1:5000 in your browser")
    print("Press Ctrl+C to stop the server")
    print("=" * 50)
    
    app.run(debug=False, host='127.0.0.1', port=5000)
