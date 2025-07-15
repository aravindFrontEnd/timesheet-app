from flask import Flask, request, jsonify, render_template_string
from timeverify_processor import TimesheetProcessor  # Import your OCR class
import json
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Initialize your OCR processor
processor = TimesheetProcessor()

@app.route('/')
def home():
    return render_template_string('''
    <!DOCTYPE html>
    <html>
    <head>
        <title>TimeVerify AI - Real OCR Integration</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f5f5; }
            .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; }
            .header { text-align: center; color: #c00; margin-bottom: 30px; }
            .section { margin: 20px 0; padding: 20px; border: 1px solid #ddd; border-radius: 8px; }
            .btn { background: #c00; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; margin: 5px; }
            .btn:hover { background: #a00; }
            .upload-area { border: 2px dashed #ccc; padding: 40px; text-align: center; }
            .results { margin-top: 20px; padding: 20px; background: #f9f9f9; border-radius: 4px; }
            .stats { display: flex; gap: 20px; margin: 20px 0; }
            .stat-card { flex: 1; padding: 20px; background: #e8f4fd; border-radius: 8px; text-align: center; }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>🔴 TimeVerify AI - Real OCR Processing</h1>
                <p>Enterprise Timesheet Processing Platform</p>
                <p><strong>Status:</strong> Using YOUR OCR algorithms + Ready for OpenShift</p>
            </div>
            
            <div class="stats">
                <div class="stat-card">
                    <h3>OCR Engine</h3>
                    <p>✅ Real Tesseract Processing</p>
                </div>
                <div class="stat-card">
                    <h3>Business Logic</h3>
                    <p>✅ Your Pattern Recognition</p>
                </div>
                <div class="stat-card">
                    <h3>Deployment</h3>
                    <p>✅ OpenShift Ready</p>
                </div>
            </div>
            
            <div class="section">
                <h3>📸 Real Screenshot Processing</h3>
                <p><strong>Upload a timesheet screenshot (PNG/JPG)</strong> - Uses your actual OCR algorithms!</p>
                <form id="uploadForm" enctype="multipart/form-data">
                    <div class="upload-area">
                        <input type="file" name="screenshot" accept="image/*" required>
                        <br><br>
                        <input type="text" name="consultant_name" placeholder="Consultant Name (e.g., John Smith)" required>
                        <br><br>
                        <button type="submit" class="btn">🔍 Process with Real OCR</button>
                    </div>
                </form>
            </div>
            
            <div class="section">
                <h3>🚀 Demo Mode (Simulated)</h3>
                <button onclick="runBulkDemo()" class="btn">Run Bulk Demo</button>
                <p>Simulate processing 100+ consultants for presentation</p>
            </div>
            
            <div id="results" class="results" style="display:none;">
                <h3>Processing Results:</h3>
                <pre id="resultData"></pre>
            </div>
        </div>
        
        <script>
            // Real file upload processing
            document.getElementById('uploadForm').onsubmit = async function(e) {
                e.preventDefault();
                
                const formData = new FormData(this);
                const resultsDiv = document.getElementById('results');
                const resultData = document.getElementById('resultData');
                
                try {
                    resultData.textContent = 'Processing with real OCR... This may take 2-3 seconds...';
                    resultsDiv.style.display = 'block';
                    
                    const response = await fetch('/api/process-screenshot', {
                        method: 'POST',
                        body: formData
                    });
                    
                    const result = await response.json();
                    resultData.textContent = JSON.stringify(result, null, 2);
                } catch (error) {
                    resultData.textContent = 'Error: ' + error.message;
                }
            };
            
            // Bulk demo simulation
            async function runBulkDemo() {
                const resultsDiv = document.getElementById('results');
                const resultData = document.getElementById('resultData');
                
                resultData.textContent = 'Running bulk processing simulation...';
                resultsDiv.style.display = 'block';
                
                try {
                    const response = await fetch('/api/bulk-demo');
                    const result = await response.json();
                    resultData.textContent = JSON.stringify(result, null, 2);
                } catch (error) {
                    resultData.textContent = 'Error: ' + error.message;
                }
            }
        </script>
    </body>
    </html>
    ''')

@app.route('/api/process-screenshot', methods=['POST'])
def process_screenshot():
    try:
        if 'screenshot' not in request.files:
            return jsonify({'error': 'No screenshot file provided'}), 400
        
        file = request.files['screenshot']
        consultant_name = request.form.get('consultant_name', 'Unknown')
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        print(f"Processing screenshot for: {consultant_name}")
        
        # Use YOUR REAL OCR processing
        result = processor.process_screenshot_from_bytes(file.read(), consultant_name)
        
        print(f"OCR Results: {result}")
        
        return jsonify(result)
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/bulk-demo')
def bulk_demo():
    """Bulk simulation for demo purposes"""
    demo_consultants = [
        {'name': 'John Smith', 'screenshot_hours': 40, 'ibm_hours': 38, 'discrepancy': True},
        {'name': 'Jane Doe', 'screenshot_hours': 40, 'ibm_hours': 40, 'discrepancy': False},
        {'name': 'Mike Johnson', 'screenshot_hours': 35, 'ibm_hours': 35, 'discrepancy': False},
        {'name': 'Sarah Wilson', 'screenshot_hours': 42, 'ibm_hours': 40, 'discrepancy': True},
        {'name': 'Alex Chen', 'screenshot_hours': 37, 'ibm_hours': 38, 'discrepancy': True}
    ]
    
    results = []
    for consultant in demo_consultants:
        results.append({
            'consultant_name': consultant['name'],
            'screenshot_hours': consultant['screenshot_hours'],
            'ibm_system_hours': consultant['ibm_hours'],
            'discrepancy_detected': consultant['discrepancy'],
            'status': 'discrepancy' if consultant['discrepancy'] else 'verified'
        })
    
    return jsonify({
        'summary': {
            'total_processed': len(demo_consultants),
            'discrepancies_found': sum(1 for c in demo_consultants if c['discrepancy']),
            'processing_time': '4.0 seconds',
            'manual_equivalent': '50+ minutes',
            'annual_savings': '$60,000+'
        },
        'results': results,
        'technology': 'Real OCR + OpenShift Ready'
    })

@app.route('/health')
def health():
    return jsonify({
        'status': 'healthy',
        'service': 'TimeVerify AI - Real OCR Integration',
        'ocr_engine': 'Tesseract with your algorithms',
        'timestamp': datetime.now().isoformat()
    })

if __name__ == '__main__':
    print("🚀 TimeVerify AI with REAL OCR - Starting...")
    print("🌐 Open: http://localhost:5000")
    print("🔍 Features: Real OCR processing + Demo simulation")
    print("✅ Ready for OpenShift deployment!")
    app.run(host='0.0.0.0', port=5000, debug=True)