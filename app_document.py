from flask import Flask, request, jsonify, render_template_string
from timeverify_processor import TimesheetProcessor  # Your OCR class
import json
from datetime import datetime
import os
import tempfile

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size for documents

# Initialize your OCR processor
processor = TimesheetProcessor()

@app.route('/')
def home():
    return render_template_string('''
    <!DOCTYPE html>
    <html>
    <head>
        <title>TimeVerify AI - Document Processing</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f5f5; }
            .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; }
            .header { text-align: center; color: #c00; margin-bottom: 30px; }
            .section { margin: 20px 0; padding: 20px; border: 1px solid #ddd; border-radius: 8px; }
            .btn { background: #c00; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; margin: 5px; }
            .btn:hover { background: #a00; }
            .upload-area { border: 2px dashed #ccc; padding: 40px; text-align: center; }
            .results { margin-top: 20px; padding: 20px; background: #f9f9f9; border-radius: 4px; max-height: 400px; overflow-y: auto; }
            .stats { display: flex; gap: 20px; margin: 20px 0; }
            .stat-card { flex: 1; padding: 20px; background: #e8f4fd; border-radius: 8px; text-align: center; }
            .file-info { background: #f0f8ff; padding: 10px; margin: 10px 0; border-radius: 4px; }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>🔴 TimeVerify AI - Document Processing</h1>
                <p>Enterprise Timesheet Processing Platform</p>
                <p><strong>Status:</strong> Process Word documents with embedded timesheet images</p>
            </div>
            
            <div class="stats">
                <div class="stat-card">
                    <h3>Document Processing</h3>
                    <p>✅ .docx/.doc Support</p>
                </div>
                <div class="stat-card">
                    <h3>Image Extraction</h3>
                    <p>✅ Auto Extract from Docs</p>
                </div>
                <div class="stat-card">
                    <h3>OCR Engine</h3>
                    <p>✅ Your Tesseract Logic</p>
                </div>
            </div>
            
            <div class="section">
                <h3>📄 Upload Word Document with Timesheet Screenshots</h3>
                <div class="file-info">
                    <strong>Supported formats:</strong> .docx, .doc files<br>
                    <strong>Process:</strong> Upload → Extract embedded images → OCR processing → Results
                </div>
                <form id="uploadForm" enctype="multipart/form-data">
                    <div class="upload-area">
                        <input type="file" name="document" accept=".docx,.doc" required>
                        <br><br>
                        <input type="text" name="consultant_name" placeholder="Consultant Name (or leave empty to extract from filename)" style="width: 300px;">
                        <br><br>
                        <button type="submit" class="btn">📄 Process Document</button>
                    </div>
                </form>
            </div>
            
            <div class="section">
                <h3>🗂️ Bulk Document Processing</h3>
                <form id="bulkUploadForm" enctype="multipart/form-data">
                    <div class="upload-area">
                        <input type="file" name="documents" accept=".docx,.doc" multiple required>
                        <br><br>
                        <button type="submit" class="btn">📁 Process Multiple Documents</button>
                        <p><small>Select multiple .docx files to process in bulk</small></p>
                    </div>
                </form>
            </div>
            
            <div class="section">
                <h3>🎯 Demo Simulation</h3>
                <button onclick="runDemo()" class="btn">🚀 Run Demo (No Upload Required)</button>
                <p>Simulate processing for presentation purposes</p>
            </div>
            
            <div id="results" class="results" style="display:none;">
                <h3>Processing Results:</h3>
                <div id="processingStatus"></div>
                <pre id="resultData"></pre>
            </div>
        </div>
        
        <script>
            // Single document processing
            document.getElementById('uploadForm').onsubmit = async function(e) {
                e.preventDefault();
                await processDocuments(this, '/api/process-document', false);
            };
            
            // Bulk document processing
            document.getElementById('bulkUploadForm').onsubmit = async function(e) {
                e.preventDefault();
                await processDocuments(this, '/api/process-documents', true);
            };
            
            async function processDocuments(form, endpoint, isBulk) {
                const formData = new FormData(form);
                const resultsDiv = document.getElementById('results');
                const resultData = document.getElementById('resultData');
                const statusDiv = document.getElementById('processingStatus');
                
                try {
                    const fileCount = isBulk ? form.documents.files.length : 1;
                    statusDiv.innerHTML = `<strong>Processing ${fileCount} document(s)...</strong><br>Extracting images and running OCR...`;
                    resultsDiv.style.display = 'block';
                    resultData.textContent = 'Processing...';
                    
                    const response = await fetch(endpoint, {
                        method: 'POST',
                        body: formData
                    });
                    
                    const result = await response.json();
                    
                    statusDiv.innerHTML = `<strong>✅ Processing Complete!</strong><br>Found ${result.total_images || 0} images, ${result.total_entries || 0} timesheet entries`;
                    resultData.textContent = JSON.stringify(result, null, 2);
                } catch (error) {
                    statusDiv.innerHTML = '<strong>❌ Error occurred</strong>';
                    resultData.textContent = 'Error: ' + error.message;
                }
            }
            
            // Demo simulation
            async function runDemo() {
                const resultsDiv = document.getElementById('results');
                const resultData = document.getElementById('resultData');
                const statusDiv = document.getElementById('processingStatus');
                
                statusDiv.innerHTML = '<strong>🎯 Running Demo Simulation...</strong>';
                resultsDiv.style.display = 'block';
                resultData.textContent = 'Generating demo data...';
                
                try {
                    const response = await fetch('/api/demo');
                    const result = await response.json();
                    
                    statusDiv.innerHTML = '<strong>✅ Demo Complete!</strong><br>Simulated processing of 5 consultants';
                    resultData.textContent = JSON.stringify(result, null, 2);
                } catch (error) {
                    statusDiv.innerHTML = '<strong>❌ Demo Error</strong>';
                    resultData.textContent = 'Error: ' + error.message;
                }
            }
        </script>
    </body>
    </html>
    ''')

@app.route('/api/process-document', methods=['POST'])
def process_document():
    """Process a single Word document"""
    try:
        if 'document' not in request.files:
            return jsonify({'error': 'No document file provided'}), 400
        
        file = request.files['document']
        consultant_name = request.form.get('consultant_name', '')
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_file:
            file.save(temp_file.name)
            temp_path = temp_file.name
        
        try:
            # Extract consultant name from filename if not provided
            if not consultant_name:
                consultant_name = processor.extract_name_from_filename(file.filename)
            
            print(f"Processing document: {file.filename} for consultant: {consultant_name}")
            
            # Extract images from Word document using your existing method
            images = processor.extract_images_from_word_file(temp_path)
            
            if not images:
                return jsonify({
                    'error': 'No images found in the document',
                    'consultant_name': consultant_name,
                    'filename': file.filename
                })
            
            # Process each image with OCR
            all_entries = []
            image_results = []
            
            for idx, image in enumerate(images, 1):
                print(f"Processing image {idx}/{len(images)}")
                
                # Extract text using your OCR method
                text = processor.extract_text_from_image(image)
                
                # Parse entries using your parsing method
                entries = processor.parse_timesheet_entries(text, consultant_name)
                all_entries.extend(entries)
                
                image_results.append({
                    'image_number': idx,
                    'text_extracted': len(text) > 0,
                    'entries_found': len(entries),
                    'raw_text_sample': text[:200] + '...' if len(text) > 200 else text
                })
            
            # Calculate totals
            total_hours = sum(entry['Hours'] for entry in all_entries 
                            if isinstance(entry['Hours'], (int, float)))
            
            # Simulate IBM system check
            ibm_hours = processor.simulate_ibm_system_check(consultant_name)
            
            result = {
                'consultant_name': consultant_name,
                'filename': file.filename,
                'total_images': len(images),
                'total_entries': len(all_entries),
                'screenshot_hours': total_hours,
                'ibm_system_hours': ibm_hours,
                'discrepancy_detected': total_hours != ibm_hours,
                'image_results': image_results,
                'all_entries': all_entries,
                'processing_timestamp': datetime.now().isoformat(),
                'status': 'Processed successfully'
            }
            
            print(f"Processing complete: {total_hours} hours extracted from {len(images)} images")
            return jsonify(result)
            
        finally:
            # Clean up temporary file
            os.unlink(temp_path)
            
    except Exception as e:
        print(f"Error processing document: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/process-documents', methods=['POST'])
def process_documents():
    """Process multiple Word documents"""
    try:
        if 'documents' not in request.files:
            return jsonify({'error': 'No document files provided'}), 400
        
        files = request.files.getlist('documents')
        
        if not files:
            return jsonify({'error': 'No files selected'}), 400
        
        all_results = []
        total_images = 0
        total_entries = 0
        
        for file in files:
            if file.filename == '':
                continue
                
            try:
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_file:
                    file.save(temp_file.name)
                    temp_path = temp_file.name
                
                try:
                    # Extract consultant name from filename
                    consultant_name = processor.extract_name_from_filename(file.filename)
                    
                    print(f"Processing document: {file.filename}")
                    
                    # Extract images from Word document
                    images = processor.extract_images_from_word_file(temp_path)
                    
                    if not images:
                        all_results.append({
                            'consultant_name': consultant_name,
                            'filename': file.filename,
                            'status': 'No images found',
                            'error': 'Document contains no embedded images'
                        })
                        continue
                    
                    # Process all images in this document
                    file_entries = []
                    for image in images:
                        text = processor.extract_text_from_image(image)
                        entries = processor.parse_timesheet_entries(text, consultant_name)
                        file_entries.extend(entries)
                    
                    # Calculate hours for this consultant
                    consultant_hours = sum(entry['Hours'] for entry in file_entries 
                                         if isinstance(entry['Hours'], (int, float)))
                    
                    ibm_hours = processor.simulate_ibm_system_check(consultant_name)
                    
                    file_result = {
                        'consultant_name': consultant_name,
                        'filename': file.filename,
                        'images_processed': len(images),
                        'entries_found': len(file_entries),
                        'screenshot_hours': consultant_hours,
                        'ibm_system_hours': ibm_hours,
                        'discrepancy_detected': consultant_hours != ibm_hours,
                        'status': 'Processed successfully'
                    }
                    
                    all_results.append(file_result)
                    total_images += len(images)
                    total_entries += len(file_entries)
                    
                finally:
                    # Clean up temporary file
                    os.unlink(temp_path)
                    
            except Exception as e:
                all_results.append({
                    'filename': file.filename,
                    'status': 'Processing failed',
                    'error': str(e)
                })
        
        # Summary statistics
        successful_files = [r for r in all_results if r.get('status') == 'Processed successfully']
        discrepancies = [r for r in successful_files if r.get('discrepancy_detected')]
        
        summary = {
            'total_documents': len(files),
            'successful_documents': len(successful_files),
            'total_images': total_images,
            'total_entries': total_entries,
            'discrepancies_found': len(discrepancies),
            'processing_time': f"{len(files) * 2} seconds (estimated)",
            'manual_equivalent': f"{len(files) * 10} minutes"
        }
        
        return jsonify({
            'summary': summary,
            'results': all_results,
            'total_images': total_images,
            'total_entries': total_entries,
            'processing_timestamp': datetime.now().isoformat()
        })
        
    except Exception as e:
        print(f"Error processing documents: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/demo')
def demo():
    """Demo simulation without file upload"""
    demo_data = {
        'summary': {
            'scenario': 'Simulated bulk document processing',
            'total_documents': 5,
            'total_images': 12,
            'total_entries': 89,
            'discrepancies_found': 2,
            'processing_time': '10 seconds',
            'manual_equivalent': '3+ hours'
        },
        'results': [
            {
                'consultant_name': 'John Smith',
                'filename': 'john_smith_week49.docx',
                'images_processed': 3,
                'screenshot_hours': 40,
                'ibm_system_hours': 38,
                'discrepancy_detected': True,
                'status': 'Processed successfully'
            },
            {
                'consultant_name': 'Jane Doe',
                'filename': 'jane_doe_timesheet.docx',
                'images_processed': 2,
                'screenshot_hours': 40,
                'ibm_system_hours': 40,
                'discrepancy_detected': False,
                'status': 'Processed successfully'
            },
            {
                'consultant_name': 'Mike Johnson',
                'filename': 'mike_johnson_hours.docx',
                'images_processed': 2,
                'screenshot_hours': 35,
                'ibm_system_hours': 35,
                'discrepancy_detected': False,
                'status': 'Processed successfully'
            },
            {
                'consultant_name': 'Sarah Wilson',
                'filename': 'sarah_wilson_week49.docx',
                'images_processed': 3,
                'screenshot_hours': 42,
                'ibm_system_hours': 40,
                'discrepancy_detected': True,
                'status': 'Processed successfully'
            },
            {
                'consultant_name': 'Alex Chen',
                'filename': 'alex_chen_timesheet.docx',
                'images_processed': 2,
                'screenshot_hours': 38,
                'ibm_system_hours': 38,
                'discrepancy_detected': False,
                'status': 'Processed successfully'
            }
        ],
        'roi_analysis': {
            'annual_savings': '$60,000+',
            'time_saved_weekly': '10+ hours',
            'error_reduction': '95%',
            'scalability': 'Unlimited consultants'
        },
        'technology_stack': 'Your OCR algorithms + Flask + OpenShift ready'
    }
    
    return jsonify(demo_data)

@app.route('/health')
def health():
    return jsonify({
        'status': 'healthy',
        'service': 'TimeVerify AI - Document Processing',
        'features': ['Word document processing', 'Image extraction', 'OCR processing'],
        'timestamp': datetime.now().isoformat()
    })

if __name__ == '__main__':
    print("🚀 TimeVerify AI Document Processor - Starting...")
    print("🌐 Open: http://localhost:5000")
    print("📄 Features: Word document upload → Image extraction → OCR processing")
    print("✅ Ready for OpenShift deployment!")
    app.run(host='0.0.0.0', port=5000, debug=True)