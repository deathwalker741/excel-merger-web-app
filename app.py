from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
import os
import traceback
from datetime import datetime
from merger import process_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'excel-merger-secret-key-2024-v2'

# Configuration
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
ALLOWED_EXTENSIONS = {'xlsx'}

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Check if file has allowed extension"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            # Check if file was uploaded
            if 'excel_file' not in request.files:
                flash('❌ No file selected!')
                return redirect(request.url)
            
            file = request.files['excel_file']
            
            # Check if file was actually selected
            if file.filename == '':
                flash('❌ No file selected!')
                return redirect(request.url)
            
            # Check if file is allowed
            if not (file and allowed_file(file.filename)):
                flash('❌ Invalid file format. Please upload a .xlsx file.')
                return redirect(request.url)

            # Secure the filename
            filename = secure_filename(file.filename)
            
            # Add timestamp to avoid conflicts
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            name, ext = os.path.splitext(filename)
            input_filename = f"{name}_{timestamp}{ext}"
            
            # Save uploaded file
            input_path = os.path.join(UPLOAD_FOLDER, input_filename)
            file.save(input_path)
            
            # Generate output filename
            output_filename = f"merged_{input_filename}"
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            
            # Process the file
            original_rows, merged_rows = process_file(input_path, output_path)
            
            # Clean up uploaded file
            try:
                os.remove(input_path)
            except:
                pass  # Don't fail if cleanup fails
            
            flash(f'✅ Success! Processed {original_rows} rows into {merged_rows} merged rows.')
            
            # Send file for download
            return send_file(
                output_path, 
                as_attachment=True,
                download_name=f"merged_{name}_{timestamp}.xlsx"
            )
            
        except Exception as e:
            # Log the full error for debugging
            print(f"Error processing file: {str(e)}")
            print(f"Traceback: {traceback.format_exc()}")
            
            # Clean up files on error
            try:
                if 'input_path' in locals() and os.path.exists(input_path):
                    os.remove(input_path)
                if 'output_path' in locals() and os.path.exists(output_path):
                    os.remove(output_path)
            except:
                pass
            
            flash(f'❌ Error processing file: {str(e)}')
            return redirect(request.url)
    
    return render_template("index.html")

@app.route("/about")
def about():
    """About page explaining what the app does"""
    return render_template("about.html")

@app.route("/health")
def health():
    """Health check endpoint"""
    return jsonify({"status": "healthy", "timestamp": datetime.now().isoformat()})

@app.errorhandler(404)
def not_found_error(error):
    return render_template("index.html"), 404

@app.errorhandler(500)
def internal_error(error):
    print(f"Internal server error: {str(error)}")
    flash('❌ Internal server error. Please try again.')
    return render_template("index.html"), 500

if __name__ == "__main__":
    import os
    port = int(os.environ.get('PORT', 5000))
    debug_mode = os.environ.get('FLASK_ENV') == 'development'
    app.run(debug=debug_mode, host='0.0.0.0', port=port)