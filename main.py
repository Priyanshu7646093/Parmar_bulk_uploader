"""
Unified Flask Application Launcher
Routes requests to Four_option.py (4 options) or Five_option.py (5 options)
based on user selection from the homepage
"""

from flask import Flask, render_template, request, redirect, url_for
import sys
import os

# Create Flask app
app = Flask(__name__, template_folder='.')

@app.route('/')
def index():
    """Serve the homepage with 4/5 option selector"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_4options():
    """Route 4-option selections to Four_option.py logic"""
    # Import and use Four_option's functions
    from Four_option import upload, uploaded_data as main_data

    # Call Four_option's upload function
    result = upload()
    return result

@app.route('/upload5', methods=['POST'])
def upload_5options():
    """Route 5-option selections to Five_option.py logic"""
    # Import and use Five_option's functions
    from Five_option import upload, uploaded_data as main5_data

    # Call Five_option's upload function
    result = upload()
    return result

@app.route('/generate', methods=['POST'])
def generate():
    """Route to generate function
    Determines which module to use based on form data
    """
    option_count = request.form.get('option_count', '4')
    
    if option_count == '5':
        from Five_option import generate as generate_main5
        return generate_main5()
    else:
        from Four_option import generate as generate_main
        return generate_main()

@app.route('/diagnose', methods=['GET'])
def diagnose():
    """Serve the diagnosis page"""
    return render_template('diagnose.html')

if __name__ == "__main__":
    print("\n" + "="*50)
    print("🚀 Starting Parmar's Bulk Uploader")
    print("="*50)
    print("📝 Access the application at:")
    print("   http://127.0.0.1:5000")
    print("\n💡 How to use:")
    print("   1. Select 4 or 5 options on the homepage")
    print("   2. Upload your PDF file")
    print("   3. Click 'Generate Quiz & Preview'")
    print("\n⏹️  Press CTRL+C to stop the application")
    print("="*50 + "\n")
    
    app.run(host="0.0.0.0", debug=True)

