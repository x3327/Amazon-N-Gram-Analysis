"""
N-gram Automation Web Application
Flask application for automating N-gram analysis of Amazon PPC search term reports.
"""

import os
import uuid
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, url_for
from werkzeug.utils import secure_filename

from utils.csv_parser import parse_csv, filter_asins, group_by_campaign, get_data_summary, validate_csv
from utils.ngram_generator import generate_ngrams, get_ngram_summary
from utils.metrics import aggregate_ngram_metrics, get_campaign_summary
from utils.suggestions import suggest_negatives, get_suggestion_summary
from utils.excel_writer import create_excel_output, generate_output_filename

# Initialize Flask app
app = Flask(__name__)

# Configuration
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'outputs')
app.config['ALLOWED_EXTENSIONS'] = {'csv'}

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)


def allowed_file(filename: str) -> bool:
    """Check if file extension is allowed."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


def process_csv_file(filepath: str, thresholds: dict = None) -> dict:
    """
    Process a CSV file through the entire N-gram analysis pipeline.
    
    Args:
        filepath: Path to the uploaded CSV file
        thresholds: Custom thresholds for NE/NP flagging
        
    Returns:
        Dictionary with processing results and output file path
    """
    # Step 1: Parse CSV
    df = parse_csv(filepath)
    
    # Validate CSV
    is_valid, missing_cols = validate_csv(df)
    if not is_valid:
        return {
            'success': False,
            'error': f"Missing required columns: {', '.join(missing_cols)}",
            'missing_columns': missing_cols
        }
    
    # Get initial summary
    initial_summary = get_data_summary(df)
    
    # Step 2: Filter ASINs
    df_filtered, asin_count = filter_asins(df)
    
    # Step 3: Group by campaign
    campaigns = group_by_campaign(df_filtered)
    
    # Step 4: Process each campaign
    processed_data = {}
    total_flagged = 0
    
    for campaign_name, campaign_df in campaigns.items():
        # Generate N-grams
        ngrams = generate_ngrams(campaign_df)
        
        # Add metrics to each N-gram type
        for ngram_type in ['monograms', 'bigrams', 'trigrams']:
            if ngram_type in ngrams and not ngrams[ngram_type].empty:
                ngrams[ngram_type] = aggregate_ngram_metrics(ngrams[ngram_type])
                ngrams[ngram_type] = suggest_negatives(ngrams[ngram_type], thresholds)
        
        # Also process search terms
        if 'search_terms' in ngrams and not ngrams['search_terms'].empty:
            ngrams['search_terms'] = aggregate_ngram_metrics(ngrams['search_terms'])
            ngrams['search_terms'] = suggest_negatives(ngrams['search_terms'], thresholds)
        
        # Get summaries
        ngram_summary = get_ngram_summary(ngrams)
        campaign_metrics = get_campaign_summary(campaign_df)
        
        # Count flagged items
        flagged_count = 0
        for ngram_type in ['monograms', 'bigrams', 'trigrams', 'search_terms']:
            if ngram_type in ngrams and not ngrams[ngram_type].empty:
                suggestion_summary = get_suggestion_summary(ngrams[ngram_type])
                flagged_count += suggestion_summary['total_flagged']
        
        total_flagged += flagged_count
        
        # Store processed data
        processed_data[campaign_name] = {
            'ngrams': ngrams,
            'summary': {
                **ngram_summary,
                **campaign_metrics,
                'total_flagged': flagged_count
            }
        }
    
    # Step 5: Generate Excel output
    output_filename = generate_output_filename()
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
    create_excel_output(processed_data, output_path)
    
    # Compile final results
    result = {
        'success': True,
        'output_file': output_filename,
        'summary': {
            'original_rows': initial_summary['total_rows'],
            'asins_removed': asin_count,
            'campaigns_processed': len(campaigns),
            'total_search_terms': initial_summary['total_search_terms'],
            'total_spend': initial_summary['total_spend'],
            'total_sales': initial_summary['total_sales'],
            'total_flagged': total_flagged
        },
        'campaigns': list(campaigns.keys())
    }
    
    return result


@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing."""
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': 'Invalid file type. Please upload a CSV file.'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        unique_filename = f"{uuid.uuid4().hex}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(filepath)
        
        # Get custom thresholds from request
        thresholds = None
        if request.form.get('min_clicks'):
            thresholds = {
                'min_clicks_for_ne': int(request.form.get('min_clicks', 3)),
                'min_spend_for_np': float(request.form.get('min_spend', 0.01))
            }
        
        # Process the file
        result = process_csv_file(filepath, thresholds)
        
        # Clean up uploaded file
        try:
            os.remove(filepath)
        except:
            pass
        
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 400
            
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'Processing error: {str(e)}'
        }), 500


@app.route('/download/<filename>')
def download_file(filename):
    """Download the generated Excel file."""
    try:
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], secure_filename(filename))
        
        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': 'File not found'}), 404
        
        return send_file(
            filepath,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/health')
def health_check():
    """Health check endpoint."""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat()
    })


# Cleanup old files periodically (files older than 1 hour)
def cleanup_old_files():
    """Remove old files from upload and output folders."""
    import time
    
    max_age = 3600  # 1 hour in seconds
    current_time = time.time()
    
    for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
        if os.path.exists(folder):
            for filename in os.listdir(folder):
                filepath = os.path.join(folder, filename)
                if os.path.isfile(filepath):
                    file_age = current_time - os.path.getmtime(filepath)
                    if file_age > max_age:
                        try:
                            os.remove(filepath)
                        except:
                            pass


if __name__ == '__main__':
    # Run cleanup on startup
    cleanup_old_files()
    
    # Run the Flask development server
    app.run(debug=True, host='0.0.0.0', port=5000)
