"""
N-gram Automation Web Application
Flask application for automating N-gram analysis of Amazon PPC search term reports.
"""

import os
import uuid
import json
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, url_for
from werkzeug.utils import secure_filename

from utils.csv_parser import parse_csv, filter_asins, group_by_campaign, get_data_summary, validate_csv, filter_active_campaigns
from utils.ngram_generator import generate_ngrams, get_ngram_summary
from utils.metrics import aggregate_ngram_metrics, get_campaign_summary
from utils.excel_writer import create_excel_output, generate_output_filename, create_asin_report

# Initialize Flask app
app = Flask(__name__)

# Configuration
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'outputs')
app.config['ARCHIVE_FOLDER'] = os.path.join(os.path.dirname(__file__), 'archive')
app.config['ARCHIVE_FILE'] = os.path.join(os.path.dirname(__file__), 'archive', 'archive.json')
app.config['ALLOWED_EXTENSIONS'] = {'csv', 'xlsx', 'xls'}

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs(app.config['ARCHIVE_FOLDER'], exist_ok=True)


def allowed_file(filename: str) -> bool:
    """Check if file extension is allowed."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


def load_archive():
    """Load archive data from JSON file."""
    if os.path.exists(app.config['ARCHIVE_FILE']):
        try:
            with open(app.config['ARCHIVE_FILE'], 'r') as f:
                return json.load(f)
        except:
            return []
    return []


def save_archive(archive_data):
    """Save archive data to JSON file."""
    with open(app.config['ARCHIVE_FILE'], 'w') as f:
        json.dump(archive_data, f, indent=2)


def process_csv_file(filepath: str) -> dict:
    """
    Process a CSV file through the entire N-gram analysis pipeline.
    
    Args:
        filepath: Path to the uploaded CSV file
        
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
    
    # Step 2: Filter to only active campaigns
    df = filter_active_campaigns(df)
    
    # Step 3: Filter ASINs and get ASIN data separately
    df_filtered, df_asins, asin_count = filter_asins(df)
    
    # Step 3: Group by campaign
    campaigns = group_by_campaign(df_filtered)
    
    # Step 4: Process each campaign
    processed_data = {}
    campaign_details = {}
    total_orders = 0
    
    for campaign_name, campaign_df in campaigns.items():
        # Generate N-grams
        ngrams = generate_ngrams(campaign_df)
        
        # Add metrics to each N-gram type
        for ngram_type in ['monograms', 'bigrams', 'trigrams']:
            if ngram_type in ngrams and not ngrams[ngram_type].empty:
                ngrams[ngram_type] = aggregate_ngram_metrics(ngrams[ngram_type])
        
        # Also process search terms
        if 'search_terms' in ngrams and not ngrams['search_terms'].empty:
            ngrams['search_terms'] = aggregate_ngram_metrics(ngrams['search_terms'])
        
        # Get summaries
        ngram_summary = get_ngram_summary(ngrams)
        campaign_metrics = get_campaign_summary(campaign_df)
        
        # Store processed data
        processed_data[campaign_name] = {
            'ngrams': ngrams,
            'summary': {
                **ngram_summary,
                **campaign_metrics
            }
        }
        
        # Store campaign details for analytics
        campaign_details[campaign_name] = {
            'monograms': ngram_summary.get('monogram_count', 0),
            'bigrams': ngram_summary.get('bigram_count', 0),
            'trigrams': ngram_summary.get('trigram_count', 0),
            'search_terms': ngram_summary.get('search_term_count', 0),
            'spend': campaign_metrics.get('total_spend', 0),
            'sales': campaign_metrics.get('total_sales', 0),
            'orders': campaign_metrics.get('total_orders', 0),
            'impressions': campaign_metrics.get('total_impressions', 0),
            'clicks': campaign_metrics.get('total_clicks', 0)
        }
        
        total_orders += campaign_metrics.get('total_orders', 0)
    
    # Step 5: Generate Excel output for N-gram analysis
    output_filename = generate_output_filename()
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
    create_excel_output(processed_data, output_path)
    
    # Step 6: Generate ASIN report if there are ASINs
    asin_filename = None
    if asin_count > 0 and not df_asins.empty:
        asin_filename = generate_output_filename(prefix="ASIN_Report")
        asin_output_path = os.path.join(app.config['OUTPUT_FOLDER'], asin_filename)
        create_asin_report(df_asins, asin_output_path)
    
    # Compile final results
    result = {
        'success': True,
        'output_file': output_filename,
        'asin_file': asin_filename,
        'summary': {
            'original_rows': initial_summary['total_rows'],
            'asins_removed': asin_count,
            'campaigns_processed': len(campaigns),
            'total_search_terms': initial_summary['total_search_terms'],
            'total_spend': initial_summary['total_spend'],
            'total_sales': initial_summary['total_sales'],
            'total_orders': total_orders
        },
        'campaigns': list(campaigns.keys()),
        'campaign_details': campaign_details
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
        
        # Process the file
        result = process_csv_file(filepath)
        
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


@app.route('/archive', methods=['GET', 'POST'])
def archive():
    """Handle archive operations - list all or create new."""
    if request.method == 'GET':
        # Return list of archived reports
        archives = load_archive()
        return jsonify({
            'success': True,
            'archives': archives
        })
    
    elif request.method == 'POST':
        # Archive a new report
        try:
            data = request.get_json()
            
            if not data or 'filename' not in data:
                return jsonify({'success': False, 'error': 'No filename provided'}), 400
            
            # Create archive entry
            archive_entry = {
                'id': str(uuid.uuid4()),
                'filename': data.get('filename'),
                'originalFilename': data.get('originalFilename', 'Unknown'),
                'summary': data.get('summary', {}),
                'processedAt': data.get('processedAt', datetime.now().isoformat())
            }
            
            # Load existing archives
            archives = load_archive()
            
            # Add new entry at the beginning
            archives.insert(0, archive_entry)
            
            # Save archives
            save_archive(archives)
            
            return jsonify({
                'success': True,
                'archive': archive_entry
            })
            
        except Exception as e:
            return jsonify({
                'success': False,
                'error': f'Failed to archive: {str(e)}'
            }), 500


@app.route('/archive/<archive_id>', methods=['GET', 'DELETE'])
def archive_item(archive_id):
    """Handle operations on a specific archive item."""
    archives = load_archive()
    
    # Find the archive
    archive_entry = None
    archive_index = None
    for i, item in enumerate(archives):
        if item.get('id') == archive_id:
            archive_entry = item
            archive_index = i
            break
    
    if archive_entry is None:
        return jsonify({'success': False, 'error': 'Archive not found'}), 404
    
    if request.method == 'GET':
        return jsonify({
            'success': True,
            'archive': archive_entry
        })
    
    elif request.method == 'DELETE':
        try:
            # Remove from archives
            archives.pop(archive_index)
            save_archive(archives)
            
            return jsonify({
                'success': True,
                'message': 'Archive deleted successfully'
            })
        except Exception as e:
            return jsonify({
                'success': False,
                'error': f'Failed to delete: {str(e)}'
            }), 500


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
