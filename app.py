from flask import Flask, request, render_template, jsonify, session, Response, stream_with_context
from email_sender import EmailSender
from functools import wraps
from datetime import datetime, timedelta
import os
from werkzeug.utils import secure_filename
import json
from pathlib import Path
import shutil
import pandas as pd
from apscheduler.schedulers.background import BackgroundScheduler
import pytz

app = Flask(__name__)
email_sender = EmailSender()
app.config['UPLOAD_FOLDER'] = 'temp_uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create upload folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Add rate limiting to prevent spam
RATE_LIMIT = timedelta(minutes=1)  # 1 email per minute
last_email_time = {}

# Add a rate limit decorator
def rate_limit(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        ip = request.remote_addr
        if ip in last_email_time:
            time_since_last = datetime.now() - last_email_time[ip]
            if time_since_last < RATE_LIMIT:
                return jsonify({
                    'success': False,
                    'message': 'Please wait before sending another email'
                }), 429
        last_email_time[ip] = datetime.now()
        return f(*args, **kwargs)
    return decorated_function

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'pdf'

# Get the directory where app.py is located
BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_FILE = BASE_DIR / 'email_templates.json'
TEMPLATE_ATTACHMENTS_DIR = BASE_DIR / 'template_attachments'

# Create template attachments directory if it doesn't exist
os.makedirs(TEMPLATE_ATTACHMENTS_DIR, exist_ok=True)

# Load templates from file
def load_templates():
    """Load templates from JSON file"""
    try:
        if Path(TEMPLATES_FILE).exists():
            with open(TEMPLATES_FILE, 'r') as f:
                templates = json.load(f)
            print(f"Loaded {len(templates)} templates from {TEMPLATES_FILE}")
            return templates
        print(f"No existing templates file found at {TEMPLATES_FILE}")
        return {}
    except Exception as e:
        print(f"Error loading templates: {e}")
        return {}

# Save templates to file
def save_templates(templates):
    """Save templates to JSON file"""
    try:
        with open(TEMPLATES_FILE, 'w') as f:
            json.dump(templates, f, indent=4)
        print(f"Saved {len(templates)} templates to {TEMPLATES_FILE}")
    except Exception as e:
        print(f"Error saving templates: {e}")

# Initialize templates
EMAIL_TEMPLATES = load_templates()

# Initialize the scheduler
scheduler = BackgroundScheduler()
scheduler.start()
SCHEDULED_JOBS = {}  # To store scheduled email jobs

@app.route('/')
def home():
    """Render the landing page"""
    return render_template('index.html')

@app.route('/email-form')
def email_form():
    """Render the email form page"""
    return render_template('email_form.html')

@app.route('/send-email', methods=['POST'])
@rate_limit
def send_email_endpoint():
    """Handle email sending requests from both API and form submissions"""
    try:
        print("Starting email send process...")
        data = request.form
        
        recipient = data.get('recipient')
        subject = data.get('subject')
        body = data.get('body')
        template_name = data.get('template')
        schedule_time = data.get('schedule_time')
        print(f"Using template: {template_name}")
        
        attachment = request.files.get('attachment')
        attachment_path = None
        template_attachment_path = None

        if not all([recipient, subject, body]):
            return jsonify({
                'success': False,
                'message': 'Missing required fields'
            }), 400

        # Handle template attachment if a template is selected
        if template_name and template_name in EMAIL_TEMPLATES:
            template = EMAIL_TEMPLATES[template_name]
            print(f"Template found: {template}")
            if template.get('attachment_name'):
                template_attachment_path = os.path.join(
                    TEMPLATE_ATTACHMENTS_DIR, 
                    template['attachment_name']
                )
                print(f"Template attachment path: {template_attachment_path}")
                if not os.path.exists(template_attachment_path):
                    print(f"Warning: Template attachment not found at {template_attachment_path}")
                else:
                    print(f"Template attachment found at {template_attachment_path}")
        
        # Handle PDF attachment if provided
        if attachment and attachment.filename:
            if not allowed_file(attachment.filename):
                return jsonify({
                    'success': False,
                    'message': 'Only PDF files are allowed'
                }), 400
            
            filename = secure_filename(attachment.filename)
            attachment_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            attachment.save(attachment_path)

        # Use template attachment if no custom attachment is provided
        if not attachment_path and template_attachment_path and os.path.exists(template_attachment_path):
            print(f"Using template attachment: {template_attachment_path}")
            attachment_path = template_attachment_path

        if schedule_time:
            # Convert schedule_time to datetime and validate
            try:
                schedule_datetime = datetime.fromisoformat(schedule_time)
                
                # Check if the scheduled time is in the future
                if schedule_datetime <= datetime.now():
                    return jsonify({
                        'success': False,
                        'message': 'Schedule time must be in the future'
                    }), 400
                
            except ValueError as e:
                return jsonify({
                    'success': False,
                    'message': 'Invalid schedule time format'
                }), 400
            
            job_id = f"email_{datetime.now().timestamp()}"
            
            # Store job info
            SCHEDULED_JOBS[job_id] = {
                'recipient': recipient,
                'subject': subject,
                'scheduled_time': schedule_datetime,
                'attachment_path': attachment_path
            }
            
            # Schedule the email
            scheduler.add_job(
                email_sender.send_email,
                'date',
                run_date=schedule_datetime,
                args=[recipient, subject, body, attachment_path],
                id=job_id,
                misfire_grace_time=300  # 5 minutes grace time
            )
            
            return jsonify({
                'success': True,
                'message': f'Email scheduled for {schedule_datetime.strftime("%Y-%m-%d %H:%M")}'
            })
        else:
            # Send email immediately
            success, error = email_sender.send_email(recipient, subject, body, attachment_path)

        # Clean up the temporary file (but not template attachments)
        if attachment_path and attachment_path != template_attachment_path and os.path.exists(attachment_path):
            os.remove(attachment_path)

        if success:
            return jsonify({
                'success': True,
                'message': 'Email sent successfully!'
            })
        else:
            return jsonify({
                'success': False,
                'message': error
            }), 400

    except Exception as e:
        # Clean up the temporary file in case of error
        if attachment_path and os.path.exists(attachment_path):
            os.remove(attachment_path)
        return jsonify({
            'success': False,
            'message': str(e)
        }), 500

@app.route('/get-templates')
def get_templates():
    """Return available email templates"""
    return jsonify({
        'success': True,
        'templates': list(EMAIL_TEMPLATES.keys())
    })

@app.route('/get-template/<template_name>')
def get_template(template_name):
    """Return specific template content"""
    template = EMAIL_TEMPLATES.get(template_name)
    if template:
        return jsonify({
            'success': True,
            'template': template
        })
    return jsonify({
        'success': False,
        'message': 'Template not found'
    }), 404

@app.route('/manage-templates')
def manage_templates():
    """Render the template management page"""
    return render_template('manage_templates.html')

@app.route('/save-template', methods=['POST'])
def save_template():
    """Save a new template"""
    try:
        name = request.form['name']
        template = {
            'subject': request.form['subject'],
            'body': request.form['body'],
            'attachment_name': None
        }
        
        # Handle PDF attachment if provided
        if 'attachment' in request.files:
            attachment = request.files['attachment']
            if attachment and allowed_file(attachment.filename):
                # Save the attachment with a unique name
                filename = secure_filename(f"{name}_{attachment.filename}")
                attachment_path = os.path.join(TEMPLATE_ATTACHMENTS_DIR, filename)
                print(f"Saving template attachment to: {attachment_path}")
                attachment.save(attachment_path)
                template['attachment_name'] = filename
                print(f"Template attachment saved as: {filename}")
        
        EMAIL_TEMPLATES[name] = template
        save_templates(EMAIL_TEMPLATES)
        print(f"Template saved: {template}")
        
        return jsonify({
            'success': True,
            'message': 'Template saved successfully'
        })
    except Exception as e:
        print(f"Error saving template: {str(e)}")
        return jsonify({
            'success': False,
            'message': str(e)
        }), 400

@app.route('/update-template', methods=['POST'])
def update_template():
    """Update an existing template"""
    try:
        data = request.get_json()
        name = data['name']
        template = {
            'subject': data['subject'],
            'body': data['body']
        }
        
        EMAIL_TEMPLATES[name] = template
        save_templates(EMAIL_TEMPLATES)
        
        return jsonify({
            'success': True,
            'message': 'Template updated successfully'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': str(e)
        }), 400

@app.route('/delete-template', methods=['POST'])
def delete_template():
    """Delete a template"""
    try:
        data = request.get_json()
        name = data['name']
        
        if name in EMAIL_TEMPLATES:
            # Delete associated attachment if it exists
            template = EMAIL_TEMPLATES[name]
            if template.get('attachment_name'):
                attachment_path = os.path.join(TEMPLATE_ATTACHMENTS_DIR, template['attachment_name'])
                if os.path.exists(attachment_path):
                    os.remove(attachment_path)
            
            del EMAIL_TEMPLATES[name]
            save_templates(EMAIL_TEMPLATES)
        
        return jsonify({
            'success': True,
            'message': 'Template deleted successfully'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': str(e)
        }), 400

@app.route('/debug/templates')
def debug_templates():
    """Debug endpoint to view current templates"""
    return jsonify({
        'templates_file': str(TEMPLATES_FILE),
        'templates_file_exists': Path(TEMPLATES_FILE).exists(),
        'templates': EMAIL_TEMPLATES,
        'attachments_dir': str(TEMPLATE_ATTACHMENTS_DIR),
        'attachments': os.listdir(TEMPLATE_ATTACHMENTS_DIR) if os.path.exists(TEMPLATE_ATTACHMENTS_DIR) else []
    })

@app.route('/bulk-email')
def bulk_email():
    """Render the bulk email page"""
    return render_template('bulk_email.html')

@app.route('/send-bulk-email', methods=['POST'])
def send_bulk_email():
    def generate():
        try:
            template_name = request.form.get('template')
            excel_file = request.files.get('excelFile')
            attachment = request.files.get('attachment')
            
            # Read Excel file
            df = pd.read_excel(excel_file)
            total_emails = len(df)
            
            # Save attachment if provided
            attachment_path = None
            if attachment and attachment.filename:
                if allowed_file(attachment.filename):
                    filename = secure_filename(attachment.filename)
                    attachment_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    attachment.save(attachment_path)

            # Get template
            template = EMAIL_TEMPLATES.get(template_name)
            if not template:
                yield json.dumps({
                    'status': {'success': False, 'message': 'Template not found'}
                }) + '\n'
                return

            # Process each email
            for index, row in df.iterrows():
                try:
                    email = row['email']
                    name = row.get('name', '')
                    
                    # Personalize subject and body
                    subject = template['subject']
                    body = template['body']
                    if name:
                        body = body.replace('[name]', name)
                    
                    # Get template attachment
                    template_attachment = None
                    if template.get('attachment_name'):
                        template_attachment = os.path.join(
                            TEMPLATE_ATTACHMENTS_DIR, 
                            template['attachment_name']
                        )
                    
                    # Use custom attachment or template attachment
                    final_attachment = attachment_path or template_attachment
                    
                    # Send email
                    success, error = email_sender.send_email(
                        recipient=email,
                        subject=subject,
                        body=body,
                        attachment_path=final_attachment
                    )
                    
                    # Report status
                    status_message = f"Email to {email}: {'Success' if success else error}"
                    yield json.dumps({
                        'progress': ((index + 1) / total_emails) * 100,
                        'current': index + 1,
                        'total': total_emails,
                        'status': {
                            'success': success,
                            'message': status_message
                        }
                    }) + '\n'
                    
                except Exception as e:
                    yield json.dumps({
                        'status': {
                            'success': False,
                            'message': f"Error processing {email}: {str(e)}"
                        }
                    }) + '\n'

            # Clean up temporary attachment
            if attachment_path and os.path.exists(attachment_path):
                os.remove(attachment_path)
                
        except Exception as e:
            yield json.dumps({
                'status': {
                    'success': False,
                    'message': f"Error processing bulk emails: {str(e)}"
                }
            }) + '\n'

    return Response(stream_with_context(generate()), mimetype='application/json')

@app.route('/get-scheduled-emails', methods=['GET'])
def get_scheduled_emails():
    scheduled_emails = []
    for job_id, job in SCHEDULED_JOBS.items():
        scheduled_emails.append({
            'id': job_id,
            'recipient': job['recipient'],
            'subject': job['subject'],
            'scheduled_time': job['scheduled_time'].strftime("%Y-%m-%d %H:%M"),
            'status': 'Pending'
        })
    return jsonify({'success': True, 'scheduled_emails': scheduled_emails})

@app.route('/cancel-scheduled-email/<job_id>', methods=['POST'])
def cancel_scheduled_email(job_id):
    if job_id in SCHEDULED_JOBS:
        scheduler.remove_job(job_id)
        del SCHEDULED_JOBS[job_id]
        return jsonify({'success': True, 'message': 'Email cancelled successfully'})
    return jsonify({'success': False, 'message': 'Scheduled email not found'})

@app.route('/scheduled-emails')
def scheduled_emails():
    return render_template('scheduled_emails.html')

@app.route('/quick-add')
def quick_add():
    """Render the quick add recipients page"""
    return render_template('quick_add.html')

@app.route('/send-quick-add-emails', methods=['POST'])
def send_quick_add_emails():
    """Handle quick add recipients email sending"""
    def generate():
        try:
            contentType = request.form.get('contentType', 'custom')
            attachment = request.files.get('attachment')
            
            # Parse recipients from form data
            recipients = []
            for key, value in request.form.items():
                if key.startswith('recipients[') and key.endswith('][name]'):
                    # Extract index from key like "recipients[1][name]"
                    index = key.split('[')[1].split(']')[0]
                    name = value
                    email_key = f'recipients[{index}][email]'
                    email = request.form.get(email_key)
                    
                    if name and email:
                        recipients.append({'name': name, 'email': email})
            
            if not recipients:
                yield json.dumps({
                    'status': {'success': False, 'message': 'No recipients found'}
                }) + '\n'
                return
            
            total_emails = len(recipients)
            
            # Save attachment if provided
            attachment_path = None
            if attachment and attachment.filename:
                if allowed_file(attachment.filename):
                    filename = secure_filename(attachment.filename)
                    attachment_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    attachment.save(attachment_path)

            # Get email content based on type
            if contentType == 'custom':
                # Use custom subject and body
                subject = request.form.get('subject', '')
                body = request.form.get('body', '')
                template_attachment = None
            else:
                # Use template
                template_name = request.form.get('template')
                template = EMAIL_TEMPLATES.get(template_name)
                if not template:
                    yield json.dumps({
                        'status': {'success': False, 'message': 'Template not found'}
                    }) + '\n'
                    return
                subject = template['subject']
                body = template['body']
                # Get template attachment
                if template.get('attachment_name'):
                    template_attachment = os.path.join(
                        TEMPLATE_ATTACHMENTS_DIR, 
                        template['attachment_name']
                    )
                else:
                    template_attachment = None

            # Process each email
            for index, recipient in enumerate(recipients):
                try:
                    email = recipient['email']
                    name = recipient['name']
                    
                    # Personalize subject and body
                    personalized_subject = subject
                    personalized_body = body
                    
                    if name:
                        # Replace [name] with actual name in both subject and body
                        personalized_subject = subject.replace('[name]', name)
                        personalized_body = body.replace('[name]', name)
                        # Convert line breaks to HTML format
                        personalized_body = personalized_body.replace('\n', '<br>')
                        # Always add "Hello [name]" at the beginning
                        personalized_body = f"Hello {name},<br><br>{personalized_body}"
                    
                    # Use custom attachment or template attachment
                    final_attachment = attachment_path or template_attachment
                    
                    # Send email
                    success, error = email_sender.send_email(
                        recipient=email,
                        subject=personalized_subject,
                        body=personalized_body,
                        attachment_path=final_attachment
                    )
                    
                    # Report status
                    status_message = f"Email to {name} ({email}): {'Success' if success else error}"
                    yield json.dumps({
                        'progress': ((index + 1) / total_emails) * 100,
                        'current': index + 1,
                        'total': total_emails,
                        'status': {
                            'success': success,
                            'message': status_message
                        }
                    }) + '\n'
                    
                except Exception as e:
                    yield json.dumps({
                        'status': {
                            'success': False,
                            'message': f"Error processing {recipient.get('email', 'unknown')}: {str(e)}"
                        }
                    }) + '\n'

            # Clean up temporary attachment
            if attachment_path and os.path.exists(attachment_path):
                os.remove(attachment_path)
                
        except Exception as e:
            yield json.dumps({
                'status': {
                    'success': False,
                    'message': f"Error processing quick add emails: {str(e)}"
                }
            }) + '\n'

    return Response(stream_with_context(generate()), mimetype='application/json')

if __name__ == '__main__':
    app.run(debug=True) 