from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_jwt_extended import JWTManager, create_access_token, create_refresh_token, jwt_required, get_jwt_identity
from werkzeug.utils import secure_filename
import os
import shutil
from datetime import datetime
import logging
from config import config
from models.user import db, User, Subscription, Favorite, ProcessingLog
from auth.auth_handler import AuthHandler
from auth.email_sender import mail, EmailSender

# ============ UPDATED IMPORTS ============ #
from utils.excel_processor import convert_xls_to_xlsx, validate_excel_file, batch_convert_xls_files
from utils.data_extractor import (
    extract_data_from_excel, 
    create_summary_report,
    get_extraction_statistics
)
from utils.visits_distributor import (
    distribute_visits,
    validate_visits_file,
    get_distribution_summary
)
import zipfile

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('xtractor.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__, static_folder='../frontend', static_url_path='')

# ========== Production Logging Setup ========== #
env = os.environ.get('FLASK_ENV', 'development')
if env == 'production':
    import logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    logger.info("="*50)
    logger.info("XTRACTOR PRO - STARTING IN PRODUCTION MODE")
    logger.info(f"Port: {os.environ.get('PORT', 8080)}")
    logger.info(f"Database: {app.config['SQLALCHEMY_DATABASE_URI']}")
    logger.info(f"Upload folder: {app.config['UPLOAD_FOLDER']}")
    logger.info("="*50)

# ========== إصلاح فوري لقاعدة البيانات ========== #
# إنشاء مجلد data في نفس المجلد
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
os.makedirs(DATA_DIR, exist_ok=True)
print(f"Data directory: {DATA_DIR}")
# ================================================ #

# Load configuration
env = os.environ.get('FLASK_ENV', 'development')
app.config.from_object(config[env])

# ========== فرض مسار واضح ومباشر ========== #
DB_FILE = os.path.join(DATA_DIR, 'xtractor.db')
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_FILE}'
print(f"Database file: {DB_FILE}")
# =========================================== #

config[env].init_app(app)

# Initialize extensions
db.init_app(app)
mail.init_app(app)
CORS(app, resources={
    r"/api/*": {
        "origins": "*",
        "methods": ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
        "allow_headers": ["Content-Type", "Authorization"]
    }
})

# JWT Manager
jwt = JWTManager(app)

# Rate Limiter
limiter = Limiter(
    app=app,
    key_func=get_remote_address,
    default_limits=["200 per day", "50 per hour"]
)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


# ======================== STATIC FILES ======================== #
@app.route('/')
def index():
    """Serve the main page"""
    return send_from_directory(app.static_folder, 'index.html')


@app.route('/login')
def login_page():
    """Serve login page"""
    return send_from_directory(app.static_folder, 'login.html')


@app.route('/register')
def register_page():
    """Serve registration page"""
    return send_from_directory(app.static_folder, 'register.html')

# ✅ أضف هذا الكود الجديد هنا:
@app.route('/activate')
def activate_page():
    """Serve activation page"""
    return send_from_directory(app.static_folder, 'activate.html')

# ======================== HEALTH CHECK ======================== #
@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'version': '3.0',
        'timestamp': datetime.now().isoformat(),
        'environment': env
    })


# ======================== AUTHENTICATION ======================== #

@app.route('/api/auth/register', methods=['POST'])
@limiter.limit("5 per hour")
def register():
    """Register new user with subscription period selection"""
    try:
        data = request.get_json()
        
        email = data.get('email')
        username = data.get('username')
        password = data.get('password')
        full_name = data.get('full_name')
        subscription_period = data.get('subscription_period', 'monthly')  # ✅ جديد
        
        if not email or not username or not password:
            return jsonify({'error': 'جميع الحقول مطلوبة'}), 400
        
        if subscription_period not in ['monthly', 'yearly']:  # ✅ جديد
            return jsonify({'error': 'مدة الاشتراك غير صالحة'}), 400
        
        result = AuthHandler.register_user(email, username, password, full_name, subscription_period)  # ✅ محدث
        
        # 4 مسافات بادئة (بافتراض أن السطر الذي يسبقها هو 'if result['success']:')
        if result['success']:
            # 8 مسافات بادئة
            # 🚀 التصحيح الحاسم: استخراج user_id من القاموس
            user_id = result['user']['id']  
            
            # Create tokens (مع تطبيق الإصلاح السابق بـ str() لمنع 422)
            access_token = create_access_token(identity=str(user_id))
            refresh_token = create_refresh_token(identity=str(user_id))
            
            return jsonify({
                'success': True,
                'user': result['user'],
                'access_token': access_token, 
                'refresh_token': refresh_token
            }), 200 # رمز الحالة 200 (OK) هو الأنسب لتسجيل الدخول
        
        # 4 مسافات بادئة
        else:
            # ... (باقي منطق رسائل الخطأ 401/403) ...
            return jsonify(result), 400
            
    except Exception as e:
        logger.error(f"Registration error: {e}")
        return jsonify({'error': 'حدث خطأ أثناء التسجيل'}), 500

@app.route('/api/auth/login', methods=['POST'])
@limiter.limit("10 per hour")
def login():
    """Login user - updated to handle activation requirement"""
    try:
        data = request.get_json()
        
        email = data.get('email')
        password = data.get('password')
        
        if not email or not password:
            return jsonify({'error': 'البريد الإلكتروني وكلمة المرور مطلوبان'}), 400
        
        result = AuthHandler.login_user(email, password)
        
        if result['success']:
            # ✅ استخراج معرف المستخدم من نتيجة تسجيل الدخول
            user_id = result['user']['id']
            
            # ✅ إنشاء الرموز (Tokens) باستخدام الـ user_id
            access_token = create_access_token(identity=str(user_id))
            refresh_token = create_refresh_token(identity=str(user_id))
            
            return jsonify({
                'success': True,
                'user': result['user'],
                'subscription': result['subscription'],
                'access_token': access_token,
                'refresh_token': refresh_token
            }), 200

        else:
            # ✅ تحقق من الحاجة إلى التفعيل
            if result.get('requires_activation'):
                return jsonify({
                    'success': False,
                    'error': result['error'],
                    'requires_activation': True,
                    'user_id': result.get('user_id')
                }), 403
            
            # ✅ تحقق من الحاجة إلى التجديد
            if result.get('requires_renewal'):
                return jsonify({
                    'success': False,
                    'error': result['error'],
                    'requires_renewal': True,
                    'user_id': result.get('user_id')
                }), 403
            
            return jsonify(result), 401
            
    except Exception as e:
        logger.error(f"Login error: {e}", exc_info=True)
        return jsonify({'error': 'حدث خطأ أثناء تسجيل الدخول'}), 500
    


@app.route('/api/auth/refresh', methods=['POST'])
@jwt_required(refresh=True)
def refresh():
    """Refresh access token"""
    try:
        user_id = get_jwt_identity()
        access_token = create_access_token(identity=str(user_id))
        
        return jsonify({
            'success': True,
            'access_token': access_token
        })
        
    except Exception as e:
        logger.error(f"Token refresh error: {e}")
        return jsonify({'error': 'فشل تجديد الرمز'}), 500


@app.route('/api/auth/me', methods=['GET'])
@jwt_required()
def get_current_user():
    """Get current user info"""
    try:
        user_id = get_jwt_identity()
        user = User.query.get(user_id)
        
        if not user:
            return jsonify({'error': 'المستخدم غير موجود'}), 404
        
        return jsonify({
            'success': True,
            'user': user.to_dict()
        })
        
    except Exception as e:
        logger.error(f"Get user error: {e}")
        return jsonify({'error': 'فشل جلب بيانات المستخدم'}), 500

    # ✅✅✅ أضف جميع هذه الـ Routes الجديدة هنا ✅✅✅

@app.route('/api/auth/activate', methods=['POST'])
@limiter.limit("10 per hour")
def activate():
    """Activate user account with activation code"""
    try:
        data = request.get_json()
        
        user_id = data.get('user_id')
        activation_code = data.get('activation_code')
        
        if not user_id or not activation_code:
            return jsonify({'error': 'معرف المستخدم وكود التفعيل مطلوبان'}), 400
        
        result = AuthHandler.activate_subscription(user_id, activation_code)
        
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 400
            
    except Exception as e:
        logger.error(f"Activation error: {e}")
        return jsonify({'error': 'فشل تفعيل الحساب'}), 500


@app.route('/api/auth/activate-by-email', methods=['POST'])
@limiter.limit("10 per hour")
def activate_by_email():
    """Activate user account by email and activation code"""
    try:
        data = request.get_json()
        
        email = data.get('email')
        activation_code = data.get('activation_code')
        
        if not email or not activation_code:
            return jsonify({'error': 'البريد الإلكتروني وكود التفعيل مطلوبان'}), 400
        
        # Find user by email
        user = User.query.filter_by(email=email).first()
        if not user:
            return jsonify({'error': 'المستخدم غير موجود'}), 404
        
        result = AuthHandler.activate_subscription(user.id, activation_code)
        
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 400
            
    except Exception as e:
        logger.error(f"Activation by email error: {e}")
        return jsonify({'error': 'فشل تفعيل الحساب'}), 500


@app.route('/api/auth/resend-activation', methods=['POST'])
@limiter.limit("3 per hour")
def resend_activation():
    """Resend activation code to user email"""
    try:
        data = request.get_json()
        
        email = data.get('email')
        
        if not email:
            return jsonify({'error': 'البريد الإلكتروني مطلوب'}), 400
        
        # Find user by email
        user = User.query.filter_by(email=email).first()
        if not user:
            return jsonify({'error': 'المستخدم غير موجود'}), 404
        
        if user.is_active and user.subscription and user.subscription.is_active():
            return jsonify({'error': 'الحساب مفعل بالفعل'}), 400
        
        # Get subscription period from existing subscription
        subscription_period = 'monthly'
        if user.subscription and user.subscription.subscription_period:
            subscription_period = user.subscription.subscription_period
        
        result = AuthHandler.request_activation_code(user.id, subscription_period)
        
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 400
            
    except Exception as e:
        logger.error(f"Resend activation error: {e}")
        return jsonify({'error': 'فشل إرسال كود التفعيل'}), 500


# ======================== SUBSCRIPTION ======================== #
@app.route('/api/subscription/status', methods=['GET'])
@jwt_required()
def subscription_status():
    """Get subscription status"""
    try:
        user_id = get_jwt_identity()
        result = AuthHandler.check_subscription(user_id)
        
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"Subscription status error: {e}")
        return jsonify({'error': 'فشل جلب حالة الاشتراك'}), 500


@app.route('/api/subscription/request-code', methods=['POST'])
@jwt_required()
@limiter.limit("3 per day")
def request_activation_code():
    """Request new activation code"""
    try:
        user_id = get_jwt_identity()
        result = AuthHandler.request_activation_code(user_id)
        
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 400
            
    except Exception as e:
        logger.error(f"Request code error: {e}")
        return jsonify({'error': 'فشل طلب كود التفعيل'}), 500


@app.route('/api/subscription/activate', methods=['POST'])
@jwt_required()
@limiter.limit("10 per day")
def activate_subscription():
    """Activate subscription with code"""
    try:
        user_id = get_jwt_identity()
        data = request.get_json()
        
        activation_code = data.get('activation_code')
        if not activation_code:
            return jsonify({'error': 'كود التفعيل مطلوب'}), 400
        
        result = AuthHandler.activate_subscription(user_id, activation_code)
        
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 400
            
    except Exception as e:
        logger.error(f"Activation error: {e}")
        return jsonify({'error': 'فشل تفعيل الاشتراك'}), 500

    # ✅ أضف هذا الـ Route الجديد هنا:

@app.route('/api/subscription/renew', methods=['POST'])
@jwt_required()
@limiter.limit("5 per day")
def renew_subscription():
    """Renew expired subscription with new code"""
    try:
        user_id = get_jwt_identity()
        data = request.get_json()
        
        activation_code = data.get('activation_code')
        subscription_period = data.get('subscription_period', 'monthly')
        
        if not activation_code:
            return jsonify({'error': 'كود التفعيل مطلوب'}), 400
        
        if subscription_period not in ['monthly', 'yearly']:
            return jsonify({'error': 'مدة الاشتراك غير صالحة'}), 400
        
        result = AuthHandler.renew_subscription(user_id, activation_code, subscription_period)
        
        if result['success']:
            return jsonify(result)
        else:
            return jsonify(result), 400
            
    except Exception as e:
        logger.error(f"Renewal error: {e}")
        return jsonify({'error': 'فشل تجديد الاشتراك'}), 500
    
#======================== FILE UPLOAD ======================== #

@app.route('/api/process', methods=['POST'])
@jwt_required()
def process_files_batch():
    """
    معالجة دفعة من الملفات مع تحسينات
    """
    try:
        user_id = get_jwt_identity()
        user = User.query.get(user_id)
        if not user:
            return jsonify({'error': 'المستخدم غير موجود'}), 404
        
        # ✅ التحقق من الملفات
        files = request.files.getlist('files')
        job_type = request.form.get('job_type')
        mode = request.form.get('mode', 'engine_idle')
        sheet_name = request.form.get('sheet_name', 'إجمالي')
        
        if not files or not any(f.filename for f in files):
            return jsonify({'error': 'الرجاء اختيار ملف واحد على الأقل'}), 400

        if job_type not in ['cars', 'visits']:
            return jsonify({'error': 'نوع العملية غير صحيح'}), 400
        
        # ✅ إنشاء مجلد مؤقت فريد
        temp_dir_name = f"{user_id}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        user_upload_folder = os.path.join(app.config['UPLOAD_FOLDER'], temp_dir_name)
        os.makedirs(user_upload_folder, exist_ok=True)
        
        uploaded_files_paths = []
        
        # ✅ حفظ الملفات المرفوعة
        for file in files:
            if file and file.filename:
                ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
                if ext in app.config['ALLOWED_EXTENSIONS']:
                    filename = secure_filename(file.filename)
                    file_path = os.path.join(user_upload_folder, filename)
                    file.save(file_path)
                    uploaded_files_paths.append(file_path)
                    logger.info(f"File uploaded: {filename}")

        if not uploaded_files_paths:
            shutil.rmtree(user_upload_folder, ignore_errors=True)
            return jsonify({'error': 'لم يتم قبول أي ملف صالح'}), 400

        # ✅ تحويل XLS إلى XLSX
        try:
            conversion_results = batch_convert_xls_files(user_upload_folder, user_upload_folder)
            logger.info(f"Conversion results: {conversion_results}")
        except Exception as conv_error:
            logger.error(f"Conversion error: {conv_error}")
            shutil.rmtree(user_upload_folder, ignore_errors=True)
            return jsonify({'error': f'فشل في تحويل الملفات: {str(conv_error)}'}), 500
        
        # ✅ جمع ملفات XLSX النهائية
        final_xlsx_files = []
        for item in conversion_results.get('converted', []):
            converted_path = os.path.join(user_upload_folder, item['converted'])
            if os.path.exists(converted_path):
                final_xlsx_files.append(converted_path)
        
        # إضافة XLSX التي لم تحتاج تحويل
        for filename in os.listdir(user_upload_folder):
            if filename.lower().endswith('.xlsx'):
                file_path = os.path.join(user_upload_folder, filename)
                if file_path not in final_xlsx_files:
                    final_xlsx_files.append(file_path)

        if not final_xlsx_files:
            shutil.rmtree(user_upload_folder, ignore_errors=True)
            return jsonify({'error': 'فشل في معالجة الملفات إلى تنسيق XLSX'}), 500
        
        logger.info(f"Final XLSX files count: {len(final_xlsx_files)}")
        
        # ✅ معالجة حسب نوع العملية
        if job_type == 'cars':
            # استخلاص البيانات
            all_extracted_data = []
            
            for file_path in final_xlsx_files:
                try:
                    if os.path.exists(file_path):
                        data = extract_data_from_excel(file_path, mode, app.config.get('ZONE_POINTS'))
                        if data:
                            all_extracted_data.extend(data)
                        logger.info(f"Extracted {len(data)} records from {os.path.basename(file_path)}")
                except Exception as extract_error:
                    logger.error(f"Extraction error for {file_path}: {extract_error}")
                    continue
            
            if not all_extracted_data:
                shutil.rmtree(user_upload_folder, ignore_errors=True)
                return jsonify({'error': 'لم يتم استخلاص أي بيانات من الملفات'}), 400
            
            # إنشاء تقرير موحد
            stats = get_extraction_statistics(all_extracted_data)
            
            report_name = f"Summary_Report_{mode}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            report_path = os.path.join(user_upload_folder, report_name)
            
            try:
                create_summary_report(all_extracted_data, report_path, mode)
                logger.info(f"Report created: {report_name}")
            except Exception as report_error:
                logger.error(f"Report creation error: {report_error}")
                shutil.rmtree(user_upload_folder, ignore_errors=True)
                return jsonify({'error': f'فشل في إنشاء التقرير: {str(report_error)}'}), 500
            
            # تسجيل العملية
            log = ProcessingLog(
                user_id=user_id,
                job_type=job_type,
                mode=mode,
                filename=f"Batch ({len(final_xlsx_files)} files)",
                status='completed',
                records_processed=stats['total_records'],
                inside_zone=stats['inside_zone'],
                outside_zone=stats['outside_zone'],
                undefined_zone=stats['undefined_zone']
            )
            db.session.add(log)
            db.session.commit()
            
            # تنظيف وإرسال
            try:
                return send_file(
                    report_path,
                    as_attachment=True,
                    download_name=report_name,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            finally:
                # تنظيف بعد الإرسال
                def cleanup():
                    import time
                    time.sleep(2)  # انتظار إتمام الإرسال
                    try:
                        shutil.rmtree(user_upload_folder, ignore_errors=True)
                    except:
                        pass
                
                import threading
                threading.Thread(target=cleanup, daemon=True).start()
        
        elif job_type == 'visits':
            # التوزيع للزيارات
            if len(final_xlsx_files) != 1:
                shutil.rmtree(user_upload_folder, ignore_errors=True)
                return jsonify({'error': 'يجب اختيار ملف واحد فقط للتوزيع'}), 400
            
            file_to_distribute = final_xlsx_files[0]
            
            output_dir_name = f"Visits_Dist_{temp_dir_name}"
            output_dir_path = os.path.join(app.config['OUTPUT_FOLDER'], output_dir_name)
            os.makedirs(output_dir_path, exist_ok=True)
            
            try:
                distribution_results = distribute_visits(file_to_distribute, sheet_name, output_dir_path)
                logger.info(f"Distribution completed: {distribution_results['total_files']} files")
            except Exception as dist_error:
                logger.error(f"Distribution error: {dist_error}")
                shutil.rmtree(user_upload_folder, ignore_errors=True)
                shutil.rmtree(output_dir_path, ignore_errors=True)
                return jsonify({'error': f'فشل في التوزيع: {str(dist_error)}'}), 500
            
            # ضغط الملفات
            zip_filename = f"Xtractor_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
            zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
            
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(output_dir_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, os.path.relpath(file_path, output_dir_path))
            
            # تسجيل
            log = ProcessingLog(
                user_id=user_id,
                job_type=job_type,
                filename=os.path.basename(file_to_distribute),
                status='completed',
                records_processed=distribution_results['total_files']
            )
            db.session.add(log)
            db.session.commit()
            
            # تنظيف وإرسال
            try:
                return send_file(
                    zip_path,
                    as_attachment=True,
                    download_name=zip_filename,
                    mimetype='application/zip'
                )
            finally:
                def cleanup():
                    import time
                    time.sleep(2)
                    try:
                        shutil.rmtree(user_upload_folder, ignore_errors=True)
                        shutil.rmtree(output_dir_path, ignore_errors=True)
                        os.remove(zip_path) if os.path.exists(zip_path) else None
                    except:
                        pass
                
                import threading
                threading.Thread(target=cleanup, daemon=True).start()
        
    except Exception as e:
        logger.error(f"Critical error in process_files_batch: {e}", exc_info=True)
        
        # تنظيف في حالة الفشل
        if 'user_upload_folder' in locals() and os.path.exists(user_upload_folder):
            shutil.rmtree(user_upload_folder, ignore_errors=True)
        if 'output_dir_path' in locals() and os.path.exists(output_dir_path):
            shutil.rmtree(output_dir_path, ignore_errors=True)
        
        return jsonify({'error': f'فشل في عملية المعالجة: {str(e)}'}), 500
    
# ======================== CAR PROCESSING ======================== #
@app.route('/api/process-cars', methods=['POST'])
@jwt_required()
def process_cars():
    """Process car tracking data with enhanced features"""
    try:
        user_id = get_jwt_identity()
        data = request.get_json()
        
        # Check subscription
        user = User.query.get(user_id)
        if not user.has_active_subscription():
            return jsonify({
                'error': 'انتهت صلاحية الاشتراك'
            }), 403
        
        filename = data.get('filename')
        mode = data.get('mode', 'engine_idle')
        
        if not filename:
            return jsonify({'error': 'اسم الملف مطلوب'}), 400
        
        if mode not in ['engine_idle', 'parking_details']:
            return jsonify({'error': 'وضع غير صالح'}), 400
        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        if not os.path.exists(filepath):
            return jsonify({'error': 'الملف غير موجود'}), 404
        
        # Create processing log
        log = ProcessingLog(
            user_id=user_id,
            job_type='cars',
            mode=mode,
            filename=filename,
            status='processing'
        )
        db.session.add(log)
        db.session.commit()
        
        try:
            # Create output folder
            job_id = datetime.now().strftime('%Y%m%d_%H%M%S')
            job_folder = os.path.join(app.config['OUTPUT_FOLDER'], f"cars_{user_id}_{job_id}")
            os.makedirs(job_folder, exist_ok=True)
            
            # Convert XLS to XLSX if needed
            if filename.endswith('.xls'):
                logger.info(f"Converting XLS to XLSX: {filename}")
                try:
                    filepath = convert_xls_to_xlsx(filepath, job_folder)
                    logger.info(f"Conversion successful: {filepath}")
                except Exception as conv_error:
                    raise Exception(f"فشل تحويل ملف XLS: {str(conv_error)}")
            
            # Extract data
            logger.info(f"Extracting data from {filename} (mode: {mode})...")
            try:
                extracted_data = extract_data_from_excel(
                    filepath,
                    mode=mode,
                    zone_points=app.config['ZONE_POINTS']
                )
            except Exception as extract_error:
                raise Exception(f"فشل استخراج البيانات: {str(extract_error)}")
            
            if not extracted_data:
                log.status = 'failed'
                log.error_message = 'No data extracted'
                log.completed_at = datetime.utcnow()
                db.session.commit()
                return jsonify({'error': 'لم يتم العثور على بيانات في الملف'}), 400
            
            # Create summary report
            report_name = f"Summary_Report_{mode}_{job_id}.xlsx"
            report_path = os.path.join(job_folder, report_name)
            
            logger.info(f"Creating summary report: {report_name}")
            try:
                create_summary_report(extracted_data, report_path, mode=mode)
            except Exception as report_error:
                raise Exception(f"فشل إنشاء التقرير: {str(report_error)}")
            
            # Calculate statistics using the new function
            stats = get_extraction_statistics(extracted_data)
            
            # Update log
            log.status = 'completed'
            log.records_processed = stats['total_records']
            log.inside_zone = stats['inside_zone']
            log.outside_zone = stats['outside_zone']
            log.undefined_zone = stats['undefined_zone']
            log.completed_at = datetime.utcnow()
            db.session.commit()
            
            result = {
                'success': True,
                'job_id': job_id,
                'report_filename': report_name,
                'total_records': stats['total_records'],
                'statistics': {
                    'inside_zone': stats['inside_zone'],
                    'outside_zone': stats['outside_zone'],
                    'undefined': stats['undefined_zone'],
                    'unique_cars': stats['unique_cars'],
                    'sheets_processed': stats['sheets_processed']
                }
            }
            
            logger.info(f"Processing complete for user {user_id}: {stats['total_records']} records")
            return jsonify(result)
            
        except Exception as proc_error:
            log.status = 'failed'
            log.error_message = str(proc_error)
            log.completed_at = datetime.utcnow()
            db.session.commit()
            raise proc_error
    
    except Exception as e:
        logger.error(f"Processing error: {e}", exc_info=True)
        error_msg = str(e)
        
        # User-friendly error messages
        if "Expected BOF record" in error_msg or "Unsupported format" in error_msg:
            error_msg = "الملف المرفوع ليس ملف Excel حقيقي. يرجى التأكد من صيغة الملف الصحيحة."
        elif "Cannot read XLS file" in error_msg:
            error_msg = "فشل قراءة ملف XLS. يرجى التأكد من أن الملف غير تالف."
        elif "No data extracted" in error_msg:
            error_msg = "لم يتم العثور على بيانات قابلة للاستخراج في الملف."
        
        return jsonify({'error': error_msg}), 500


# ======================== VISITS DISTRIBUTION ======================== #
@app.route('/api/process-visits', methods=['POST'])
@jwt_required()
def process_visits():
    """Process and distribute visits report with validation"""
    try:
        user_id = get_jwt_identity()
        data = request.get_json()
        
        # Check subscription
        user = User.query.get(user_id)
        if not user.has_active_subscription():
            return jsonify({'error': 'انتهت صلاحية الاشتراك'}), 403
        
        filename = data.get('filename')
        sheet_name = data.get('sheet_name', 'إجمالي')
        
        if not filename:
            return jsonify({'error': 'اسم الملف مطلوب'}), 400
        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        if not os.path.exists(filepath):
            return jsonify({'error': 'الملف غير موجود'}), 404
        
        # Validate visits file structure
        logger.info(f"Validating visits file: {filename}")
        is_valid, validation_message, validation_info = validate_visits_file(filepath, sheet_name)
        
        if not is_valid:
            logger.error(f"Visits file validation failed: {validation_message}")
            return jsonify({
                'error': validation_message,
                'info': validation_info
            }), 400
        
        # Create processing log
        log = ProcessingLog(
            user_id=user_id,
            job_type='visits',
            filename=filename,
            status='processing'
        )
        db.session.add(log)
        db.session.commit()
        
        try:
            # Create output folder
            job_id = datetime.now().strftime('%Y%m%d_%H%M%S')
            job_folder = os.path.join(app.config['OUTPUT_FOLDER'], f"visits_{user_id}_{job_id}")
            os.makedirs(job_folder, exist_ok=True)
            
            # Distribute visits
            logger.info(f"Distributing visits from {filename} (sheet: {sheet_name})...")
            try:
                results = distribute_visits(filepath, sheet_name, job_folder)
            except Exception as dist_error:
                raise Exception(f"فشل توزيع الزيارات: {str(dist_error)}")
            
            # Create ZIP file
            zip_filename = f"visits_distribution_{user_id}_{job_id}.zip"
            zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
            
            logger.info(f"Creating ZIP archive: {zip_filename}")
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(job_folder):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.basename(file_path)
                        zipf.write(file_path, arcname)
            
            # Generate summary
            summary = get_distribution_summary(results)
            logger.info(f"Distribution summary:\n{summary}")
            
            # Update log
            log.status = 'completed'
            log.records_processed = results['total_files']
            log.completed_at = datetime.utcnow()
            db.session.commit()
            
            result = {
                'success': True,
                'job_id': job_id,
                'zip_filename': zip_filename,
                'supervisors_count': len(results['supervisors']),
                'representatives_count': len(results['representatives']),
                'total_files': results['total_files'],
                'supervisors': results['supervisors'],
                'representatives': results['representatives'],
                'validation_info': validation_info
            }
            
            logger.info(f"Distribution complete for user {user_id}: {results['total_files']} files")
            return jsonify(result)
            
        except Exception as proc_error:
            log.status = 'failed'
            log.error_message = str(proc_error)
            log.completed_at = datetime.utcnow()
            db.session.commit()
            raise proc_error
    
    except Exception as e:
        logger.error(f"Distribution error: {e}", exc_info=True)
        error_msg = str(e)
        
        # User-friendly error messages
        if "Expected BOF record" in error_msg or "Unsupported format" in error_msg:
            error_msg = "الملف المرفوع ليس ملف Excel حقيقي."
        elif "File must have at least 3 columns" in error_msg:
            error_msg = "الملف يجب أن يحتوي على 3 أعمدة على الأقل: المشرف، المندوب، الخط."
        elif "Cannot read Excel file" in error_msg:
            error_msg = "فشل قراءة ملف Excel. يرجى التأكد من صحة الملف."
        
        return jsonify({'error': error_msg}), 500


# ======================== NEW: VALIDATE FILE ======================== #
@app.route('/api/validate-file', methods=['POST'])
@jwt_required()
def validate_file():
    """Validate uploaded Excel file before processing"""
    try:
        user_id = get_jwt_identity()
        data = request.get_json()
        
        filename = data.get('filename')
        file_type = data.get('type', 'cars')  # 'cars' or 'visits'
        
        if not filename:
            return jsonify({'error': 'اسم الملف مطلوب'}), 400
        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        if not os.path.exists(filepath):
            return jsonify({'error': 'الملف غير موجود'}), 404
        
        if file_type == 'visits':
            sheet_name = data.get('sheet_name', 'إجمالي')
            is_valid, message, info = validate_visits_file(filepath, sheet_name)
        else:
            is_valid, file_format, message = validate_excel_file(filepath)
            info = {'file_type': file_format} if is_valid else {}
        
        return jsonify({
            'success': True,
            'is_valid': is_valid,
            'message': message,
            'info': info
        })
        
    except Exception as e:
        logger.error(f"Validation error: {e}")
        return jsonify({'error': 'فشل التحقق من الملف'}), 500


# ======================== NEW: FILE INFO ======================== #
@app.route('/api/file-info/<filename>', methods=['GET'])
@jwt_required()
def get_file_info(filename):
    """Get detailed information about uploaded file"""
    try:
        user_id = get_jwt_identity()
        filename = secure_filename(filename)
        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        if not os.path.exists(filepath):
            return jsonify({'error': 'الملف غير موجود'}), 404
        
        # Get file info
        file_stats = os.stat(filepath)
        file_ext = os.path.splitext(filename)[1].lower()
        
        info = {
            'filename': filename,
            'size': file_stats.st_size,
            'size_mb': round(file_stats.st_size / (1024 * 1024), 2),
            'extension': file_ext,
            'created_at': datetime.fromtimestamp(file_stats.st_ctime).isoformat(),
            'modified_at': datetime.fromtimestamp(file_stats.st_mtime).isoformat()
        }
        
        # Try to get Excel-specific info
        if file_ext in ['.xlsx', '.xls']:
            try:
                if file_ext == '.xlsx':
                    import openpyxl
                    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
                    info['sheets'] = [ws.title for ws in wb.worksheets]
                    info['sheets_count'] = len(wb.worksheets)
                    wb.close()
                else:
                    import pandas as pd
                    xls = pd.ExcelFile(filepath, engine='xlrd')
                    info['sheets'] = xls.sheet_names
                    info['sheets_count'] = len(xls.sheet_names)
                    xls.close()
            except Exception as excel_error:
                logger.warning(f"Could not read Excel file info: {excel_error}")
                info['sheets'] = []
                info['sheets_count'] = 0
        
        return jsonify({
            'success': True,
            'info': info
        })
        
    except Exception as e:
        logger.error(f"File info error: {e}")
        return jsonify({'error': 'فشل جلب معلومات الملف'}), 500


# ======================== DOWNLOADS ======================== #
@app.route('/api/download/<job_type>/<filename>', methods=['GET'])
@jwt_required()
def download_file(job_type, filename):
    """Download processed file"""
    try:
        user_id = get_jwt_identity()
        filename = secure_filename(filename)
        
        if job_type == 'cars':
            # Find the job folder
            for folder in os.listdir(app.config['OUTPUT_FOLDER']):
                if folder.startswith(f'cars_{user_id}_'):
                    folder_path = os.path.join(app.config['OUTPUT_FOLDER'], folder)
                    file_path = os.path.join(folder_path, filename)
                    if os.path.exists(file_path):
                        return send_file(file_path, as_attachment=True, download_name=filename)
        
        elif job_type == 'visits':
            file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
            if os.path.exists(file_path):
                return send_file(file_path, as_attachment=True, download_name=filename)
        
        return jsonify({'error': 'الملف غير موجود'}), 404
    
    except Exception as e:
        logger.error(f"Download error: {e}")
        return jsonify({'error': 'فشل تحميل الملف'}), 500


# ======================== FAVORITES ======================== #
@app.route('/api/favorites', methods=['GET'])
@jwt_required()
def get_favorites():
    """Get user's favorite folders"""
    try:
        user_id = get_jwt_identity()
        favorites = Favorite.query.filter_by(user_id=user_id).order_by(Favorite.created_at.desc()).all()
        
        return jsonify({
            'success': True,
            'favorites': [fav.to_dict() for fav in favorites]
        })
        
    except Exception as e:
        logger.error(f"Get favorites error: {e}")
        return jsonify({'error': 'فشل جلب المفضلة'}), 500


@app.route('/api/favorites', methods=['POST'])
@jwt_required()
def add_favorite():
    """Add folder to favorites"""
    try:
        user_id = get_jwt_identity()
        data = request.get_json()
        
        folder_path = data.get('folder_path')
        name = data.get('name')
        
        if not folder_path:
            return jsonify({'error': 'مسار المجلد مطلوب'}), 400
        
        # Check if already exists
        existing = Favorite.query.filter_by(user_id=user_id, folder_path=folder_path).first()
        if existing:
            return jsonify({'error': 'المجلد موجود بالفعل في المفضلة'}), 400
        
        favorite = Favorite(
            user_id=user_id,
            folder_path=folder_path,
            name=name or os.path.basename(folder_path)
        )
        
        db.session.add(favorite)
        db.session.commit()
        
        logger.info(f"Favorite added by user {user_id}: {folder_path}")
        
        return jsonify({
            'success': True,
            'message': 'تمت إضافة المجلد للمفضلة',
            'favorite': favorite.to_dict()
        })
        
    except Exception as e:
        db.session.rollback()
        logger.error(f"Add favorite error: {e}")
        return jsonify({'error': 'فشل إضافة المفضلة'}), 500


@app.route('/api/favorites/<int:favorite_id>', methods=['DELETE'])
@jwt_required()
def delete_favorite(favorite_id):
    """Delete favorite"""
    try:
        user_id = get_jwt_identity()
        favorite = Favorite.query.filter_by(id=favorite_id, user_id=user_id).first()
        
        if not favorite:
            return jsonify({'error': 'المفضلة غير موجودة'}), 404
        
        db.session.delete(favorite)
        db.session.commit()
        
        logger.info(f"Favorite deleted by user {user_id}: {favorite_id}")
        
        return jsonify({
            'success': True,
            'message': 'تم حذف المفضلة'
        })
        
    except Exception as e:
        db.session.rollback()
        logger.error(f"Delete favorite error: {e}")
        return jsonify({'error': 'فشل حذف المفضلة'}), 500


# ======================== PROCESSING LOGS ======================== #
@app.route('/api/logs', methods=['GET'])
@jwt_required()
def get_processing_logs():
    """Get user's processing logs"""
    try:
        user_id = get_jwt_identity()
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 20, type=int)
        
        logs = ProcessingLog.query.filter_by(user_id=user_id)\
            .order_by(ProcessingLog.started_at.desc())\
            .paginate(page=page, per_page=per_page, error_out=False)
        
        return jsonify({
            'success': True,
            'logs': [log.to_dict() for log in logs.items],
            'total': logs.total,
            'pages': logs.pages,
            'current_page': logs.page
        })
        
    except Exception as e:
        logger.error(f"Get logs error: {e}")
        return jsonify({'error': 'فشل جلب السجلات'}), 500


# ======================== CLEANUP ======================== #
@app.route('/api/cleanup', methods=['POST'])
@jwt_required()
def cleanup():
    """Clean up old files with detailed statistics"""
    try:
        user_id = get_jwt_identity()
        user = User.query.get(user_id)
        
        # Only admins can trigger cleanup
        if not user.is_admin:
            return jsonify({'error': 'غير مصرح'}), 403
        
        cleanup_time = datetime.now().timestamp() - (app.config['CLEANUP_AFTER_HOURS'] * 3600)
        
        cleanup_stats = {
            'uploads_deleted': 0,
            'folders_deleted': 0,
            'zips_deleted': 0,
            'total_space_freed': 0
        }
        
        # Clean upload folder
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            if os.path.isfile(file_path) and os.path.getctime(file_path) < cleanup_time:
                try:
                    file_size = os.path.getsize(file_path)
                    os.remove(file_path)
                    cleanup_stats['uploads_deleted'] += 1
                    cleanup_stats['total_space_freed'] += file_size
                except Exception as e:
                    logger.error(f"Error deleting file {filename}: {e}")
        
        # Clean output folders
        for folder in os.listdir(app.config['OUTPUT_FOLDER']):
            folder_path = os.path.join(app.config['OUTPUT_FOLDER'], folder)
            if os.path.isdir(folder_path) and os.path.getctime(folder_path) < cleanup_time:
                try:
                    folder_size = sum(
                        os.path.getsize(os.path.join(dirpath, filename))
                        for dirpath, dirnames, filenames in os.walk(folder_path)
                        for filename in filenames
                    )
                    shutil.rmtree(folder_path)
                    cleanup_stats['folders_deleted'] += 1
                    cleanup_stats['total_space_freed'] += folder_size
                except Exception as e:
                    logger.error(f"Error deleting folder {folder}: {e}")
        
        # Clean old ZIP files
        for filename in os.listdir(app.config['OUTPUT_FOLDER']):
            if filename.endswith('.zip'):
                file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
                if os.path.getctime(file_path) < cleanup_time:
                    try:
                        file_size = os.path.getsize(file_path)
                        os.remove(file_path)
                        cleanup_stats['zips_deleted'] += 1
                        cleanup_stats['total_space_freed'] += file_size
                    except Exception as e:
                        logger.error(f"Error deleting zip {filename}: {e}")
        
        # Convert bytes to MB
        cleanup_stats['space_freed_mb'] = round(cleanup_stats['total_space_freed'] / (1024 * 1024), 2)
        
        logger.info(f"Cleanup executed by user {user_id}: {cleanup_stats}")
        
        return jsonify({
            'success': True,
            'message': 'تم التنظيف بنجاح',
            'stats': cleanup_stats
        })
    
    except Exception as e:
        logger.error(f"Cleanup error: {e}")
        return jsonify({'error': 'فشل التنظيف'}), 500


# ======================== ERROR HANDLERS ======================== #
@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'الصفحة غير موجودة'}), 404


@app.errorhandler(500)
def internal_error(e):
    return jsonify({'error': 'خطأ داخلي في الخادم'}), 500


@app.errorhandler(429)
def ratelimit_handler(e):
    return jsonify({'error': 'تجاوزت الحد المسموح من الطلبات. يرجى المحاولة لاحقاً'}), 429


# ======================== DATABASE INITIALIZATION ======================== #
if __name__ == '__main__':
    with app.app_context():
        # Create all database tables
        db.create_all()
        logger.info("Database tables created successfully")
        logger.info(f"Database location: {DB_FILE}")
        
        # Log application startup
        logger.info("=" * 70)
        logger.info("Xtractor Pro v3.0 - Starting Application")
        logger.info(f"Environment: {env}")
        logger.info(f"Upload folder: {app.config['UPLOAD_FOLDER']}")
        logger.info(f"Output folder: {app.config['OUTPUT_FOLDER']}")
        logger.info("=" * 70)
    
    # Run the application
    port = int(os.environ.get("PORT", 8080))
    app.run(debug=(env == 'development'), host='0.0.0.0', port=port)
