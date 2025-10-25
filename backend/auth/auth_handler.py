import secrets
import string
import hashlib
from datetime import datetime, timedelta
from flask import current_app
from models.user import db, User, Subscription
import logging

logger = logging.getLogger(__name__)


class AuthHandler:
    """Handle authentication and subscription logic"""
    
    @staticmethod
    def generate_activation_code():
        """Generate a random activation code"""
        characters = string.ascii_uppercase + string.digits
        code = ''.join(secrets.choice(characters) for _ in range(12))
        return f"{code[:4]}-{code[4:8]}-{code[8:]}"
    
    @staticmethod
    def hash_code(code):
        """Hash activation code using SHA-256"""
        normalized = str(code).strip().replace("-", "").upper()
        return hashlib.sha256(normalized.encode()).hexdigest()
    
    @staticmethod
    def register_user(email, username, password, full_name=None, subscription_period='monthly'):
        """Register a new user with subscription period selection"""
        try:
            # Check if user exists
            if User.query.filter_by(email=email).first():
                return {'success': False, 'error': 'البريد الإلكتروني مسجل بالفعل'}
            
            if User.query.filter_by(username=username).first():
                return {'success': False, 'error': 'اسم المستخدم موجود بالفعل'}
            
            # Validate subscription period
            if subscription_period not in ['monthly', 'yearly']:
                return {'success': False, 'error': 'مدة الاشتراك غير صالحة'}
            
            # Create user (inactive until activation)
            user = User(
                email=email,
                username=username,
                password=password,
                full_name=full_name
            )
            user.is_active = False  # User must activate first
            
            db.session.add(user)
            db.session.commit()
            
            # Generate activation code
            activation_code = AuthHandler.generate_activation_code()
            code_hash = AuthHandler.hash_code(activation_code)
            
            # Create inactive subscription with selected period
            duration = timedelta(days=30) if subscription_period == 'monthly' else timedelta(days=365)
            
            subscription = Subscription(
                user_id=user.id,
                activation_code_hash=code_hash,
                subscription_period=subscription_period,
                is_trial=False,
                is_cancelled=False
            )
            
            db.session.add(subscription)
            db.session.commit()
            
            # Send activation code via email
            from auth.email_sender import EmailSender
            email_sent = EmailSender.send_activation_code(
                recipient_email=email,
                recipient_name=full_name or username,
                activation_code=activation_code,
                subscription_period=subscription_period
            )
            
            if not email_sent:
                logger.error(f"Failed to send activation email to {email}")
                # Don't fail registration, user can request code again
            
            logger.info(f"New user registered: {username} (period: {subscription_period})")
            
            return {
                'success': True,
                'user': user.to_dict(),
                'message': f'تم التسجيل بنجاح! تم إرسال كود التفعيل إلى {email}',
                'requires_activation': True
            }
            
        except Exception as e:
            db.session.rollback()
            logger.error(f"Registration error: {e}")
            return {'success': False, 'error': str(e)}
    
    @staticmethod
    def login_user(email, password):
        """Login user and return user info"""
        try:
            user = User.query.filter_by(email=email).first()
            
            if not user:
                return {'success': False, 'error': 'البريد الإلكتروني أو كلمة المرور غير صحيحة'}
            
            if not user.check_password(password):
                return {'success': False, 'error': 'البريد الإلكتروني أو كلمة المرور غير صحيحة'}
            
            if not user.is_active:
                return {
                    'success': False, 
                    'error': 'الحساب غير مفعل. يرجى تفعيل حسابك أولاً',
                    'requires_activation': True,
                    'user_id': user.id
                }
            
            # Update last login
            user.update_last_login()
            
            # Check subscription
            subscription_status = user.get_subscription_status()
            
            if not subscription_status['active']:
                return {
                    'success': False,
                    'error': 'انتهت صلاحية الاشتراك. يرجى تجديد اشتراكك',
                    'requires_renewal': True,
                    'user_id': user.id
                }
            
            logger.info(f"User logged in: {user.username}")
            
            return {
                'success': True,
                'user': user.to_dict(),
                'subscription': subscription_status
            }
            
        except Exception as e:
            logger.error(f"Login error: {e}")
            return {'success': False, 'error': 'حدث خطأ أثناء تسجيل الدخول'}
    
    @staticmethod
    def activate_subscription(user_id, activation_code):
        """Activate user subscription with code"""
        try:
            user = User.query.get(user_id)
            if not user:
                return {'success': False, 'error': 'المستخدم غير موجود'}
            
            subscription = user.subscription
            if not subscription:
                return {'success': False, 'error': 'لا يوجد اشتراك لهذا المستخدم'}
            
            # Check if already activated
            if subscription.activated_at and subscription.is_active():
                return {'success': False, 'error': 'الاشتراك مفعل بالفعل'}
            
            # Verify activation code
            code_hash = AuthHandler.hash_code(activation_code)
            
            if subscription.activation_code_hash != code_hash:
                return {'success': False, 'error': 'كود التفعيل غير صحيح'}
            
            # Activate subscription
            subscription.activated_at = datetime.utcnow()
            
            # Set expiration based on subscription period
            if subscription.subscription_period == 'monthly':
                subscription.expires_at = datetime.utcnow() + timedelta(days=30)
            else:  # yearly
                subscription.expires_at = datetime.utcnow() + timedelta(days=365)
            
            subscription.is_cancelled = False
            
            # Activate user account
            user.is_active = True
            
            db.session.commit()
            
            logger.info(f"Subscription activated for user: {user.username}")
            
            return {
                'success': True,
                'message': 'تم تفعيل الاشتراك بنجاح!',
                'subscription': subscription.get_status()
            }
            
        except Exception as e:
            db.session.rollback()
            logger.error(f"Activation error: {e}")
            return {'success': False, 'error': 'فشل تفعيل الاشتراك'}
    
    @staticmethod
    def check_subscription(user_id):
        """Check if user has active subscription"""
        try:
            user = User.query.get(user_id)
            if not user:
                return {'success': False, 'error': 'المستخدم غير موجود'}
            
            status = user.get_subscription_status()
            
            return {
                'success': True,
                'subscription': status
            }
            
        except Exception as e:
            logger.error(f"Subscription check error: {e}")
            return {'success': False, 'error': str(e)}
    
    @staticmethod
    def renew_subscription(user_id, activation_code, subscription_period='monthly'):
        """Renew expired subscription with new code"""
        try:
            user = User.query.get(user_id)
            if not user or not user.subscription:
                return {'success': False, 'error': 'المستخدم أو الاشتراك غير موجود'}
            
            subscription = user.subscription
            
            # Hash and verify new code
            code_hash = AuthHandler.hash_code(activation_code)
            
            # For renewal, we accept the code and update
            subscription.activation_code_hash = code_hash
            subscription.subscription_period = subscription_period
            
            # Extend subscription from now
            subscription.activated_at = datetime.utcnow()
            
            if subscription_period == 'monthly':
                subscription.expires_at = datetime.utcnow() + timedelta(days=30)
            else:  # yearly
                subscription.expires_at = datetime.utcnow() + timedelta(days=365)
            
            subscription.is_cancelled = False
            
            db.session.commit()
            
            logger.info(f"Subscription renewed for user: {user.username} (period: {subscription_period})")
            
            return {
                'success': True,
                'message': 'تم تجديد الاشتراك بنجاح!',
                'subscription': subscription.get_status()
            }
            
        except Exception as e:
            db.session.rollback()
            logger.error(f"Renewal error: {e}")
            return {'success': False, 'error': 'فشل تجديد الاشتراك'}
    
    @staticmethod
    def request_activation_code(user_id, subscription_period='monthly'):
        """Generate and send new activation code to user"""
        try:
            user = User.query.get(user_id)
            if not user:
                return {'success': False, 'error': 'المستخدم غير موجود'}
            
            # Generate new code
            code = AuthHandler.generate_activation_code()
            
            # Send email
            from auth.email_sender import EmailSender
            email_sent = EmailSender.send_activation_code(
                user.email, 
                user.full_name or user.username, 
                code,
                subscription_period
            )
            
            if not email_sent:
                return {'success': False, 'error': 'فشل إرسال البريد الإلكتروني'}
            
            logger.info(f"Activation code sent to: {user.email}")
            
            return {
                'success': True,
                'message': f'تم إرسال كود التفعيل إلى {user.email}'
            }
            
        except Exception as e:
            logger.error(f"Request activation code error: {e}")
            return {'success': False, 'error': str(e)}
