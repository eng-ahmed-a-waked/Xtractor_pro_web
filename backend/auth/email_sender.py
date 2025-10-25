from flask import current_app
from flask_mail import Mail, Message
from datetime import datetime
import logging

logger = logging.getLogger(__name__)
mail = Mail()


class EmailSender:
    """Handle email sending for activation codes and notifications"""
    
    @staticmethod
    def send_activation_code(recipient_email, recipient_name, activation_code, subscription_period='monthly'):
        """
        Send activation code email to ADMIN only
        
        Args:
            recipient_email: User's email (for reference only, NOT the email recipient)
            recipient_name: User's name
            activation_code: The activation code
            subscription_period: Subscription period (monthly/yearly)
            
        Returns:
            bool: Success status
        """
        # تحويل المتغير للتوافق مع باقي الكود
        user_email = recipient_email
        try:
            # إرسال الكود للأدمن فقط - وليس للمستخدم
            admin_email = current_app.config.get('ADMIN_EMAIL', 'eng.ahmed.a.waked@gmail.com')
            
            period_text = "شهر واحد (30 يوم)" if subscription_period == 'monthly' else "سنة كاملة (365 يوم)"
            
            msg = Message(
                subject=f'🔐 كود تفعيل جديد - {recipient_name}',
                recipients=[admin_email],  # إرسال للأدمن فقط
                sender=current_app.config['MAIL_DEFAULT_SENDER']
            )
            
            # HTML email body
            msg.html = f"""
            <html dir="rtl">
                <body style="font-family: 'Segoe UI', Arial, sans-serif; direction: rtl; text-align: right; background-color: #f5f5f5; padding: 20px;">
                    <div style="max-width: 600px; margin: 0 auto; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.2);">
                        
                        <!-- Header -->
                        <div style="text-align: center; margin-bottom: 30px;">
                            <h1 style="color: white; font-size: 32px; margin: 0;">🌊 Xtractor Pro</h1>
                            <p style="color: #E8F4F8; font-size: 14px; margin: 5px 0;">كود تفعيل لمستخدم جديد</p>
                        </div>
                        
                        <!-- Main Content -->
                        <div style="background: white; padding: 40px 30px; border-radius: 10px; margin: 20px 0;">
                            <h2 style="color: #333; text-align: center; font-size: 24px; margin-bottom: 20px;">
                                📋 معلومات المستخدم الجديد
                            </h2>
                            
                            <!-- User Info -->
                            <div style="background: #e3f2fd; padding: 20px; border-radius: 8px; border-right: 4px solid #2196f3; margin: 20px 0;">
                                <h3 style="color: #1976d2; margin-top: 0; font-size: 16px;">👤 بيانات المستخدم:</h3>
                                <p style="font-size: 14px; color: #555; margin: 5px 0;">
                                    <strong>الاسم:</strong> {recipient_name}
                                </p>
                                <p style="font-size: 14px; color: #555; margin: 5px 0;">
                                    <strong>البريد الإلكتروني:</strong> {user_email}
                                </p>
                                <p style="font-size: 14px; color: #555; margin: 5px 0;">
                                    <strong>مدة الاشتراك:</strong> {period_text}
                                </p>
                                <p style="font-size: 14px; color: #555; margin: 5px 0;">
                                    <strong>تاريخ التسجيل:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
                                </p>
                            </div>
                            
                            <!-- Activation Code Box -->
                            <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 25px; border-radius: 10px; text-align: center; margin: 30px 0; box-shadow: 0 5px 15px rgba(102, 126, 234, 0.3);">
                                <p style="color: white; font-size: 14px; margin: 0 0 10px 0;">كود التفعيل:</p>
                                <h1 style="color: white; font-size: 36px; letter-spacing: 5px; margin: 0; font-family: 'Courier New', monospace; text-shadow: 2px 2px 4px rgba(0,0,0,0.2);">
                                    {activation_code}
                                </h1>
                            </div>
                            
                            <!-- Instructions -->
                            <div style="margin-top: 30px; padding: 20px; background: #fff3cd; border-radius: 8px; border-right: 4px solid #ffc107;">
                                <h3 style="color: #856404; margin-top: 0; font-size: 16px;">💡 تعليمات:</h3>
                                <ul style="color: #856404; font-size: 14px; margin: 10px 0; padding-right: 20px;">
                                    <li>قم بإرسال هذا الكود للمستخدم: <strong>{user_email}</strong></li>
                                    <li>الكود صالح لمرة واحدة فقط</li>
                                    <li>بعد التفعيل، سيبدأ الاشتراك تلقائياً</li>
                                    <li>مدة الاشتراك: {period_text}</li>
                                </ul>
                            </div>
                        </div>
                        
                        <!-- Footer -->
                        <div style="text-align: center; margin-top: 20px;">
                            <p style="color: rgba(255,255,255,0.7); font-size: 12px; margin: 5px 0;">
                                Xtractor Pro v3.0 Ocean Edition
                            </p>
                            <p style="color: rgba(255,255,255,0.7); font-size: 11px; margin: 5px 0;">
                                © 2024 All rights reserved
                            </p>
                        </div>
                    </div>
                </body>
            </html>
            """
            
            # Plain text fallback
            msg.body = f"""
            طلب تفعيل حساب جديد - Xtractor Pro
            
            معلومات المستخدم:
            الاسم: {recipient_name}
            البريد الإلكتروني: {user_email}
            مدة الاشتراك: {period_text}
            تاريخ التسجيل: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            
            كود التفعيل: {activation_code}
            
            يرجى إرسال هذا الكود إلى المستخدم على البريد: {user_email}
            
            Xtractor Pro - Eng. Ahmed Waked
            """
            
            mail.send(msg)
            logger.info(f"Activation code sent to ADMIN for user: {user_email}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to send activation email: {e}")
            return False
    
    @staticmethod
    def send_welcome_email(recipient_email, recipient_name, subscription_period='monthly'):
        """Send welcome email after successful activation (optional)"""
        try:
            period_text = "شهر واحد" if subscription_period == 'monthly' else "سنة كاملة"
            
            msg = Message(
                subject='🎉 مرحباً بك في Xtractor Pro',
                recipients=[recipient_email],
                sender=current_app.config['MAIL_DEFAULT_SENDER']
            )
            
            msg.html = f"""
            <html dir="rtl">
                <body style="font-family: 'Segoe UI', Arial, sans-serif; direction: rtl; text-align: right; background-color: #f5f5f5; padding: 20px;">
                    <div style="max-width: 600px; margin: 0 auto; background: linear-gradient(135deg, #0A2463 0%, #3CAEA3 100%); padding: 30px; border-radius: 15px;">
                        <h1 style="color: white; text-align: center;">🌊 مرحباً بك في Xtractor Pro</h1>
                        
                        <div style="background: white; padding: 30px; border-radius: 10px; margin: 20px 0;">
                            <h2 style="color: #0A2463;">مرحباً {recipient_name}! 👋</h2>
                            
                            <p style="color: #555; font-size: 16px; line-height: 1.6;">
                                تم تفعيل اشتراكك بنجاح في <strong>Xtractor Pro</strong>!
                            </p>
                            
                            <div style="background: #E8F4F8; padding: 20px; border-radius: 8px; margin: 20px 0;">
                                <h3 style="color: #0A2463; margin-top: 0;">✅ تفاصيل الاشتراك</h3>
                                <p style="color: #555;">
                                    مدة الاشتراك: <strong>{period_text}</strong><br>
                                    تاريخ التفعيل: <strong>{datetime.now().strftime('%Y-%m-%d')}</strong>
                                </p>
                            </div>
                            
                            <h3 style="color: #0A2463;">✨ المزايا المتاحة:</h3>
                            <ul style="color: #555; line-height: 1.8;">
                                <li>✅ تحويل ملفات XLS إلى XLSX</li>
                                <li>✅ استخلاص بيانات السيارات</li>
                                <li>✅ توزيع تقارير المبيعات</li>
                                <li>✅ إنشاء تقارير موحدة</li>
                                <li>✅ فحص المناطق الجغرافية</li>
                            </ul>
                            
                            <div style="text-align: center; margin: 30px 0;">
                                <a href="{current_app.config.get('FRONTEND_URL', 'http://localhost:5000')}" 
                                   style="display: inline-block; background: linear-gradient(135deg, #3CAEA3, #158FAD); color: white; padding: 15px 40px; text-decoration: none; border-radius: 25px; font-weight: bold; box-shadow: 0 5px 15px rgba(60, 174, 163, 0.3);">
                                    🚀 ابدأ الآن
                                </a>
                            </div>
                        </div>
                        
                        <p style="color: white; text-align: center; font-size: 12px; margin-top: 20px;">
                            Developed by Eng. Ahmed Waked<br>
                            v3.0 Ocean Edition
                        </p>
                    </div>
                </body>
            </html>
            """
            
            mail.send(msg)
            logger.info(f"Welcome email sent to: {recipient_email}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to send welcome email: {e}")
            return False
    
    @staticmethod
    def send_subscription_expiry_warning(recipient_email, recipient_name, days_left):
        """Send warning email when subscription is about to expire"""
        try:
            msg = Message(
                subject='⚠️ تنبيه: اشتراكك على وشك الانتهاء',
                recipients=[recipient_email],
                sender=current_app.config['MAIL_DEFAULT_SENDER']
            )
            
            msg.html = f"""
            <html dir="rtl">
                <body style="font-family: 'Segoe UI', Arial, sans-serif; direction: rtl; text-align: right; background-color: #f5f5f5; padding: 20px;">
                    <div style="max-width: 600px; margin: 0 auto; background: linear-gradient(135deg, #F39C12 0%, #E74C3C 100%); padding: 30px; border-radius: 15px;">
                        <h1 style="color: white; text-align: center;">⏰ تنبيه الاشتراك</h1>
                        
                        <div style="background: white; padding: 30px; border-radius: 10px; margin: 20px 0;">
                            <h2 style="color: #E74C3C;">مرحباً {recipient_name}!</h2>
                            
                            <p style="color: #555; font-size: 16px;">
                                اشتراكك في Xtractor Pro على وشك الانتهاء.
                            </p>
                            
                            <div style="background: #FFF3CD; padding: 20px; border-radius: 8px; border-right: 4px solid #F39C12; margin: 20px 0;">
                                <p style="color: #856404; font-size: 18px; margin: 0;">
                                    ⏳ الوقت المتبقي: <strong>{days_left} يوم</strong>
                                </p>
                            </div>
                            
                            <p style="color: #555;">
                                لتجديد اشتراكك والاستمرار في استخدام جميع المزايا، يرجى التواصل معنا للحصول على كود تفعيل جديد.
                            </p>
                            
                            <div style="text-align: center; margin: 30px 0;">
                                <p style="color: #555; font-size: 14px;">
                                    هل تحتاج مساعدة؟ تواصل معنا:<br>
                                    📧 eng.ahmed.a.waked@gmail.com
                                </p>
                            </div>
                        </div>
                    </div>
                </body>
            </html>
            """
            
            mail.send(msg)
            logger.info(f"Expiry warning sent to: {recipient_email}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to send expiry warning: {e}")
            return False
