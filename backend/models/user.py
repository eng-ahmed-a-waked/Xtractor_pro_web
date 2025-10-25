from datetime import datetime, timedelta
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(120), unique=True, nullable=False, index=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(120))
    
    # Account status
    is_active = db.Column(db.Boolean, default=True)
    is_admin = db.Column(db.Boolean, default=False)
    email_verified = db.Column(db.Boolean, default=False)
    
    # Timestamps
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_login = db.Column(db.DateTime)
    
    # Relationships
    subscription = db.relationship('Subscription', backref='user', lazy=True, uselist=False)
    favorites = db.relationship('Favorite', backref='user', lazy=True)
    processing_logs = db.relationship('ProcessingLog', backref='user', lazy=True)
    
    def __init__(self, email, username, password, full_name=None):
        self.email = email
        self.username = username
        self.set_password(password)
        self.full_name = full_name
    
    def set_password(self, password):
        """Hash and set password"""
        self.password_hash = generate_password_hash(password)
    
    def check_password(self, password):
        """Verify password"""
        return check_password_hash(self.password_hash, password)
    
    def update_last_login(self):
        """Update last login timestamp"""
        self.last_login = datetime.utcnow()
        db.session.commit()
    
    def has_active_subscription(self):
        """Check if user has active subscription"""
        if not self.subscription:
            return False
        return self.subscription.is_active()
    
    def get_subscription_status(self):
        """Get detailed subscription status"""
        if not self.subscription:
            return {
                'active': False,
                'message': 'No subscription found',
                'hours_left': 0
            }
        return self.subscription.get_status()
    
    def to_dict(self):
        """Convert user to dictionary"""
        return {
            'id': self.id,
            'email': self.email,
            'username': self.username,
            'full_name': self.full_name,
            'is_active': self.is_active,
            'is_admin': self.is_admin,
            'email_verified': self.email_verified,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'last_login': self.last_login.isoformat() if self.last_login else None,
            'subscription': self.get_subscription_status()
        }
    
    def __repr__(self):
        return f'<User {self.username}>'


class Subscription(db.Model):
    __tablename__ = 'subscriptions'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False, unique=True)
    
    # Activation details
    activation_code_hash = db.Column(db.String(255))
    activated_at = db.Column(db.DateTime)
    expires_at = db.Column(db.DateTime)
    
    # Subscription period - NEW FIELD
    subscription_period = db.Column(db.String(20), default='monthly')  # 'monthly' or 'yearly'
    
    # Trial
    is_trial = db.Column(db.Boolean, default=False)
    
    # Status
    is_cancelled = db.Column(db.Boolean, default=False)
    cancelled_at = db.Column(db.DateTime)
    
    def is_active(self):
        """Check if subscription is currently active"""
        if self.is_cancelled:
            return False
        if not self.expires_at:
            return False
        return datetime.utcnow() < self.expires_at
    
    def get_remaining_time(self):
        """Get remaining subscription time"""
        if not self.is_active():
            return timedelta(0)
        return self.expires_at - datetime.utcnow()
    
    def get_status(self):
        """Get detailed subscription status"""
        if not self.is_active():
            return {
                'active': False,
                'message': 'انتهت صلاحية الاشتراك',
                'hours_left': 0,
                'expires_at': self.expires_at.isoformat() if self.expires_at else None,
                'subscription_period': self.subscription_period
            }
        
        remaining = self.get_remaining_time()
        hours = int(remaining.total_seconds() / 3600)
        minutes = int((remaining.total_seconds() % 3600) / 60)
        
        return {
            'active': True,
            'message': 'الاشتراك نشط',
            'hours_left': hours,
            'minutes_left': minutes,
            'expires_at': self.expires_at.isoformat(),
            'is_trial': self.is_trial,
            'subscription_period': self.subscription_period
        }
    
    def __repr__(self):
        return f'<Subscription user_id={self.user_id}>'


class Favorite(db.Model):
    __tablename__ = 'favorites'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    folder_path = db.Column(db.String(500), nullable=False)
    name = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    def to_dict(self):
        return {
            'id': self.id,
            'folder_path': self.folder_path,
            'name': self.name,
            'created_at': self.created_at.isoformat()
        }
    
    def __repr__(self):
        return f'<Favorite {self.name}>'


class ProcessingLog(db.Model):
    __tablename__ = 'processing_logs'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    
    # Processing details
    job_type = db.Column(db.String(50))  # 'cars' or 'visits'
    mode = db.Column(db.String(50))      # 'engine_idle' or 'parking_details'
    filename = db.Column(db.String(300))
    status = db.Column(db.String(50))    # 'processing', 'completed', 'failed'
    
    # Statistics
    records_processed = db.Column(db.Integer, default=0)
    inside_zone = db.Column(db.Integer, default=0)
    outside_zone = db.Column(db.Integer, default=0)
    undefined_zone = db.Column(db.Integer, default=0)
    
    # Timing
    started_at = db.Column(db.DateTime, default=datetime.utcnow)
    completed_at = db.Column(db.DateTime)
    
    # Error handling
    error_message = db.Column(db.Text)
    
    def to_dict(self):
        return {
            'id': self.id,
            'job_type': self.job_type,
            'mode': self.mode,
            'filename': self.filename,
            'status': self.status,
            'records_processed': self.records_processed,
            'statistics': {
                'inside_zone': self.inside_zone,
                'outside_zone': self.outside_zone,
                'undefined_zone': self.undefined_zone
            },
            'started_at': self.started_at.isoformat(),
            'completed_at': self.completed_at.isoformat() if self.completed_at else None
        }
    
    def __repr__(self):
        return f'<ProcessingLog {self.id} - {self.status}>'
