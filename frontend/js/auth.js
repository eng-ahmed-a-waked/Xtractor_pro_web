// API Base URL
const API_URL = window.location.hostname === 'localhost' 
    ? 'http://localhost:5000/api' 
    : '/api';

// Authentication functions
class AuthService {
    constructor() {
        this.accessToken = localStorage.getItem('access_token');
        this.refreshToken = localStorage.getItem('refresh_token');
        this.user = JSON.parse(localStorage.getItem('user') || 'null');
        this.pendingUserId = localStorage.getItem('pending_user_id');
    }

    // Register new user
    async register(email, username, password, fullName, subscriptionPeriod = 'monthly') {
        try {
            const response = await fetch(`${API_URL}/auth/register`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    email,
                    username,
                    password,
                    full_name: fullName,
                    subscription_period: subscriptionPeriod
                })
            });

            const data = await response.json();

            if (data.success) {
                // Save user ID for activation
                if (data.user && data.user.id) {
                    localStorage.setItem('pending_user_id', data.user.id);
                }
                return { success: true, user: data.user, message: data.message };
            } else {
                return { success: false, error: data.error };
            }
        } catch (error) {
            console.error('Register error:', error);
            return { success: false, error: 'حدث خطأ في الاتصال' };
        }
    }

    // Activate account
    async activateAccount(userId, activationCode) {
        try {
            const response = await fetch(`${API_URL}/auth/activate`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    user_id: userId,
                    activation_code: activationCode
                })
            });

            const data = await response.json();

            if (data.success) {
                // Clear pending user
                localStorage.removeItem('pending_user_id');
                return { success: true, message: data.message };
            } else {
                return { success: false, error: data.error };
            }
        } catch (error) {
            console.error('Activation error:', error);
            return { success: false, error: 'حدث خطأ في الاتصال' };
        }
    }

    // Activate by email (convenience method)
    async activateByEmail(email, activationCode) {
        try {
            const response = await fetch(`${API_URL}/auth/activate-by-email`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    email,
                    activation_code: activationCode
                })
            });

            const data = await response.json();

            if (data.success) {
                return { success: true, message: data.message };
            } else {
                return { success: false, error: data.error };
            }
        } catch (error) {
            console.error('Activation error:', error);
            return { success: false, error: 'حدث خطأ في الاتصال' };
        }
    }

    // Resend activation code
    async resendActivationCode(email) {
        try {
            const response = await fetch(`${API_URL}/auth/resend-activation`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ email })
            });

            const data = await response.json();

            if (data.success) {
                return { success: true, message: data.message };
            } else {
                return { success: false, error: data.error };
            }
        } catch (error) {
            console.error('Resend activation error:', error);
            return { success: false, error: 'حدث خطأ في الاتصال' };
        }
    }

    // Login user
    async login(email, password) {
        try {
            const response = await fetch(`${API_URL}/auth/login`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ email, password })
            });

            const data = await response.json();

            if (data.success) {
                // Save tokens and user
                this.saveAuth(data.access_token, data.refresh_token, data.user);
                return { success: true, user: data.user, subscription: data.subscription };
            } else {
                // Check if requires activation
                if (data.requires_activation) {
                    return { 
                        success: false, 
                        error: data.error,
                        requires_activation: true,
                        user_id: data.user_id
                    };
                }
                return { success: false, error: data.error };
            }
        } catch (error) {
            console.error('Login error:', error);
            return { success: false, error: 'حدث خطأ في الاتصال' };
        }
    }

    // Logout
    logout() {
        localStorage.removeItem('access_token');
        localStorage.removeItem('refresh_token');
        localStorage.removeItem('user');
        localStorage.removeItem('pending_user_id');
        this.accessToken = null;
        this.refreshToken = null;
        this.user = null;
        this.pendingUserId = null;
        window.location.href = '/login';
    }

    // Save authentication data
    saveAuth(accessToken, refreshToken, user) {
        localStorage.setItem('access_token', accessToken);
        localStorage.setItem('refresh_token', refreshToken);
        localStorage.setItem('user', JSON.stringify(user));
        this.accessToken = accessToken;
        this.refreshToken = refreshToken;
        this.user = user;
    }

    // Check if user is authenticated
    isAuthenticated() {
        return !!this.accessToken && !!this.user;
    }

    // Get current user
    getUser() {
        return this.user;
    }

    // Get access token
    getToken() {
        return this.accessToken;
    }

    // Refresh access token
    async refreshAccessToken() {
        try {
            const response = await fetch(`${API_URL}/auth/refresh`, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${this.refreshToken}`
                }
            });

            const data = await response.json();

            if (data.success) {
                this.accessToken = data.access_token;
                localStorage.setItem('access_token', data.access_token);
                return true;
            } else {
                this.logout();
                return false;
            }
        } catch (error) {
            console.error('Token refresh error:', error);
            this.logout();
            return false;
        }
    }

    // Make authenticated API call
    async authenticatedFetch(url, options = {}) {
        if (!this.isAuthenticated()) {
            throw new Error('Not authenticated');
        }

        const headers = {
            ...options.headers,
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json'
        };

        try {
            let response = await fetch(url, { ...options, headers });

            // If token expired, refresh and retry
            if (response.status === 401) {
                const refreshed = await this.refreshAccessToken();
                if (refreshed) {
                    headers['Authorization'] = `Bearer ${this.accessToken}`;
                    response = await fetch(url, { ...options, headers });
                } else {
                    throw new Error('Session expired');
                }
            }

            return response;
        } catch (error) {
            console.error('Authenticated fetch error:', error);
            throw error;
        }
    }

    // Get subscription status
    async getSubscriptionStatus() {
        try {
            const response = await this.authenticatedFetch(`${API_URL}/subscription/status`);
            const data = await response.json();
            return data;
        } catch (error) {
            console.error('Get subscription error:', error);
            return { success: false, error: 'فشل جلب حالة الاشتراك' };
        }
    }

    // Request activation code
    async requestActivationCode() {
        try {
            const response = await this.authenticatedFetch(`${API_URL}/subscription/request-code`, {
                method: 'POST'
            });
            const data = await response.json();
            return data;
        } catch (error) {
            console.error('Request code error:', error);
            return { success: false, error: 'فشل طلب كود التفعيل' };
        }
    }

    // Renew subscription
    async renewSubscription(activationCode, subscriptionPeriod = 'monthly') {
        try {
            const response = await this.authenticatedFetch(`${API_URL}/subscription/renew`, {
                method: 'POST',
                body: JSON.stringify({ 
                    activation_code: activationCode,
                    subscription_period: subscriptionPeriod
                })
            });
            const data = await response.json();
            return data;
        } catch (error) {
            console.error('Renewal error:', error);
            return { success: false, error: 'فشل تجديد الاشتراك' };
        }
    }
}

// Create global auth service instance
const authService = new AuthService();

// Export functions for backward compatibility
async function register(email, username, password, fullName, subscriptionPeriod = 'monthly') {
    return await authService.register(email, username, password, fullName, subscriptionPeriod);
}

async function login(email, password) {
    return await authService.login(email, password);
}

function logout() {
    authService.logout();
}

function isAuthenticated() {
    return authService.isAuthenticated();
}

function getUser() {
    return authService.getUser();
}

function getToken() {
    return authService.getToken();
}

async function activateAccount(email, activationCode) {
    return await authService.activateByEmail(email, activationCode);
}

async function resendActivationCode(email) {
    return await authService.resendActivationCode(email);
}

// Check authentication on page load
document.addEventListener('DOMContentLoaded', () => {
    const currentPage = window.location.pathname;
    
    // Public pages that don't require authentication
    const publicPages = ['/login', '/register', '/activate'];
    const isPublicPage = publicPages.some(page => currentPage.startsWith(page));
    
    // Redirect to login if not authenticated (except on public pages)
    if (!isAuthenticated() && !isPublicPage) {
        window.location.href = '/login';
    }
    
    // Redirect to home if authenticated and on auth pages (except activate)
    if (isAuthenticated() && (currentPage === '/login' || currentPage === '/register')) {
        window.location.href = '/';
    }
});

// Add request interceptor for handling 401 errors globally
window.addEventListener('unhandledrejection', (event) => {
    if (event.reason && event.reason.message === 'Session expired') {
        authService.logout();
    }
});