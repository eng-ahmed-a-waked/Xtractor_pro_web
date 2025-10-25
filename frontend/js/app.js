// API Base URL
const API_URL = window.location.hostname === 'localhost' 
    ? 'http://localhost:5000/api' 
    : '/api';

// Global state
let currentMode = 'engine_idle';
let selectedFiles = [];
let currentJobId = null;

console.log('🔧 App.js loaded successfully');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Application initializing...');
    
    // التحقق من وجود جميع العناصر المطلوبة
    const requiredElements = {
        'filesInput': document.getElementById('filesInput'),
        'selectFilesBtn': document.getElementById('selectFilesBtn'),
        'fileCountDisplay': document.getElementById('fileCountDisplay'),
        'carProcessForm': document.getElementById('carProcessForm'),
        'visitsProcessForm': document.getElementById('visitsProcessForm'),
        'currentTime': document.getElementById('currentTime')
    };
    
    console.log('🔍 Checking required elements:');
    Object.entries(requiredElements).forEach(([name, element]) => {
        console.log(`  ${element ? '✅' : '❌'} ${name}`);
    });
    
    initializeClock();
    initializeMenuButtons();
    initializeModeButtons();
    initializeFileInput();
    initializeCarForm();
    initializeVisitsForm();
    
    // Check authentication
    if (!isAuthenticated()) {
        console.warn('⚠️ User not authenticated, redirecting to login...');
        window.location.href = '/login';
        return;
    }
    
    // Display user info
    displayUserInfo();
    
    console.log('✅ Application initialized successfully');
});

// Emergency initialization fallback
window.addEventListener('load', () => {
    console.log('🔄 Window fully loaded - Running emergency checks...');
    
    const fileInput = document.getElementById('filesInput');
    const selectBtn = document.getElementById('selectFilesBtn');
    
    if (fileInput && selectBtn && !selectBtn.onclick) {
        console.log('🚨 Running emergency file input initialization...');
        
        selectBtn.onclick = (e) => {
            e.preventDefault();
            e.stopPropagation();
            console.log('🖱️ Emergency button click handler');
            fileInput.click();
        };
        
        fileInput.onchange = (e) => {
            console.log('📂 Emergency file change handler - Files:', e.target.files.length);
            selectedFiles = Array.from(e.target.files);
            updateFileCountDisplay();
            
            selectedFiles.forEach((file, index) => {
                const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
                console.log(`  File ${index + 1}: ${file.name} (${sizeMB} MB)`);
            });
        };
        
        console.log('✅ Emergency initialization complete');
    }
});

// ==================== CLOCK ====================
function initializeClock() {
    function updateTime() {
        const now = new Date();
        const timeString = now.toLocaleString('ar-EG', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            hour12: false
        });
        const clockEl = document.getElementById('currentTime');
        if (clockEl) {
            clockEl.textContent = `📅 ${timeString}`;
        }
    }
    updateTime();
    setInterval(updateTime, 1000);
    console.log('✅ Clock initialized');
}

// ==================== MENU NAVIGATION ====================
function initializeMenuButtons() {
    const menuButtons = document.querySelectorAll('.menu-btn');
    const sections = document.querySelectorAll('.section');
    
    console.log(`📋 Found ${menuButtons.length} menu buttons and ${sections.length} sections`);
    
    menuButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetSection = btn.dataset.section;
            console.log(`📂 Switching to section: ${targetSection}`);
            
            menuButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            sections.forEach(section => {
                section.classList.toggle('active', section.id === targetSection);
            });
            
            hideAllProgressCards();
        });
    });
    
    console.log('✅ Menu buttons initialized');
}

function hideAllProgressCards() {
    const cards = [
        'carProgressCard', 
        'resultsCard', 
        'visitsProgressCard', 
        'visitsResultsCard'
    ];
    cards.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
    });
}

// ==================== MODE BUTTONS ====================
function initializeModeButtons() {
    const modeButtons = document.querySelectorAll('.mode-btn');
    
    console.log(`⚙️ Found ${modeButtons.length} mode buttons`);
    
    modeButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            currentMode = btn.dataset.mode;
            console.log(`🔧 Mode changed to: ${currentMode}`);
            
            modeButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            const modeText = currentMode === 'engine_idle' 
                ? 'وضع: محرك في وضع الخمول (Engine Idle)'
                : 'وضع: تفاصيل أماكن التوقف (Parking Details)';
            
            const displayEl = document.getElementById('selectedModeDisplay');
            if (displayEl) {
                displayEl.textContent = modeText;
            }
        });
    });
    
    console.log('✅ Mode buttons initialized');
}

// ==================== FILE INPUT ====================
function initializeFileInput() {
    const fileInput = document.getElementById('filesInput');
    const selectBtn = document.getElementById('selectFilesBtn');
    const displayEl = document.getElementById('fileCountDisplay');
    
    console.log('📁 Initializing file input...');
    console.log('File input elements:', {
        fileInput: fileInput ? '✅ Found' : '❌ NOT FOUND',
        selectBtn: selectBtn ? '✅ Found' : '❌ NOT FOUND',
        displayEl: displayEl ? '✅ Found' : '❌ NOT FOUND'
    });
    
    if (!fileInput) {
        console.error('❌ File input element not found!');
        return;
    }
    
    if (!selectBtn) {
        console.error('❌ Select button not found!');
        return;
    }
    
    if (!displayEl) {
        console.error('❌ Display element not found!');
        return;
    }
    
    // Add hover effects
    selectBtn.addEventListener('mouseenter', () => {
        selectBtn.style.background = 'rgba(255,255,255,0.35)';
        selectBtn.style.borderStyle = 'solid';
    });
    
    selectBtn.addEventListener('mouseleave', () => {
        selectBtn.style.background = 'rgba(255,255,255,0.25)';
        selectBtn.style.borderStyle = 'dashed';
    });
    
    // Open file dialog when button clicked
    selectBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        console.log('📂 Select button clicked - opening file dialog');
        fileInput.click();
    });
    
    // Handle file selection
    fileInput.addEventListener('change', (e) => {
        console.log('📄 Files changed event triggered');
        console.log('Files selected:', e.target.files.length);
        
        selectedFiles = Array.from(e.target.files);
        updateFileCountDisplay();
        
        // Log selected files
        selectedFiles.forEach((file, index) => {
            const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
            console.log(`File ${index + 1}: ${file.name} (${sizeMB} MB)`);
        });
    });
    
    console.log('✅ File input initialized successfully');
}

function updateFileCountDisplay() {
    const displayEl = document.getElementById('fileCountDisplay');
    if (!displayEl) {
        console.error('❌ Display element not found in updateFileCountDisplay');
        return;
    }
    
    if (selectedFiles.length > 0) {
        console.log(`✅ Updating display for ${selectedFiles.length} files`);
        
        // Calculate total size
        const totalSize = selectedFiles.reduce((sum, file) => sum + file.size, 0);
        const totalSizeMB = (totalSize / (1024 * 1024)).toFixed(2);
        
        // Create file list HTML
        let filesHtml = '<div style="margin-top: 10px;">';
        selectedFiles.forEach((file, index) => {
            const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
            const ext = file.name.split('.').pop().toUpperCase();
            filesHtml += `
                <div class="file-item" style="background: rgba(32, 191, 85, 0.1); padding: 10px 15px; margin: 5px 0; border-radius: 8px; display: flex; justify-content: space-between; align-items: center; transition: all 0.2s;">
                    <span style="font-weight: 600;">📄 ${file.name}</span>
                    <span style="font-size: 12px; opacity: 0.9; background: rgba(255,255,255,0.2); padding: 4px 10px; border-radius: 5px;">${ext} • ${sizeMB} MB</span>
                </div>
            `;
        });
        filesHtml += '</div>';
        
        displayEl.innerHTML = `
            <div style="font-size: 18px; font-weight: bold; margin-bottom: 12px; display: flex; align-items: center; justify-content: center; gap: 10px;">
                <span style="font-size: 24px;">✅</span>
                <span>تم اختيار ${selectedFiles.length} ملف</span>
            </div>
            <div style="font-size: 14px; margin-bottom: 10px; opacity: 0.9;">
                📦 الحجم الإجمالي: <strong>${totalSizeMB} MB</strong>
            </div>
            ${filesHtml}
        `;
        displayEl.style.background = 'rgba(32, 191, 85, 0.2)';
        displayEl.style.borderColor = '#20BF55';
        displayEl.style.color = '#20BF55';
        displayEl.style.borderStyle = 'solid';
        
        // Add hover effect to file items
        setTimeout(() => {
            const fileItems = displayEl.querySelectorAll('.file-item');
            fileItems.forEach(item => {
                item.addEventListener('mouseover', () => {
                    item.style.background = 'rgba(32, 191, 85, 0.2)';
                    item.style.transform = 'translateX(-5px)';
                });
                item.addEventListener('mouseout', () => {
                    item.style.background = 'rgba(32, 191, 85, 0.1)';
                    item.style.transform = 'translateX(0)';
                });
            });
        }, 100);
        
    } else {
        console.log('⚠️ No files selected');
        displayEl.innerHTML = '<span style="font-size: 20px;">⚠️</span> لم يتم اختيار ملفات';
        displayEl.style.background = 'rgba(243, 156, 18, 0.2)';
        displayEl.style.borderColor = '#F39C12';
        displayEl.style.color = '#F39C12';
        displayEl.style.borderStyle = 'dashed';
    }
}

// ==================== CAR PROCESSING ====================
function initializeCarForm() {
    const form = document.getElementById('carProcessForm');
    if (!form) {
        console.error('❌ Car process form not found');
        return;
    }
    
    form.addEventListener('submit', handleCarProcess);
    console.log('✅ Car form initialized');
}

async function handleCarProcess(event) {
    event.preventDefault();
    console.log('🚗 Starting car processing...');
    
    const processBtn = document.getElementById('carProcessBtn');
    const progressCard = document.getElementById('carProgressCard');
    const resultsCard = document.getElementById('resultsCard');
    const progressFill = document.getElementById('carProgressFill');
    const progressText = document.getElementById('carProgressText');
    
    // Validation
    if (selectedFiles.length === 0) {
        alert('⚠️ الرجاء اختيار ملف واحد على الأقل');
        return;
    }
    
    // Check file types
    const invalidFiles = selectedFiles.filter(f => {
        const ext = f.name.split('.').pop().toLowerCase();
        return ext !== 'xls' && ext !== 'xlsx';
    });
    
    if (invalidFiles.length > 0) {
        alert('⚠️ يجب أن تكون جميع الملفات من نوع Excel (.xls أو .xlsx)');
        return;
    }
    
    console.log(`📊 Processing ${selectedFiles.length} files in mode: ${currentMode}`);
    
    // UI updates
    processBtn.disabled = true;
    progressCard.style.display = 'block';
    resultsCard.style.display = 'none';
    progressFill.style.width = '10%';
    progressText.textContent = 'جاري رفع الملفات...';
    
    try {
        // Prepare FormData
        const formData = new FormData();
        
        selectedFiles.forEach(file => {
            formData.append('files', file);
            console.log(`📎 Added file to FormData: ${file.name}`);
        });
        
        formData.append('job_type', 'cars');
        formData.append('mode', currentMode);
        
        progressFill.style.width = '30%';
        progressText.textContent = 'جاري تحويل ومعالجة الملفات...';
        
        // Get token
        const token = getToken();
        if (!token) {
            throw new Error('Session expired');
        }
        
        console.log('🔐 Token acquired, sending request...');
        
        // Send request
        const response = await fetch(`${API_URL}/process`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`
            },
            body: formData
        });
        
        progressFill.style.width = '70%';
        progressText.textContent = 'جاري إنشاء التقرير النهائي...';
        
        if (!response.ok) {
            const errorData = await response.json();
            console.error('❌ Server error:', errorData);
            throw new Error(errorData.error || 'فشل في المعالجة');
        }
        
        console.log('✅ Server response received');
        
        // Get filename from headers
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = 'Xtractor_Report.xlsx';
        
        if (contentDisposition) {
            const match = contentDisposition.match(/filename\*?=utf-8''(.+)/i);
            if (match) {
                filename = decodeURIComponent(match[1]);
            }
        }
        
        console.log(`📥 Downloading file: ${filename}`);
        
        // Download file
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        console.log('✅ File downloaded successfully');
        
        // Update UI - success
        progressFill.style.width = '100%';
        progressFill.textContent = '✅ اكتملت المعالجة بنجاح!';
        progressFill.style.background = 'linear-gradient(90deg, #20BF55, #3CAEA3)';
        
        // Show results
        resultsCard.style.display = 'block';
        const resultsContent = document.getElementById('resultsContent');
        resultsContent.innerHTML = `
            <div style="text-align: center; padding: 20px;">
                <div style="font-size: 48px; margin-bottom: 20px;">🎉</div>
                <h3 style="color: var(--text-white); margin-bottom: 15px;">
                    تمت معالجة ${selectedFiles.length} ملف بنجاح!
                </h3>
                <div class="stat-item">
                    <span class="stat-label">📄 اسم الملف:</span>
                    <span class="stat-value">${filename}</span>
                </div>
                <p style="margin-top: 20px; color: var(--light-aqua); line-height: 1.8;">
                    ✅ تم تحويل جميع ملفات XLS إلى XLSX<br>
                    ✅ تم استخراج البيانات بنجاح<br>
                    ✅ تم فحص المواقع الجغرافية<br>
                    ✅ تم إنشاء تقرير موحد شامل<br>
                    📥 تم تحميل التقرير تلقائياً
                </p>
            </div>
        `;
        
        // Clear selected files
        selectedFiles = [];
        document.getElementById('filesInput').value = '';
        updateFileCountDisplay();
        
        showNotification('تمت المعالجة بنجاح! 🎉', 'success');
        
    } catch (error) {
        console.error('❌ Processing error:', error);
        
        progressFill.style.width = '100%';
        progressFill.textContent = '❌ فشلت العملية';
        progressFill.style.background = 'linear-gradient(90deg, #E74C3C, #C0392B)';
        progressText.textContent = `خطأ: ${error.message}`;
        
        alert(`❌ فشلت عملية المعالجة:\n\n${error.message}\n\nالرجاء المحاولة مرة أخرى`);
        
    } finally {
        processBtn.disabled = false;
        
        // Hide progress after delay
        setTimeout(() => {
            if (progressCard.style.display === 'block') {
                progressCard.style.display = 'none';
            }
        }, 5000);
    }
}

// ==================== VISITS PROCESSING ====================
function initializeVisitsForm() {
    const form = document.getElementById('visitsProcessForm');
    if (!form) {
        console.error('❌ Visits process form not found');
        return;
    }
    
    form.addEventListener('submit', handleVisitsProcess);
    console.log('✅ Visits form initialized');
}

async function handleVisitsProcess(event) {
    event.preventDefault();
    console.log('📊 Starting visits processing...');
    
    const processBtn = document.getElementById('visitsProcessBtn');
    const progressCard = document.getElementById('visitsProgressCard');
    const resultsCard = document.getElementById('visitsResultsCard');
    const progressFill = document.getElementById('visitsProgressFill');
    const progressText = document.getElementById('visitsProgressText');
    const sheetNameInput = document.getElementById('sheetNameInput');
    
    // Validation - for visits, we need exactly ONE file
    if (selectedFiles.length !== 1) {
        alert('⚠️ لعملية التوزيع، يجب اختيار ملف واحد فقط');
        return;
    }
    
    const file = selectedFiles[0];
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext !== 'xlsx' && ext !== 'xls') {
        alert('⚠️ يجب أن يكون الملف من نوع Excel (.xls أو .xlsx)');
        return;
    }
    
    const sheetName = sheetNameInput.value.trim() || 'إجمالي';
    console.log(`📋 Processing visits from sheet: ${sheetName}`);
    
    // UI updates
    processBtn.disabled = true;
    progressCard.style.display = 'block';
    resultsCard.style.display = 'none';
    progressFill.style.width = '10%';
    progressText.textContent = 'جاري رفع الملف...';
    
    try {
        // Prepare FormData
        const formData = new FormData();
        formData.append('files', file);
        formData.append('job_type', 'visits');
        formData.append('sheet_name', sheetName);
        
        console.log(`📎 Added file: ${file.name}`);
        
        progressFill.style.width = '30%';
        progressText.textContent = 'جاري معالجة البيانات وإنشاء الملفات...';
        
        // Get token
        const token = getToken();
        if (!token) {
            throw new Error('Session expired');
        }
        
        console.log('🔐 Token acquired, sending request...');
        
        // Send request
        const response = await fetch(`${API_URL}/process`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`
            },
            body: formData
        });
        
        progressFill.style.width = '70%';
        progressText.textContent = 'جاري ضغط الملفات...';
        
        if (!response.ok) {
            const errorData = await response.json();
            console.error('❌ Server error:', errorData);
            throw new Error(errorData.error || 'فشل في التوزيع');
        }
        
        console.log('✅ Server response received');
        
        // Get filename from headers
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = 'Xtractor_Visits.zip';
        
        if (contentDisposition) {
            const match = contentDisposition.match(/filename\*?=utf-8''(.+)/i);
            if (match) {
                filename = decodeURIComponent(match[1]);
            }
        }
        
        console.log(`📥 Downloading ZIP: ${filename}`);
        
        // Download ZIP file
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        console.log('✅ ZIP file downloaded successfully');
        
        // Update UI - success
        progressFill.style.width = '100%';
        progressFill.textContent = '✅ اكتمل التوزيع بنجاح!';
        progressFill.style.background = 'linear-gradient(90deg, #20BF55, #3CAEA3)';
        
        // Show results
        resultsCard.style.display = 'block';
        const resultsContent = document.getElementById('visitsResultsContent');
        resultsContent.innerHTML = `
            <div style="text-align: center; padding: 20px;">
                <div style="font-size: 48px; margin-bottom: 20px;">📦</div>
                <h3 style="color: var(--text-white); margin-bottom: 15px;">
                    تم توزيع التقرير بنجاح!
                </h3>
                <div class="stat-item">
                    <span class="stat-label">📦 ملف التوزيع:</span>
                    <span class="stat-value">${filename}</span>
                </div>
                <div class="stat-item">
                    <span class="stat-label">📋 الشيت المعالج:</span>
                    <span class="stat-value">${sheetName}</span>
                </div>
                <p style="margin-top: 20px; color: var(--light-aqua); line-height: 1.8;">
                    ✅ تم إنشاء ملفات المشرفين<br>
                    ✅ تم إنشاء ملفات المندوبين<br>
                    ✅ كل ملف يحتوي على شيتات منفصلة حسب الخطوط<br>
                    📥 تم تحميل ملف ZIP يحتوي على جميع الملفات
                </p>
            </div>
        `;
        
        // Clear selected files
        selectedFiles = [];
        document.getElementById('filesInput').value = '';
        updateFileCountDisplay();
        
        showNotification('تم التوزيع بنجاح! 📦', 'success');
        
    } catch (error) {
        console.error('❌ Distribution error:', error);
        
        progressFill.style.width = '100%';
        progressFill.textContent = '❌ فشلت العملية';
        progressFill.style.background = 'linear-gradient(90deg, #E74C3C, #C0392B)';
        progressText.textContent = `خطأ: ${error.message}`;
        
        alert(`❌ فشلت عملية التوزيع:\n\n${error.message}\n\nالرجاء المحاولة مرة أخرى`);
        
    } finally {
        processBtn.disabled = false;
        
        // Hide progress after delay
        setTimeout(() => {
            if (progressCard.style.display === 'block') {
                progressCard.style.display = 'none';
            }
        }, 5000);
    }
}

// ==================== USER INFO ====================
function displayUserInfo() {
    const user = getUser();
    if (!user) {
        console.warn('⚠️ No user data available');
        return;
    }
    
    console.log('👤 Current user:', user.username);
    console.log('📧 Email:', user.email);
}

// ==================== UTILITY FUNCTIONS ====================
function showNotification(message, type = 'info') {
    console.log(`📢 Notification: ${message} (${type})`);
    
    // Simple notification system
    const notification = document.createElement('div');
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 15px 25px;
        background: ${type === 'success' ? '#20BF55' : type === 'error' ? '#E74C3C' : '#3CAEA3'};
        color: white;
        border-radius: 10px;
        box-shadow: 0 5px 20px rgba(0,0,0,0.3);
        z-index: 10000;
        animation: slideIn 0.3s ease-out;
        font-weight: 600;
        font-size: 15px;
        max-width: 350px;
    `;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease-out';
        setTimeout(() => {
            if (document.body.contains(notification)) {
                document.body.removeChild(notification);
            }
        }, 300);
    }, 4000);
}

// Add CSS animations
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from {
            transform: translateX(400px);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
    
    @keyframes slideOut {
        from {
            transform: translateX(0);
            opacity: 1;
        }
        to {
            transform: translateX(400px);
            opacity: 0;
        }
    }
`;
document.head.appendChild(style);