
// ==========================================
// 替身 mock GAS_API
// ==========================================
const API_URL = "https://script.google.com/macros/s/AKfycbwDuDK2BYwykf0Z-u2FNFwxqyu0NZbE4emYceSMAIa3oD5JRUB9zIzRHfbVxtHdEzfnlg/exec"; // 請替換為部署後的網址

window.GAS_API = (function () {
  function createRunner(successHandler, failureHandler) {
    const handlerObj = {
      withSuccessHandler: function (callback) {
        return createRunner(callback, failureHandler);
      },
      withFailureHandler: function (callback) {
        return createRunner(successHandler, callback);
      }
    };

    return new Proxy(handlerObj, {
      get: function (target, prop) {
        if (prop in target) {
          return target[prop];
        }

        return function (...args) {
          console.log(`發送 API 請求：${prop.toString()}`, args);
          fetch(API_URL, {
            method: 'POST',
            body: JSON.stringify({
              action: prop,
              args: args
            })
          })
            .then(res => res.json())
            .then(result => {
              console.log(`收到 API 回應：${prop.toString()}`, result);
              if (result.status === 'success') {
                if (successHandler) successHandler(result.data);
              } else {
                if (failureHandler) failureHandler(new Error(result.message));
                else console.error("API Error:", result.message);
              }
            })
            .catch(error => {
              if (failureHandler) failureHandler(error);
              else console.error("Fetch Error:", error);
            });
        };
      }
    });
  }

  return createRunner(null, null);
})();

// ============================================
// 綜合施工處 每日工程日誌系統 JavaScript v2.1
// 修正日期：2025-01-18
// ============================================

// ============================================
// 全域變數
// ============================================
let allProjectsData = [];
let allInspectors = [];
let disasterOptions = [];
let currentUserInfo = null;
let currentHolidayInfo = null;
let currentSummaryData = [];
let filledDates = [];
let currentCalendarYear = new Date().getFullYear();
let currentCalendarMonth = new Date().getMonth();
let currentMonthHolidays = {};
let allInspectorsWithStatus = [];
let isGuestMode = true;
let guestViewMode = 'tomorrow'; // 新增：預設訪客檢視模式

// 檢驗員部門編號前綴映射
const DEPT_CODE_MAP = {
  '土木隊': 'CV',
  '建築隊': 'AR',
  '電氣隊': 'EL',
  '機械隊': 'ME',
  '中部隊': 'CT',
  '南部隊': 'ST',
  '委外監造': 'OS'
};

// ============================================
// 主題切換 (Dark/Light Mode)
// ============================================
function toggleTheme() {
  const body = document.body;
  const btn = document.getElementById('themeToggleBtn');
  
  if (body.classList.contains('dark')) {
    body.classList.remove('dark');
    body.classList.add('light');
    if (btn) btn.textContent = '🌙';
    localStorage.setItem('theme', 'light');
  } else {
    body.classList.remove('light');
    body.classList.add('dark');
    if (btn) btn.textContent = '☀️';
    localStorage.setItem('theme', 'dark');
  }
}

function initTheme() {
  const saved = localStorage.getItem('theme');
  const btn = document.getElementById('themeToggleBtn');
  
  if (saved === 'dark') {
    document.body.classList.add('dark');
    if (btn) btn.textContent = '☀️';
  } else if (saved === 'light') {
    document.body.classList.add('light');
    if (btn) btn.textContent = '🌙';
  } else {
    // Follow system preference
    if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
      document.body.classList.add('dark');
      if (btn) btn.textContent = '☀️';
    }
  }
}

// ============================================
// 初始化 - 訪客模式
// ============================================
document.addEventListener('DOMContentLoaded', function () {
  initTheme();
  initGuestMode();
});

function initGuestMode() {
  isGuestMode = true;
  currentUserInfo = null;
  showMainInterface();

  // 使用 setTimeout 確保 DOM 完全渲染後再載入資料
  setTimeout(() => {
    loadGuestData();
  }, 100);

  updateUIForGuestMode();
}

// 新增：切換訪客日期模式 (今日/明日)
function toggleGuestDate(mode) {
  guestViewMode = mode;

  // 更新按鈕樣式
  const btnToday = document.getElementById('guestBtnToday');
  const btnTomorrow = document.getElementById('guestBtnTomorrow');

  if (btnToday && btnTomorrow) {
    if (mode === 'today') {
      btnToday.classList.add('active');
      btnTomorrow.classList.remove('active');

      const label = document.getElementById('guestProjectLabel');
      if (label) label.textContent = '今日施工工程';
    } else {
      btnToday.classList.remove('active');
      btnTomorrow.classList.add('active');

      const label = document.getElementById('guestProjectLabel');
      if (label) label.textContent = '明日施工工程';
    }
  }

  // 重新載入資料
  loadGuestData();
}

// 修正：完整的 loadGuestData 函式
function loadGuestData() {
  // 先切換到總表頁籤
  showSummaryTab();

  // 使用 requestAnimationFrame 確保 DOM 完全更新後再設定值
  requestAnimationFrame(() => {
    // 根據模式設定目標日期
    const targetDate = new Date();
    if (guestViewMode === 'tomorrow') {
      targetDate.setDate(targetDate.getDate() + 1);
    }

    // 修正：使用本地時間格式化，避免時區問題
    const yyyy = targetDate.getFullYear();
    const mm = String(targetDate.getMonth() + 1).padStart(2, '0');
    const dd = String(targetDate.getDate()).padStart(2, '0');
    const dateString = `${yyyy}-${mm}-${dd}`;

    const datePickerElement = document.getElementById('summaryDatePicker');

    if (datePickerElement) {
      datePickerElement.value = dateString;
      console.log('訪客模式：設定總表日期為', dateString, '模式:', guestViewMode);
    }

    // 設定工程狀態為「施工中」
    const statusRadios = document.querySelectorAll('input[name="summaryStatusFilter"]');
    statusRadios.forEach(radio => {
      if (radio.value === 'active') {
        radio.checked = true;
      } else {
        radio.checked = false;
      }
    });

    // 設定部門為「全部部門」
    const deptFilter = document.getElementById('summaryDeptFilter');
    if (deptFilter) {
      deptFilter.value = 'all';
    }

    // 再次使用 requestAnimationFrame 確保所有設定完成後再載入資料
    requestAnimationFrame(() => {
      loadSummaryReport();
    });
  });
}

function checkLoginStatus() {
  showLoading();
  window.GAS_API
    .withSuccessHandler(function (session) {
      hideLoading();
      if (session.isLoggedIn) {
        currentUserInfo = session;
        isGuestMode = false;
        updateUIForLoggedIn();
        loadInitialData();
      } else {
        showToast('登入驗證失敗', true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('系統錯誤：' + error.message, true);
    })
    .getCurrentSession();
}

function showLoginInterface() {
  document.getElementById('loginModal').style.display = 'flex';
}

function hideLoginInterface() {
  document.getElementById('loginModal').style.display = 'none';
}

// ============================================
// 忘記密碼功能
// ============================================
function showForgotPasswordModal() {
  document.getElementById('forgotPasswordModal').style.display = 'flex';
  document.getElementById('forgotPasswordInput').value = '';
  document.getElementById('forgotPasswordInput').focus();
}

function closeForgotPasswordModal() {
  document.getElementById('forgotPasswordModal').style.display = 'none';
}

function submitForgotPassword() {
  const input = document.getElementById('forgotPasswordInput').value.trim();

  if (!input) {
    showToast('請輸入帳號或信箱', true);
    return;
  }

  showLoading();
  closeForgotPasswordModal();

  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();
      if (result.success) {
        showToast('✓ ' + result.message);
      } else {
        showToast(result.message, true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('發送失敗：' + error.message, true);
    })
    .sendTemporaryPassword(input);
}

function showMainInterface() {
  document.getElementById('loginContainer').style.display = 'none';
  document.getElementById('mainContainer').style.display = 'block';
  setupEventListeners();
}

function updateUIForGuestMode() {
  document.getElementById('userInfoPanel').style.display = 'none';
  document.getElementById('guestLoginBtn').style.display = 'flex';

  // 隱藏導航列
  const tabs = document.querySelector('.tabs');
  if (tabs) tabs.style.display = 'none';

  // 隱藏所有頁籤，只顯示總表
  document.querySelectorAll('.tab-pane').forEach(pane => {
    pane.style.display = 'none';
  });
  const summaryReport = document.getElementById('summaryReport');
  if (summaryReport) summaryReport.style.display = 'block';

  // 顯示訪客模式卡片
  const guestCards = document.getElementById('guestSummaryCards');
  if (guestCards) guestCards.style.display = 'block';

  // 隱藏控制區（篩選器）
  const summaryControls = document.querySelector('.summary-controls');
  if (summaryControls) summaryControls.style.display = 'none';

  // 隱藏 TBM-KY 文件生成
  const tbmkyCard = document.getElementById('tbmkyCard');
  if (tbmkyCard) tbmkyCard.style.display = 'none';

  // 移除日曆模式按鈕（訪客模式專用）
  const calendarModeBtn = document.getElementById('calendarModeBtn');
  if (calendarModeBtn && calendarModeBtn.parentNode) {
    calendarModeBtn.parentNode.removeChild(calendarModeBtn);
  }
}

function updateUIForLoggedIn() {
  if (currentUserInfo) {
    document.getElementById('currentUserName').textContent = currentUserInfo.name;
    document.getElementById('currentUserRole').textContent = currentUserInfo.role;

    // 顯示導航列
    const tabs = document.querySelector('.tabs');
    if (tabs) tabs.style.display = 'flex';

    // 顯示控制區（篩選器）
    const summaryControls = document.querySelector('.summary-controls');
    if (summaryControls) summaryControls.style.display = 'block';

    // 隱藏訪客模式卡片
    const guestCards = document.getElementById('guestSummaryCards');
    if (guestCards) guestCards.style.display = 'none';

    // 顯示 TBM-KY 文件生成
    const tbmkyCard = document.getElementById('tbmkyCard');
    if (tbmkyCard) tbmkyCard.style.display = 'block';

    // 重新添加日曆模式按鈕（如果不存在則重新建立）
    const modeToggle = document.getElementById('modeToggle');
    let calendarModeBtn = document.getElementById('calendarModeBtn');

    if (!calendarModeBtn && modeToggle) {
      // 如果按鈕被移除了，重新建立
      calendarModeBtn = document.createElement('button');
      calendarModeBtn.id = 'calendarModeBtn';
      calendarModeBtn.className = 'mode-btn';
      calendarModeBtn.setAttribute('data-mode', 'calendar');
      calendarModeBtn.innerHTML = '<span>📅 日曆模式</span>';
      modeToggle.appendChild(calendarModeBtn);
    } else if (calendarModeBtn) {
      calendarModeBtn.style.display = 'block';
    }

    // 移除所有頁籤的 inline style，讓 CSS 類別控制顯示
    document.querySelectorAll('.tab-pane').forEach(pane => {
      pane.style.display = '';
    });

    // 填表人：隱藏特定導航列
    if (currentUserInfo.role === '填表人') {
      document.querySelector('.tab-logStatus').style.display = 'none';
      document.querySelector('.tab-inspectorManagement').style.display = 'none';
      document.querySelector('.tab-userManagement').style.display = 'none';
    }
    document.getElementById('userInfoPanel').style.display = 'flex';
    document.getElementById('helpBtn').style.display = 'flex';
    document.getElementById('changePasswordBtn').style.display = 'flex';
    document.getElementById('logoutBtn').style.display = 'flex';
    document.getElementById('guestLoginBtn').style.display = 'none';

    // 根據角色顯示/隱藏使用者管理Tab
    const userManagementTab = document.querySelector('.tab[data-tab="userManagement"]');
    if (userManagementTab) {
      if (currentUserInfo.role === '填表人') {
        // 填表人：隱藏使用者管理
        userManagementTab.style.display = 'none';
      } else {
        // 超級管理員、聯絡員：顯示使用者管理
        userManagementTab.style.display = 'flex';

        // 更新按鈕文字
        const userManagementTitle = document.getElementById('userManagementTitle');
        if (userManagementTitle) {
          if (currentUserInfo.role === '聯絡員') {
            userManagementTitle.textContent = '填表人管理';
          } else {
            userManagementTitle.textContent = '使用者管理';
          }
        }
      }
    }
  }
}

// ============================================
// 修正2：密碼顯示/隱藏切換
// ============================================
function togglePasswordVisibility() {
  const passwordInput = document.getElementById('loginPassword');
  const toggleIcon = document.getElementById('passwordToggleIcon');

  if (passwordInput.type === 'password') {
    passwordInput.type = 'text';
    toggleIcon.textContent = '🙈';
  } else {
    passwordInput.type = 'password';
    toggleIcon.textContent = '👁️';
  }
}

// ============================================
// 登入/登出功能
// ============================================
function decodeJwtResponse(token) {
  var base64Url = token.split('.')[1];
  var base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
  var jsonPayload = decodeURIComponent(window.atob(base64).split('').map(function (c) {
    return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
  }).join(''));
  return JSON.parse(jsonPayload);
}

function handleCredentialResponse(response) {
  const responsePayload = decodeJwtResponse(response.credential);
  const email = responsePayload.email;

  showLoading();
  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();
      if (result.success) {
        currentUserInfo = result.user;
        isGuestMode = false;
        hideLoginInterface();
        updateUIForLoggedIn();
        showToast(result.message);
        setTimeout(() => {
          loadInitialData();
          if (currentUserInfo.role === '填表人') {
            checkFillerReminders();
          }
        }, 500);
      } else {
        showToast(result.message, true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('Google 登入失敗：' + error.message, true);
    })
    .authenticateGoogleUser(email);
}

function handleLogin(event) {
  event.preventDefault();

  const identifier = document.getElementById('loginIdentifier').value.trim();
  const password = document.getElementById('loginPassword').value;

  if (!identifier || !password) {
    showToast('請輸入帳號/信箱和密碼', true);
    return;
  }

  showLoading();
  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();
      if (result.success) {
        currentUserInfo = result.user;
        isGuestMode = false;
        hideLoginInterface();
        updateUIForLoggedIn();
        showToast(result.message);
        setTimeout(() => {
          loadInitialData();
          // 填表人登入後檢查未填寫工程
          if (currentUserInfo.role === '填表人') {
            checkFillerReminders();
          }
        }, 500);
      } else {
        showToast(result.message, true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('登入失敗：' + error.message, true);
    })
    .authenticateUser(identifier, password);
}

function handleLogout() {
  const confirmMessage = `
    <p><strong>🚪 確認登出</strong></p>
    <p>您確定要登出系統嗎？</p>
    <p>登出後將進入訪客模式</p>
  `;

  showConfirmModal(confirmMessage, function () {
    showLoading();
    window.GAS_API
      .withSuccessHandler(function (result) {
        hideLoading();
        if (result.success) {
          showToast(result.message);
          setTimeout(() => {
            location.reload();
          }, 1000);
        }
      })
      .withFailureHandler(function (error) {
        hideLoading();
        showToast('登出失敗：' + error.message, true);
      })
      .logoutUser();
    closeConfirmModal();
  });
}

// ============================================
// 變更密碼功能
// ============================================
function showRoleGuideModal() {
  const modal = document.getElementById('roleGuideModal');
  const fillerGuide = document.getElementById('fillerGuide');
  const contactGuide = document.getElementById('contactGuide');

  // 根據使用者角色顯示對應的說明
  if (currentUserInfo && currentUserInfo.role === '聯絡員') {
    fillerGuide.style.display = 'none';
    contactGuide.style.display = 'block';
  } else {
    // 預設顯示填表人說明
    fillerGuide.style.display = 'block';
    contactGuide.style.display = 'none';
  }

  modal.style.display = 'flex';
}

function closeRoleGuideModal() {
  document.getElementById('roleGuideModal').style.display = 'none';
}

function showChangePasswordModal() {
  document.getElementById('oldPassword').value = '';
  document.getElementById('newPassword').value = '';
  document.getElementById('confirmPassword').value = '';
  document.getElementById('changePasswordModal').style.display = 'flex';
}

function closeChangePasswordModal() {
  document.getElementById('changePasswordModal').style.display = 'none';
}

function submitChangePassword() {
  const oldPassword = document.getElementById('oldPassword').value.trim();
  const newPassword = document.getElementById('newPassword').value.trim();
  const confirmPassword = document.getElementById('confirmPassword').value.trim();

  if (!oldPassword || !newPassword || !confirmPassword) {
    showToast('請填寫所有欄位', true);
    return;
  }

  if (newPassword.length < 6) {
    showToast('新密碼長度至少 6 位', true);
    return;
  }

  if (newPassword !== confirmPassword) {
    showToast('新密碼與確認密碼不一致', true);
    return;
  }

  if (!currentUserInfo) {
    showToast('請先登入', true);
    return;
  }

  showLoading();
  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();
      if (result.success) {
        showToast(result.message);
        closeChangePasswordModal();
      } else {
        showToast(result.message, true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('變更密碼失敗：' + error.message, true);
    })
    .changeUserPassword(currentUserInfo.account, oldPassword, newPassword);
}

// ============================================
// 填表人未填寫提醒功能
// ============================================
function checkFillerReminders() {
  if (!currentUserInfo || currentUserInfo.role !== '填表人') return;

  // Debug：輸出 managedProjects 資訊
  console.log('[checkFillerReminders] currentUserInfo:', currentUserInfo);
  console.log('[checkFillerReminders] managedProjects type:', typeof currentUserInfo.managedProjects);
  console.log('[checkFillerReminders] managedProjects value:', currentUserInfo.managedProjects);

  // 確保 managedProjects 是字串格式
  const managedProjectsStr = Array.isArray(currentUserInfo.managedProjects)
    ? currentUserInfo.managedProjects.join(',')
    : String(currentUserInfo.managedProjects || '');

  console.log('[checkFillerReminders] 傳遞給後端的字串:', managedProjectsStr);

  window.GAS_API
    .withSuccessHandler(function (result) {
      console.log('[checkFillerReminders] 後端返回結果:', result);
      if (result.unfilledProjects.length > 0 || result.incompleteProjects.length > 0) {
        showFillerReminderModal(result);
      }
    })
    .withFailureHandler(function (error) {
      console.error('檢查未填寫提醒失敗：' + error.message);
    })
    .getFillerReminders(managedProjectsStr);
}

function showFillerReminderModal(data) {
  const content = document.getElementById('fillerReminderContent');
  let html = '';

  if (data.unfilledProjects.length > 0) {
    html += `
      <div style="margin-bottom: 1.5rem;">
        <h4 style="color: var(--danger); margin-bottom: 1rem;">⚠️ 明日（${data.tomorrowDate}）尚未填寫的工程（${data.unfilledProjects.length}）</h4>
        <div style="display: flex; flex-direction: column; gap: 0.75rem;">
          ${data.unfilledProjects.map(proj => `
            <div style="padding: 1rem; background: #fef2f2; border-left: 4px solid var(--danger); border-radius: 8px; display: flex; justify-content: space-between; align-items: center; gap: 1rem;">
              <div style="flex: 1;">
                <div style="font-weight: 600; color: #374151; margin-bottom: 0.25rem;">${proj.fullName}</div>
                <div style="font-size: 0.875rem; color: #6b7280;">序號：${proj.seqNo} | 承攬商：${proj.contractor}</div>
              </div>
              <button class="btn btn-primary" onclick="goToLogEntry('${proj.seqNo}')" style="white-space: nowrap;">
                <span class="btn-icon">📝</span>
                <span>前往填寫</span>
              </button>
            </div>
          `).join('')}
        </div>
      </div>
    `;
  }

  if (data.incompleteProjects.length > 0) {
    html += `
      <div>
        <h4 style="color: var(--warning); margin-bottom: 1rem;">📝 工程設定未完整（${data.incompleteProjects.length}）</h4>
        <div style="display: flex; flex-direction: column; gap: 0.75rem;">
          ${data.incompleteProjects.map(proj => `
            <div style="padding: 1rem; background: #fffbeb; border-left: 4px solid var(--warning); border-radius: 8px; display: flex; justify-content: space-between; align-items: center; gap: 1rem;">
              <div style="flex: 1;">
                <div style="font-weight: 600; color: #374151; margin-bottom: 0.5rem;">${proj.fullName}</div>
                <div style="font-size: 0.875rem; color: #6b7280; margin-bottom: 0.25rem;">序號：${proj.seqNo}</div>
                <div style="font-size: 0.875rem; color: var(--danger);">
                  ${proj.missingFields.join('、')} 未填寫
                </div>
              </div>
              <button class="btn btn-secondary" onclick="goToProjectSetup('${proj.seqNo}')" style="white-space: nowrap;">
                <span class="btn-icon">⚙️</span>
                <span>前往設定</span>
              </button>
            </div>
          `).join('')}
        </div>
      </div>
    `;
  }

  content.innerHTML = html;
  document.getElementById('fillerReminderModal').style.display = 'flex';
}

function closeFillerReminderModal() {
  document.getElementById('fillerReminderModal').style.display = 'none';
}

function goToLogEntry(projectSeqNo) {
  // 關閉提醒視窗
  closeFillerReminderModal();

  // 切換到日誌填報頁籤
  switchTab('logEntry');

  // 自動選擇該工程
  setTimeout(() => {
    const projectSelect = document.getElementById('logProjectSelect');
    if (projectSelect) {
      projectSelect.value = projectSeqNo;

      // 觸發 change 事件以載入工程資料
      const event = new Event('change');
      projectSelect.dispatchEvent(event);

      // 滾動到表單頂部
      document.getElementById('logEntry').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, 300);
}

function goToProjectSetup(projectSeqNo) {
  // 關閉提醒視窗
  closeFillerReminderModal();

  // 切換到工程設定頁籤
  switchTab('projectSetup');

  // 找到該工程卡片並滾動到視窗中
  setTimeout(() => {
    const projectCards = document.querySelectorAll('.project-card');
    projectCards.forEach(card => {
      const seqNoElement = card.querySelector('[data-seq-no]');
      if (seqNoElement && seqNoElement.getAttribute('data-seq-no') === projectSeqNo) {
        card.scrollIntoView({ behavior: 'smooth', block: 'center' });

        // 高亮該卡片（添加短暫的動畫效果）
        card.style.boxShadow = '0 0 0 3px var(--primary)';
        setTimeout(() => {
          card.style.boxShadow = '';
        }, 2000);
      }
    });
  }, 300);
}

function updateUnfilledCardsDisplay() {
  if (!currentUserInfo || currentUserInfo.role !== '填表人') {
    document.getElementById('unfilledCardsContainer').style.display = 'none';
    return;
  }

  window.GAS_API
    .withSuccessHandler(function (result) {
      const container = document.getElementById('unfilledCardsContainer');

      if (result.unfilledProjects.length === 0 && result.incompleteProjects.length === 0) {
        container.style.display = 'none';
        return;
      }

      let html = '';

      if (result.unfilledProjects.length > 0) {
        html += `
          <div class="alert-warning" style="margin-bottom: 1rem;">
            <strong>⚠️ 明日（${result.tomorrowDate}）尚未填寫的工程（${result.unfilledProjects.length}）</strong>
            <div style="margin-top: 0.75rem; display: flex; flex-direction: column; gap: 0.5rem;">
              ${result.unfilledProjects.map(proj => `
                <div style="padding: 0.75rem; background: white; border-radius: 6px; display: flex; justify-content: space-between; align-items: center; gap: 1rem;">
                  <div style="flex: 1;">
                    <strong>${proj.fullName}</strong> - 序號：${proj.seqNo}
                  </div>
                  <button class="btn btn-primary btn-sm" onclick="goToLogEntry('${proj.seqNo}')" style="white-space: nowrap; padding: 0.5rem 1rem; font-size: 0.875rem;">
                    📝 前往填寫
                  </button>
                </div>
              `).join('')}
            </div>
          </div>
        `;
      }

      if (result.incompleteProjects.length > 0) {
        html += `
          <div class="alert-info">
            <strong>📝 工程設定未完整（${result.incompleteProjects.length}）</strong>
            <div style="margin-top: 0.75rem; display: flex; flex-direction: column; gap: 0.5rem;">
              ${result.incompleteProjects.map(proj => `
                <div style="padding: 0.75rem; background: white; border-radius: 6px; display: flex; justify-content: space-between; align-items: center; gap: 1rem;">
                  <div style="flex: 1;">
                    <div><strong>${proj.fullName}</strong> - 序號：${proj.seqNo}</div>
                    <div style="color: var(--danger); font-size: 0.875rem; margin-top: 0.25rem;">
                      ${proj.missingFields.join('、')} 未填寫
                    </div>
                  </div>
                  <button class="btn btn-secondary btn-sm" onclick="goToProjectSetup('${proj.seqNo}')" style="white-space: nowrap; padding: 0.5rem 1rem; font-size: 0.875rem;">
                    ⚙️ 前往設定
                  </button>
                </div>
              `).join('')}
            </div>
          </div>
        `;
      }

      container.innerHTML = html;
      container.style.display = 'block';
    })
    .withFailureHandler(function (error) {
      console.error('更新提醒區塊失敗：' + error.message);
    })
    .getFillerReminders(Array.isArray(currentUserInfo.managedProjects)
      ? currentUserInfo.managedProjects.join(',')
      : String(currentUserInfo.managedProjects || ''));
}

// ============================================
// 事件監聽器設置
// ============================================
function setupEventListeners() {
  // 登入表單
  const loginForm = document.getElementById('loginForm');
  if (loginForm) {
    loginForm.addEventListener('submit', handleLogin);
  }

  // 登出按鈕
  const logoutBtn = document.getElementById('logoutBtn');
  if (logoutBtn) {
    logoutBtn.addEventListener('click', handleLogout);
  }

  // 頁籤切換
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', function () {
      const targetTab = this.getAttribute('data-tab');
      switchTab(targetTab);
    });
  });

  // 日誌填報表單
  document.getElementById('dailyLogForm').addEventListener('submit', handleDailyLogSubmit);

  // 日期選擇器 - 設置為明天
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);

  // 修正：使用本地時間格式化
  const yyyy = tomorrow.getFullYear();
  const mm = String(tomorrow.getMonth() + 1).padStart(2, '0');
  const dd = String(tomorrow.getDate()).padStart(2, '0');
  document.getElementById('logDatePicker').value = `${yyyy}-${mm}-${dd}`;


  // 工程選擇變更
  document.getElementById('logProjectSelect').addEventListener('change', handleProjectChange);

  // 假日選項互斥
  document.getElementById('isHolidayWork').addEventListener('change', function () {
    if (this.checked) {
      document.getElementById('isHolidayNoWork').checked = false;
      document.getElementById('holidayNoWorkDetails').style.display = 'none'; // [新增] 隱藏細項
      toggleWorkFields(false); // 假日施工：顯示所有欄位
    }
  });

  document.getElementById('isHolidayNoWork').addEventListener('change', function () {
    const detailsDiv = document.getElementById('holidayNoWorkDetails');
    if (this.checked) {
      document.getElementById('isHolidayWork').checked = false;
      detailsDiv.style.display = 'flex'; // [新增] 展開細項
      toggleWorkFields(true); // 假日不施工：隱藏欄位
    } else {
      detailsDiv.style.display = 'none'; // [新增] 隱藏細項
      // 取消假日不施工：顯示欄位
      toggleWorkFields(false);
    }
  });

  // 修正7：新增工項按鈕
  document.getElementById('addWorkItemBtn').addEventListener('click', addWorkItemPair);

  // 修改檢驗員按鈕
  document.getElementById('changeInspectorBtn').addEventListener('click', toggleInspectorEditMode);

  // 總表功能
  document.getElementById('refreshSummary').addEventListener('click', loadSummaryReport);
  document.getElementById('summaryDatePicker').addEventListener('change', loadSummaryReport);
  document.querySelectorAll('input[name="summaryStatusFilter"]').forEach(radio => {
    radio.addEventListener('change', loadSummaryReport);
  });
  document.getElementById('summaryDeptFilter').addEventListener('change', loadSummaryReport);
  document.getElementById('summaryInspectorFilter').addEventListener('change', loadSummaryReport); // [新增]

  // 總表模式切換
  document.querySelectorAll('.mode-btn').forEach(btn => {
    btn.addEventListener('click', function () {
      const mode = this.getAttribute('data-mode');
      switchSummaryMode(mode);
    });
  });

  // TBM-KY 生成
  document.getElementById('generateTBMKYBtn').addEventListener('click', openTBMKYModal);

  // 日誌狀況
  document.getElementById('refreshLogStatus').addEventListener('click', loadLogStatus);

  // 工程設定
  document.getElementById('refreshProjectList').addEventListener('click', loadAndRenderProjectCards);
  document.querySelectorAll('input[name="projectStatusFilter"]').forEach(radio => {
    radio.addEventListener('change', loadAndRenderProjectCards);
  });
  document.getElementById('projectDeptFilter').addEventListener('change', loadAndRenderProjectCards);

  // 工程狀態變更監聽
  const editProjectStatus = document.getElementById('editProjectStatus');
  if (editProjectStatus) {
    editProjectStatus.addEventListener('change', function () {
      const remarkGroup = document.getElementById('editStatusRemarkGroup');
      const remarkTextarea = document.getElementById('editStatusRemark');

      if (this.value !== '施工中') {
        remarkGroup.style.display = 'block';
        remarkTextarea.setAttribute('required', 'required');
      } else {
        remarkGroup.style.display = 'none';
        remarkTextarea.removeAttribute('required');
        remarkTextarea.value = '';
      }
    });
  }

  // 檢驗員管理
  document.getElementById('addInspectorBtn').addEventListener('click', openAddInspectorModal);
  document.getElementById('refreshInspectorList').addEventListener('click', loadInspectorManagement);
  document.querySelectorAll('input[name="inspectorStatusFilter"]').forEach(radio => {
    radio.addEventListener('change', loadInspectorManagement);
  });
  document.getElementById('inspectorDeptFilter').addEventListener('change', loadInspectorManagement);

  // 日期變更時檢查假日
  document.getElementById('logDatePicker').addEventListener('change', checkAndShowHolidayAlert);
}

// ============================================
// 載入初始資料
// ============================================
function loadInitialData() {
  showLoading();

  window.GAS_API
    .withSuccessHandler(function (data) {
      allProjectsData = data.projects || [];
      disasterOptions = data.disasters || [];
      allInspectors = data.inspectors || [];

      renderProjectSelect('logProjectSelect', allProjectsData, true);

      const depts = extractDepartments(allProjectsData);
      renderDepartmentFilters(depts);

      // [新增] 載入檢驗員篩選器
      loadInspectorsForFilter();

      checkAndShowHolidayAlert();
      loadUnfilledCount();
      setupSummaryDate();
      loadFilledDates();

      hideLoading();

      // 填表人：直接彈出明日未填報工程
      if (currentUserInfo && currentUserInfo.role === '填表人') {
        openFillerStartupModal();
      } else {
        // 其他角色：預設進入總表頁籤
        showSummaryTab();

        // 設定工程狀態為「施工中」
        const statusRadios = document.querySelectorAll('input[name="summaryStatusFilter"]');
        statusRadios.forEach(radio => {
          if (radio.value === 'active') {
            radio.checked = true;
          } else {
            radio.checked = false;
          }
        });

        // 設定部門為「全部部門」
        document.getElementById('summaryDeptFilter').value = 'all';

        // 載入總表資料
        loadSummaryReport();
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('載入初始資料失敗：' + error.message, true);
    })
    .loadLogSetupData();
}

function extractDepartments(projects) {
  // 從 DEPT_CODE_MAP 取得所有可用的部門（確保所有部門都可選）
  const deptSet = new Set(Object.keys(DEPT_CODE_MAP));

  // 同時加入工程資料中的部門（向下相容）
  projects.forEach(p => {
    if (p.dept) {
      deptSet.add(p.dept);
    }
  });

  const deptArray = Array.from(deptSet);

  // 修正5：排序 - 「隊」優先，「委外監造」最後
  deptArray.sort((a, b) => {
    const aIsTeam = a.includes('隊');
    const bIsTeam = b.includes('隊');

    if (aIsTeam && !bIsTeam) return -1;
    if (!aIsTeam && bIsTeam) return 1;

    if (a === '委外監造' && b !== '委外監造') return 1;
    if (a !== '委外監造' && b === '委外監造') return -1;

    return a.localeCompare(b, 'zh-TW');
  });

  return deptArray;
}

function renderDepartmentFilters(depts) {
  const summaryDeptFilter = document.getElementById('summaryDeptFilter');
  const projectDeptFilter = document.getElementById('projectDeptFilter');
  const inspectorDeptFilter = document.getElementById('inspectorDeptFilter');
  const addInspectorDept = document.getElementById('addInspectorDept');
  const editInspectorDept = document.getElementById('editInspectorDept');

  const filters = [summaryDeptFilter, projectDeptFilter, inspectorDeptFilter];
  const selects = [addInspectorDept, editInspectorDept];

  filters.forEach(filter => {
    if (filter) {
      const currentValue = filter.value;
      filter.innerHTML = '<option value="all">全部部門</option>';
      depts.forEach(dept => {
        const option = document.createElement('option');
        option.value = dept;
        option.textContent = dept;
        filter.appendChild(option);
      });
      filter.value = currentValue;
    }
  });

  selects.forEach(select => {
    if (select) {
      const currentValue = select.value;
      select.innerHTML = '<option value="">請選擇部門</option>';
      depts.forEach(dept => {
        const option = document.createElement('option');
        option.value = dept;
        option.textContent = dept;
        select.appendChild(option);
      });
      select.value = currentValue;
    }
  });
}

// ============================================
// 頁籤切換
// ============================================
function switchTab(tabName) {
  // 訪客模式權限檢查
  if (isGuestMode && tabName !== 'summaryReport') {
    showToast('請先登入才能使用此功能', true);
    showLoginInterface();
    return;
  }

  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));

  document.querySelector(`.tab[data-tab="${tabName}"]`).classList.add('active');
  document.getElementById(tabName).classList.add('active');

  // 顯示讀取中（除了日誌填報頁籤）
  if (tabName !== 'logEntry') {
    showLoading();
  }

  // 載入對應頁籤的資料
  switch (tabName) {
    case 'summaryReport':
      loadSummaryReport();
      break;
    case 'logEntry':
      // 切換到日誌填報時，更新提醒區塊
      if (currentUserInfo && currentUserInfo.role === '填表人') {
        updateUnfilledCardsDisplay();
      }
      break;
    case 'logStatus':
      loadLogStatus();
      break;
    case 'projectSetup':
      loadAndRenderProjectCards();
      break;
    case 'inspectorManagement':
      loadInspectorManagement();
      break;
    case 'userManagement':
      loadUserManagement();
      break;
  }
}

function showSummaryTab() {
  switchTab('summaryReport');
}

// ============================================
// 假日檢查與提示
// ============================================
function checkAndShowHolidayAlert() {
  const dateString = document.getElementById('logDatePicker').value;
  if (!dateString) return;

  // 重置勾選框
  document.getElementById('checkSat').checked = false;
  document.getElementById('checkSun').checked = false;
  document.getElementById('checkHoliday').checked = false;

  // [修正] 啟用假日勾選框
  document.getElementById('checkSat').disabled = false;
  document.getElementById('checkSun').disabled = false;
  document.getElementById('checkHoliday').disabled = false;

  window.GAS_API
    .withSuccessHandler(function (holidayInfo) {
      currentHolidayInfo = holidayInfo;

      const holidayAlert = document.getElementById('holidayAlert');
      const holidayDateInfo = document.getElementById('holidayDateInfo');
      const holidayRemark = document.getElementById('holidayRemark');
      const holidayWorkGroup = document.getElementById('holidayWorkGroup');
      const isHolidayNoWorkCheckbox = document.getElementById('isHolidayNoWork');
      const isHolidayWorkCheckbox = document.getElementById('isHolidayWork');

      // 判斷星期並自動勾選 UI
      const dateObj = new Date(dateString);
      const day = dateObj.getDay(); // 0 is Sunday, 6 is Saturday

      if (day === 6) document.getElementById('checkSat').checked = true;
      if (day === 0) document.getElementById('checkSun').checked = true;

      // 判斷是否為例假
      if (holidayInfo.isHoliday && day !== 0 && day !== 6) {
        document.getElementById('checkHoliday').checked = true;
      }

      if (holidayInfo.isHoliday) {
        holidayAlert.style.display = 'block';
        holidayDateInfo.textContent = `${dateString} (星期${holidayInfo.weekday})`;
        holidayRemark.textContent = holidayInfo.remark || '假日';
        holidayWorkGroup.style.display = 'block';

        // [修正] 預設勾選假日不施工，但允許修改
        isHolidayNoWorkCheckbox.checked = true;
        isHolidayWorkCheckbox.checked = false;

        // 隱藏其他欄位
        toggleWorkFields(true);

        // 自動提交假日不施工記錄
        autoSubmitHolidayNoWork(dateString);
      } else {
        holidayAlert.style.display = 'none';
        holidayWorkGroup.style.display = 'none';

        // 非假日：取消勾選並顯示欄位
        isHolidayNoWorkCheckbox.checked = false;
        isHolidayWorkCheckbox.checked = false;
        toggleWorkFields(false);
      }
    })
    .withFailureHandler(function (error) {
      console.error('檢查假日失敗：', error);
    })
    .checkHoliday(dateString);
}

// ============================================
// 假日自動提交不施工記錄
// ============================================
function autoSubmitHolidayNoWork(dateString) {
  // 取得工程下拉選單
  const projectSelect = document.getElementById('logProjectSelect');
  const projectOptions = projectSelect.options;

  if (projectOptions.length <= 1) {
    // 沒有可選工程
    return;
  }

  // 取得所有可選工程（排除第一個 placeholder）
  const projects = [];
  for (let i = 1; i < projectOptions.length; i++) {
    const option = projectOptions[i];
    projects.push({
      seqNo: option.value,
      shortName: option.getAttribute('data-short-name') || option.text
    });
  }

  // 檢查哪些工程尚未填報
  showLoading();
  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();

      const unfilledProjects = result.unfilledProjects || [];
      const alreadyFilledCount = result.alreadyFilledCount || 0;

      if (unfilledProjects.length === 0) {
        // 所有工程都已填報
        showToast(`✓ ${dateString} 假日不施工記錄已全部完成（${alreadyFilledCount} 項工程）`);
        return;
      }

      // 批次提交未填報工程的假日不施工記錄
      batchSubmitHolidayNoWork(dateString, unfilledProjects);
    })
    .withFailureHandler(function (error) {
      hideLoading();
      console.error('檢查假日填報狀態失敗：', error);
    })
    .checkHolidayFilledStatus(dateString, projects);
}

function batchSubmitHolidayNoWork(dateString, projects) {
  if (projects.length === 0) {
    return;
  }

  let successCount = 0;
  let totalCount = projects.length;
  let errors = [];

  showLoading();

  // 逐一提交
  projects.forEach((project, index) => {
    window.GAS_API
      .withSuccessHandler(function (result) {
        if (result.success) {
          successCount++;
        } else {
          errors.push(`${project.shortName}: ${result.message}`);
        }

        // 最後一個完成時顯示結果
        if (index === totalCount - 1) {
          hideLoading();
          if (successCount === totalCount) {
            showToast(`✓ 已自動完成 ${successCount} 項工程的假日不施工記錄`);
          } else {
            showToast(`完成 ${successCount}/${totalCount} 項工程，${errors.length} 項失敗`, errors.length > 0);
          }
          // 重新載入未填報數量
          loadUnfilledCount();
        }
      })
      .withFailureHandler(function (error) {
        errors.push(`${project.shortName}: ${error.message}`);

        if (index === totalCount - 1) {
          hideLoading();
          showToast(`完成 ${successCount}/${totalCount} 項工程，${errors.length} 項失敗`, true);
          loadUnfilledCount();
        }
      })
      .submitDailyLog({
        logDate: dateString,
        projectSeqNo: project.seqNo,
        projectShortName: project.shortName,
        isHolidayNoWork: true,
        isHolidayWork: false,
        inspectorIds: [],
        workersCount: 0,
        workItems: []
      });
  });
}

// ============================================
// 填表人啟動：顯示明日未填報工程
// ============================================
function openFillerStartupModal() {
  // 切換到日誌填報頁籤
  switchTab('logEntry');

  // 確認使用者資訊
  if (!currentUserInfo || !currentUserInfo.managedProjects) {
    console.error('openFillerStartupModal: 無使用者資訊或管理工程');
    showToast('無法取得使用者管理工程資訊', true);
    return;
  }

  console.log('[openFillerStartupModal] currentUserInfo =', currentUserInfo);
  console.log('[openFillerStartupModal] managedProjects type =', typeof currentUserInfo.managedProjects);
  console.log('[openFillerStartupModal] managedProjects value =', currentUserInfo.managedProjects);

  // 確保 managedProjects 是字串格式
  const managedProjectsStr = Array.isArray(currentUserInfo.managedProjects)
    ? currentUserInfo.managedProjects.join(',')
    : String(currentUserInfo.managedProjects || '');

  console.log('[openFillerStartupModal] 傳遞給後端的字串:', managedProjectsStr);

  // 顯示 loading
  showLoading();

  // 使用新的 getFillerReminders API
  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();

      console.log('[openFillerStartupModal] getFillerReminders 返回結果:', result);

      // 檢查是否有任何未完成項目
      const hasUnfilled = result.unfilledProjects && result.unfilledProjects.length > 0;
      const hasIncomplete = result.incompleteProjects && result.incompleteProjects.length > 0;

      console.log('[openFillerStartupModal] hasUnfilled:', hasUnfilled, '未填寫數量:', result.unfilledProjects ? result.unfilledProjects.length : 0);
      console.log('[openFillerStartupModal] hasIncomplete:', hasIncomplete, '未完整數量:', result.incompleteProjects ? result.incompleteProjects.length : 0);

      if (!hasUnfilled && !hasIncomplete) {
        showToast('✓ 明日所有工程都已填報完成！');
        return;
      }

      // 顯示提醒彈窗
      showFillerReminderModal(result);
    })
    .withFailureHandler(function (error) {
      hideLoading();
      console.error('[openFillerStartupModal] getFillerReminders 失敗:', error);
      showToast('載入提醒資訊失敗：' + error.message, true);
    })
    .getFillerReminders(managedProjectsStr);
}

function selectProjectAndStartLog(seqNo, shortName) {
  closeConfirmModal();

  // 設定工程選擇器
  const projectSelect = document.getElementById('logProjectSelect');
  projectSelect.value = seqNo;

  // 觸發工程變更事件以載入檢驗員
  handleProjectChange();

  // 滾動到表單頂部
  setTimeout(() => {
    document.querySelector('.glass-card').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 300);

  showToast(`✓ 已選擇：${shortName}，請填寫日誌`);
}

// ============================================
// 工程選擇變更
// ============================================
function handleProjectChange() {
  const projectSeqNo = document.getElementById('logProjectSelect').value;
  if (!projectSeqNo) return;

  // [修正] 確保檢驗員資料已載入
  if (!allInspectors || allInspectors.length === 0) {
    showLoading();
    window.GAS_API
      .withSuccessHandler(function (inspectors) {
        allInspectors = inspectors || [];
        handleProjectChange();
      })
      .withFailureHandler(function (error) {
        hideLoading();
        showToast('載入檢驗員資料失敗：' + error.message, true);
      })
      .getAllInspectors();
    return;
  }

  const dateString = document.getElementById('logDatePicker').value;

  showLoading();
  window.GAS_API
    .withSuccessHandler(function (inspectorIds) {
      hideLoading();

      // 顯示預設檢驗員
      if (inspectorIds && inspectorIds.length > 0) {
        displayDefaultInspectors(inspectorIds);
      } else {
        // 沒有預設檢驗員，直接顯示選擇介面
        showInspectorCheckboxes(inspectorIds);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('載入檢驗員失敗：' + error.message, true);
    })
    .getLastInspectors(projectSeqNo, dateString);
}

// 顯示預設檢驗員
function displayDefaultInspectors(inspectorIds) {
  const displayDiv = document.getElementById('inspectorDisplay');
  const displayText = document.getElementById('inspectorDisplayText');
  const checkboxDiv = document.getElementById('inspectorCheckboxes');
  const changeBtn = document.getElementById('changeInspectorBtn');

  // 取得檢驗員名稱
  const inspectorNames = inspectorIds.map(id => {
    const inspector = allInspectors.find(ins => ins.id === id);
    return inspector ? inspector.name : id;
  }).join('、');

  displayText.innerHTML = `✓ 預設檢驗員：<strong>${inspectorNames}</strong>`;

  // 顯示預設檢驗員區域
  displayDiv.style.display = 'block';
  changeBtn.style.display = 'inline-flex';

  // 隱藏複選框，但仍然渲染以保存選擇
  renderInspectorCheckboxes('inspectorCheckboxes', inspectorIds);
  checkboxDiv.style.display = 'none';
}

// 顯示檢驗員複選框
function showInspectorCheckboxes(selectedIds) {
  const displayDiv = document.getElementById('inspectorDisplay');
  const checkboxDiv = document.getElementById('inspectorCheckboxes');
  const changeBtn = document.getElementById('changeInspectorBtn');

  displayDiv.style.display = 'none';
  checkboxDiv.style.display = 'grid';
  changeBtn.style.display = 'none';

  renderInspectorCheckboxes('inspectorCheckboxes', selectedIds || []);
}

// 切換到修改模式
function toggleInspectorEditMode() {
  const checkboxDiv = document.getElementById('inspectorCheckboxes');
  const displayDiv = document.getElementById('inspectorDisplay');
  const changeBtn = document.getElementById('changeInspectorBtn');

  if (checkboxDiv.style.display === 'none') {
    // 切換到編輯模式
    checkboxDiv.style.display = 'grid';
    displayDiv.style.display = 'none';
    changeBtn.innerHTML = '<span>✓ 確認選擇</span>';
  } else {
    // 切換回顯示模式
    const selectedIds = getSelectedInspectors('inspectorCheckboxes');
    if (selectedIds.length === 0) {
      showToast('請至少選擇一位檢驗員', true);
      return;
    }
    displayDefaultInspectors(selectedIds);
  }
}

// ============================================
// 檢驗員複選框渲染（按部門分組）
// ============================================
function renderInspectorCheckboxes(containerId, selectedIds) {
  const container = document.getElementById(containerId);
  if (!container) return;

  container.innerHTML = '';

  // 按部門分組
  const inspectorsByDept = {};
  allInspectors.forEach(inspector => {
    const dept = inspector.dept || '未分類';
    if (!inspectorsByDept[dept]) {
      inspectorsByDept[dept] = [];
    }
    inspectorsByDept[dept].push(inspector);
  });

  // 修正5：部門排序 - 「隊」優先，「委外監造」最後
  const deptNames = Object.keys(inspectorsByDept);
  deptNames.sort((a, b) => {
    const aIsTeam = a.includes('隊');
    const bIsTeam = b.includes('隊');

    if (aIsTeam && !bIsTeam) return -1;
    if (!aIsTeam && bIsTeam) return 1;

    if (a === '委外監造' && b !== '委外監造') return 1;
    if (a !== '委外監造' && b === '委外監造') return -1;

    return a.localeCompare(b, 'zh-TW');
  });

  deptNames.forEach(dept => {
    const deptHeader = document.createElement('div');
    deptHeader.className = 'dept-group-header';
    deptHeader.innerHTML = `<span>🏢 ${dept}</span>`;
    container.appendChild(deptHeader);

    inspectorsByDept[dept].forEach(inspector => {
      const checkboxDiv = document.createElement('div');
      checkboxDiv.className = 'checkbox-item';

      const isChecked = selectedIds.includes(inspector.id);
      const checkboxId = `inspector_${inspector.id}`;

      // 顯示格式：名字(部門)
      const displayName = `${inspector.name}(${inspector.dept})`;

      checkboxDiv.innerHTML = `
        <input type="checkbox" id="${checkboxId}" value="${inspector.id}" ${isChecked ? 'checked' : ''}>
        <label for="${checkboxId}">${displayName}</label>
      `;

      container.appendChild(checkboxDiv);
    });
  });
}

function getSelectedInspectors(containerId) {
  const container = document.getElementById(containerId);
  if (!container) return [];

  // 即使容器被隱藏，仍然可以獲取選中的複選框
  const checkboxes = container.querySelectorAll('input[type="checkbox"]:checked');
  const selectedIds = Array.from(checkboxes).map(cb => cb.value);

  // 如果沒有選中的，可能是因為使用了預設顯示模式
  // 嘗試從所有複選框中找到已勾選的
  if (selectedIds.length === 0) {
    const allCheckboxes = container.querySelectorAll('input[type="checkbox"]');
    return Array.from(allCheckboxes)
      .filter(cb => cb.checked)
      .map(cb => cb.value);
  }

  return selectedIds;
}

// [新增] 載入檢驗員清單到篩選器
function loadInspectorsForFilter() {
  window.GAS_API
    .withSuccessHandler(function (inspectors) {
      const select = document.getElementById('summaryInspectorFilter');
      if (!select) return;

      // 保留 "全部檢驗員" 選項
      select.innerHTML = '<option value="all">全部檢驗員</option>';

      inspectors.forEach(inspector => {
        // 使用 ID 作為值
        const option = document.createElement('option');
        option.value = inspector.id; // [修正] idCode -> id
        option.textContent = inspector.name;
        select.appendChild(option);
      });
    })
    .withFailureHandler(function (err) {
      console.error('載入檢驗員失敗', err);
    })
    .getAllInspectors();
}

// ============================================
// 修正6&7：工項配對（含災害類型「其他」選項）
// ============================================
function addWorkItemPair() {
  const container = document.getElementById('workItemsContainer');
  const pairCount = container.querySelectorAll('.work-item-pair').length + 1;

  const pairDiv = document.createElement('div');
  pairDiv.className = 'work-item-pair';

  pairDiv.innerHTML = `
    <div class="pair-header">
      <div class="pair-number">工項 ${pairCount}</div>
      <button type="button" class="btn-remove" onclick="removeWorkItemPair(this)">✕ 移除</button>
    </div>
    
    <div class="form-group">
      <label class="form-label">
        <span class="label-icon">🛠️</span>
        <span>工作工項 <span class="required">*</span></span>
      </label>
      <textarea class="form-textarea work-item-text" rows="2" required placeholder="請描述主要工作內容"></textarea>
    </div>
    
    <div class="form-group">
      <label class="form-label">
        <span class="label-icon">⚠️</span>
        <span>災害類型（可多選） <span class="required">*</span></span>
      </label>
      <div class="disaster-checkboxes-grid">
        ${renderDisasterCheckboxes(pairCount)}
      </div>
    </div>
    
    <div class="form-group">
      <label class="form-label">
        <span class="label-icon">🛡️</span>
        <span>危害對策 <span class="required">*</span></span>
      </label>
      <textarea class="form-textarea countermeasures-text" rows="2" required placeholder="請描述具體的危害對策"></textarea>
    </div>
    
    <div class="form-group">
      <label class="form-label">
        <span class="label-icon">📍</span>
        <span>工作地點 <span class="required">*</span></span>
      </label>
      <input type="text" class="form-input work-location-text" required placeholder="請輸入具體工作地點">
    </div>
  `;

  container.appendChild(pairDiv);
  updatePairNumbers();
}

function renderDisasterCheckboxes(pairIndex) {
  let html = '';

  disasterOptions.forEach(disaster => {
    if (disaster.type === '其他') {
      // 修正6：「其他」選項特殊處理
      const checkboxId = `disaster_${pairIndex}_other`;
      html += `
        <div class="disaster-checkbox-item">
          <input type="checkbox" id="${checkboxId}" value="其他" onchange="toggleCustomDisasterInput(this, ${pairIndex})">
          <label for="${checkboxId}">
            <div class="disaster-checkbox-title">${disaster.icon} ${disaster.type}</div>
            <div class="disaster-checkbox-desc">${disaster.description}</div>
          </label>
        </div>
      `;
    } else {
      const checkboxId = `disaster_${pairIndex}_${disaster.type.replace(/[、\/]/g, '_')}`;
      html += `
        <div class="disaster-checkbox-item">
          <input type="checkbox" id="${checkboxId}" value="${disaster.type}">
          <label for="${checkboxId}">
            <div class="disaster-checkbox-title">${disaster.icon} ${disaster.type}</div>
            <div class="disaster-checkbox-desc">${disaster.description}</div>
          </label>
        </div>
      `;
    }
  });

  // 修正6：自訂災害類型輸入框
  html += `
    <div id="customDisasterContainer_${pairIndex}" style="display: none; grid-column: 1 / -1;">
      <input type="text" id="customDisasterInput_${pairIndex}" class="custom-disaster-input" placeholder="請輸入自訂災害類型">
    </div>
  `;

  return html;
}

// 修正6：切換自訂災害類型輸入框
function toggleCustomDisasterInput(checkbox, pairIndex) {
  const container = document.getElementById(`customDisasterContainer_${pairIndex}`);
  const input = document.getElementById(`customDisasterInput_${pairIndex}`);

  if (checkbox.checked) {
    container.style.display = 'block';
    input.focus();
  } else {
    container.style.display = 'none';
    input.value = '';
  }
}

function removeWorkItemPair(button) {
  button.closest('.work-item-pair').remove();
  updatePairNumbers();
}

function updatePairNumbers() {
  const pairs = document.querySelectorAll('.work-item-pair');
  pairs.forEach((pair, index) => {
    pair.querySelector('.pair-number').textContent = `工項 ${index + 1}`;
  });
}

function copyLastWorkItems() {
  const projectSeqNo = document.getElementById('logProjectSelect').value;

  if (!projectSeqNo) {
    showToast('請先選擇工程', true);
    return;
  }

  showLoading();
  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();
      if (result.success && result.data) {
        const lastLog = result.data;

        // 清空現有工作項目
        const container = document.getElementById('workItemsContainer');
        container.innerHTML = '';

        // 複製每個工作項目
        lastLog.workItems.forEach((item, index) => {
          addWorkItemPair();

          const pair = container.querySelectorAll('.work-item-pair')[index];

          // 填入工作工項
          pair.querySelector('.work-item-text').value = item.workItem || '';

          // 填入危害對策
          pair.querySelector('.countermeasures-text').value = item.countermeasures || '';

          // 填入工作地點
          pair.querySelector('.work-location-text').value = item.location || '';

          // 勾選災害類型
          if (item.disasters && item.disasters.length > 0) {
            item.disasters.forEach(disaster => {
              const pairIndex = index + 1;

              // 處理「其他」選項
              if (disaster.includes('其他')) {
                const otherCheckbox = pair.querySelector(`#disaster_${pairIndex}_other`);
                if (otherCheckbox) {
                  otherCheckbox.checked = true;
                  toggleCustomDisasterInput(otherCheckbox, pairIndex);

                  // 提取自訂災害內容
                  const customText = disaster.replace('其他：', '').trim();
                  const customInput = pair.querySelector(`#customDisasterInput_${pairIndex}`);
                  if (customInput) customInput.value = customText;
                }
              } else {
                // 一般災害類型
                const checkbox = pair.querySelector(`input[type="checkbox"][value="${disaster}"]`);
                if (checkbox) checkbox.checked = true;
              }
            });
          }
        });

        showToast(`✓ 已複製上次填寫內容（${lastLog.date}）`);
      } else {
        showToast(result.message || '此工程尚無歷史填報記錄', true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('複製失敗：' + error.message, true);
    })
    .getLastLogForProject(projectSeqNo);
}

function toggleWorkFields(hide) {
  const inspectorGroup = document.getElementById('inspectorGroup');
  const workersCountGroup = document.getElementById('workersCountGroup');
  const workItemsGroup = document.getElementById('workItemsGroup');

  if (hide) {
    inspectorGroup.style.display = 'none';
    workersCountGroup.style.display = 'none';
    workItemsGroup.style.display = 'none';
  } else {
    inspectorGroup.style.display = 'block';
    workersCountGroup.style.display = 'block';
    workItemsGroup.style.display = 'block';
  }
}

// ============================================
// 日誌提交
// ============================================
function handleDailyLogSubmit(event) {
  event.preventDefault();

  const logDate = document.getElementById('logDatePicker').value;
  const projectSeqNo = document.getElementById('logProjectSelect').value;

  // 驗證必填欄位
  if (!projectSeqNo) {
    showToast('請選擇工程', true);
    return;
  }

  const projectSelect = document.getElementById('logProjectSelect');
  const projectShortName = projectSelect.selectedOptions[0] ?
    projectSelect.selectedOptions[0].getAttribute('data-short-name') : '';

  const isHolidayNoWork = document.getElementById('isHolidayNoWork').checked;

  // 假日不施工
  if (isHolidayNoWork) {
    // [修正] 取得假日勾選細節
    const noWorkDetails = [];
    if (document.getElementById('checkSatNoWork').checked) noWorkDetails.push('星期六');
    if (document.getElementById('checkSunNoWork').checked) noWorkDetails.push('星期日');
    if (document.getElementById('checkHolidayNoWork').checked) noWorkDetails.push('例假');

    // 如果有勾選細項，組合字串；若無則保持預設
    const noWorkReason = noWorkDetails.length > 0
      ? `[假日不施工] ${noWorkDetails.join('、')}`
      : '🏖️ 假日不施工';

    const confirmMessage = `
      <p><strong>🏖️ 假日不施工設定</strong></p>
      <p><strong>📅 日期：</strong>${logDate}</p>
      <p><strong>🏗️ 工程：</strong>${document.getElementById('logProjectSelect').selectedOptions[0].text}</p>
      <p style="margin-top: 0.5rem; color: var(--primary);"><strong>包含：</strong>${noWorkDetails.length > 0 ? noWorkDetails.join('、') : '純假日不施工'}</p>
      <p style="margin-top: 1rem; color: var(--info);">確認提交此記錄嗎？</p>
    `;

    showConfirmModal(confirmMessage, function () {
      showLoading();
      executeSubmitDailyLog({
        logDate: logDate,
        projectSeqNo: projectSeqNo,
        projectShortName: projectShortName,
        isHolidayNoWork: true,
        // 我們將新的理由注入原本用來存工作的 workItems，這是後端解析字串的地方
        isHolidayWork: false,
        inspectorIds: [],
        workersCount: 0,
        workItems: [{
          workItem: noWorkReason,
          disasterTypes: [],
          countermeasures: '',
          location: ''
        }]
      });
      closeConfirmModal();
    });
    return;
  }

  // 一般日誌
  const inspectorIds = getSelectedInspectors('inspectorCheckboxes');
  const workersCount = document.getElementById('logWorkersCount').value;
  const isHolidayWork = document.getElementById('isHolidayWork').checked;

  if (inspectorIds.length === 0) {
    showToast('請至少選擇一位檢驗員', true);
    return;
  }

  if (!workersCount || workersCount <= 0) {
    showToast('請填寫施工人數', true);
    return;
  }

  const workItems = collectWorkItems();
  if (workItems.length === 0) {
    showToast('請至少填寫一組工項資料', true);
    return;
  }

  // 取得檢驗員名稱用於確認訊息
  const inspectorNames = inspectorIds.map(id => {
    const inspector = allInspectors.find(ins => ins.id === id);
    return inspector ? inspector.name : id;
  }).join('、');

  // 生成工項詳細資訊
  let workItemsDetail = '';
  workItems.forEach((item, index) => {
    const disasterText = (item.disasterTypes || []).join('、');
    const workItemName = item.workItem || '未命名工項';
    workItemsDetail += `
      <div style="margin-left: 1rem; margin-bottom: 0.5rem; padding: 0.5rem; background: var(--gray-50); border-left: 3px solid var(--primary); border-radius: var(--radius-sm);">
        <strong>工項 ${index + 1}：</strong>${workItemName}<br>
        <span style="font-size: 0.9rem; color: var(--text-secondary);">災害類型：${disasterText}</span>
      </div>
    `;
  });

  const confirmMessage = `
    <div style="max-height: 60vh; overflow-y: auto;">
      <p><strong>📅 日期：</strong>${logDate}</p>
      <p><strong>🏗️ 工程：</strong>${document.getElementById('logProjectSelect').selectedOptions[0].text}</p>
      ${isHolidayWork ? '<p style="color: var(--warning); font-weight: 700;">🏗️ 假日施工</p>' : ''}
      <p><strong>👥 檢驗員：</strong>${inspectorNames}</p>
      <p><strong>🧑‍🔧 施工人數：</strong>${workersCount} 人</p>
      <p style="margin-top: 1rem;"><strong>📝 工作項目明細：</strong></p>
      ${workItemsDetail}
      <p style="margin-top: 1.5rem; padding: 1rem; background: rgba(234, 88, 12, 0.1); border-radius: var(--radius); color: var(--warning-dark); font-weight: 600; text-align: center;">
        ⚠️ 確認提交日誌嗎？
      </p>
    </div>
  `;

  showConfirmModal(confirmMessage, function () {
    showLoading();
    executeSubmitDailyLog({
      logDate: logDate,
      projectSeqNo: projectSeqNo,
      projectShortName: projectShortName,
      isHolidayNoWork: false,
      isHolidayWork: isHolidayWork,
      inspectorIds: inspectorIds,
      workersCount: parseInt(workersCount),
      workItems: workItems
    });
    closeConfirmModal();
  });
}

function collectWorkItems() {
  const workItems = [];
  const pairs = document.querySelectorAll('.work-item-pair');

  pairs.forEach((pair, index) => {
    const workItemText = pair.querySelector('.work-item-text').value.trim();
    const countermeasuresText = pair.querySelector('.countermeasures-text').value.trim();
    const workLocationText = pair.querySelector('.work-location-text').value.trim();

    const disasterCheckboxes = pair.querySelectorAll('.disaster-checkboxes-grid input[type="checkbox"]:checked');
    let disasterTypes = Array.from(disasterCheckboxes).map(cb => cb.value);

    // 修正6：處理自訂災害類型
    if (disasterTypes.includes('其他')) {
      const pairIndex = index + 1;
      const customInput = document.getElementById(`customDisasterInput_${pairIndex}`);
      if (customInput && customInput.value.trim()) {
        // 移除「其他」，加入自訂類型
        disasterTypes = disasterTypes.filter(d => d !== '其他');
        disasterTypes.push(`其他:${customInput.value.trim()}`);
      }
    }

    if (workItemText && disasterTypes.length > 0 && countermeasuresText && workLocationText) {
      workItems.push({
        workItem: workItemText,
        disasterTypes: disasterTypes,
        countermeasures: countermeasuresText,
        workLocation: workLocationText
      });
    }
  });

  return workItems;
}

function executeSubmitDailyLog(data) {
  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();
      if (result.success) {
        showToast(`✓ ${result.message}`);




        document.getElementById('dailyLogForm').reset();
        const tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);

        // 修正：使用本地時間格式化
        const yyyy = tomorrow.getFullYear();
        const mm = String(tomorrow.getMonth() + 1).padStart(2, '0');
        const dd = String(tomorrow.getDate()).padStart(2, '0');
        document.getElementById('logDatePicker').value = `${yyyy}-${mm}-${dd}`;

        document.getElementById('workItemsContainer').innerHTML = '';




        checkAndShowHolidayAlert();
        loadUnfilledCount();
      } else {
        showToast('提交失敗：' + result.message, true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('伺服器錯誤：' + error.message, true);
    })
    .submitDailyLog(data);
}

// ============================================
// 未填寫數量統計
// ============================================
function loadUnfilledCount() {
  window.GAS_API
    .withSuccessHandler(function (data) {
      const container = document.getElementById('unfilledCardsContainer');

      if (data.unfilled > 0 && data.unfilledProjects) {
        // 渲染多個卡片
        renderUnfilledCards(data.unfilledProjects, data.date);
        container.style.display = 'block';
      } else {
        container.style.display = 'none';
      }
    })
    .withFailureHandler(function (error) {
      console.error('載入未填寫數量失敗：', error);
    })
    .getUnfilledCount();
}

// 渲染未填寫提醒卡片
function renderUnfilledCards(projects, date) {
  const container = document.getElementById('unfilledCardsContainer');
  container.innerHTML = '';

  const gridDiv = document.createElement('div');
  gridDiv.className = 'unfilled-cards-grid';

  projects.forEach((project, index) => {
    const card = document.createElement('div');
    card.className = 'unfilled-card';
    card.onclick = function () {
      // 點擊卡片後填入該工程並開始填報
      fillProjectAndStartLog(project.seqNo, project.shortName);
    };

    card.innerHTML = `
      <div class="unfilled-card-header">
        <div class="unfilled-card-icon">⚠️</div>
        <div class="unfilled-card-title">待填報 #${index + 1}</div>
      </div>
      <div class="unfilled-card-body">
        <div class="unfilled-card-info">
          <strong>日期：</strong><span>${date}</span>
        </div>
        <div class="unfilled-card-info">
          <strong>工程：</strong><span>${project.shortName}</span>
        </div>
        <div class="unfilled-card-info">
          <strong>承攬商：</strong><span>${project.contractor}</span>
        </div>
        <div class="unfilled-card-info">
          <strong>部門：</strong><span>${project.dept}</span>
        </div>
      </div>
      <div class="unfilled-card-footer">
        點擊開始填報
      </div>
    `;

    gridDiv.appendChild(card);
  });

  container.appendChild(gridDiv);
}

// 填入工程並開始填報
function fillProjectAndStartLog(seqNo, shortName) {
  const projectSelect = document.getElementById('logProjectSelect');
  projectSelect.value = seqNo;

  // 觸發工程變更事件以載入檢驗員
  handleProjectChange();

  // 滾動到表單頂部
  document.querySelector('.glass-card').scrollIntoView({ behavior: 'smooth', block: 'start' });

  showToast(`已選擇：${shortName}，請繼續填寫日誌`);
}

// ============================================
// 總表功能
// ============================================




function setupSummaryDate() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);

  // 修正：使用本地時間格式化
  const yyyy = tomorrow.getFullYear();
  const mm = String(tomorrow.getMonth() + 1).padStart(2, '0');
  const dd = String(tomorrow.getDate()).padStart(2, '0');
  document.getElementById('summaryDatePicker').value = `${yyyy}-${mm}-${dd}`;
}




function formatInspectorDisplay(inspectorText, inspectorDetails) {
  if (!inspectorText || inspectorText === '-') {
    return '-';
  }

  // 如果有詳細資訊（inspectorDetails 是陣列，包含 {name, profession, dept}）
  if (inspectorDetails && Array.isArray(inspectorDetails) && inspectorDetails.length > 0) {
    return inspectorDetails.map(ins => {
      const isOutsource = ins.dept === '委外監造';
      return `${ins.name}(${ins.profession})${isOutsource ? '委' : ''}`;
    }).join('、');
  }

  // 如果沒有詳細資訊，直接返回原文字
  return inspectorText;
}

function loadSummaryReport() {
  const dateString = document.getElementById('summaryDatePicker').value;
  if (!dateString) {
    showToast('請選擇日期', true);
    return;
  }

  const filterStatus = document.querySelector('input[name="summaryStatusFilter"]:checked').value;
  const filterDept = document.getElementById('summaryDeptFilter').value;
  const filterInspector = document.getElementById('summaryInspectorFilter').value; // [新增] 取得檢驗員篩選值

  showLoading();
  window.GAS_API

    .withSuccessHandler(function (summaryData) {
      hideLoading();
      currentSummaryData = summaryData;
      renderSummaryTable(summaryData);

      // [新增] 同步渲染手機版卡片視圖
      renderMobileSummary(summaryData);

      // 更新訪客模式卡片
      if (isGuestMode) {
        updateGuestSummaryCards(dateString, summaryData);
      }
    })







    .withFailureHandler(function (error) {
      hideLoading();
      showToast('載入總表失敗：' + error.message, true);
    })
    .getDailySummaryReport(dateString, filterStatus, filterDept, filterInspector, isGuestMode, currentUserInfo);
}

// [新增] 渲染檢驗員篩選選單
function renderInspectorFilter() {
  const select = document.getElementById('summaryInspectorFilter');
  if (!select) return;

  // 清空選項 (保留預設)
  select.innerHTML = '<option value="all">全部檢驗員</option>';

  // 取得所有檢驗員 (已在全域變數 allInspectors)
  // 排序：隊 -> 委外
  const sortedInspectors = [...allInspectors].sort((a, b) => {
    const aIsTeam = a.dept && a.dept.includes('隊');
    const bIsTeam = b.dept && b.dept.includes('隊');
    if (aIsTeam && !bIsTeam) return -1;
    if (!aIsTeam && bIsTeam) return 1;
    return a.dept.localeCompare(b.dept, 'zh-TW') || a.name.localeCompare(b.name, 'zh-TW');
  });

  sortedInspectors.forEach(ins => {
    if (ins.status === 'active') { // 只顯示啟用中的
      const option = document.createElement('option');
      option.value = ins.id;
      option.textContent = `${ins.name} (${ins.dept})`;
      select.appendChild(option);
    }
  });
}

// [新增] 渲染檢驗員篩選選單
function renderInspectorFilter() {
  const select = document.getElementById('summaryInspectorFilter');
  if (!select) return;

  // 清空選項 (保留預設)
  select.innerHTML = '<option value="all">全部檢驗員</option>';

  // 取得所有檢驗員 (已在全域變數 allInspectors)
  // 排序：隊 -> 委外
  const sortedInspectors = [...allInspectors].sort((a, b) => {
    const aIsTeam = a.dept && a.dept.includes('隊');
    const bIsTeam = b.dept && b.dept.includes('隊');
    if (aIsTeam && !bIsTeam) return -1;
    if (!aIsTeam && bIsTeam) return 1;
    return a.dept.localeCompare(b.dept, 'zh-TW') || a.name.localeCompare(b.name, 'zh-TW');
  });

  sortedInspectors.forEach(ins => {
    if (ins.status === 'active') { // 只顯示啟用中的
      const option = document.createElement('option');
      option.value = ins.id;
      option.textContent = `${ins.name} (${ins.dept})`;
      select.appendChild(option);
    }
  });
}

function updateGuestSummaryCards(dateString, summaryData) {
  // 更新日期顯示
  const guestDateDisplay = document.getElementById('guestDateDisplay');
  if (guestDateDisplay) {
    guestDateDisplay.textContent = dateString;
  }

  // 計算當日有填報的工程數（排除未填報的）
  const filledCount = summaryData.filter(row => row.hasFilled).length;

  // 更新工程數顯示
  const guestProjectCount = document.getElementById('guestProjectCount');
  if (guestProjectCount) {
    guestProjectCount.textContent = filledCount;
  }
}

function renderSummaryTable(summaryData) {
  const tbody = document.getElementById('summaryTableBody');
  const thead = document.getElementById('summaryTableHead');

  // 根據登入狀態調整表頭
  if (isGuestMode) {
    // 訪客模式：不顯示序號和操作
    thead.innerHTML = `
      <tr>
        <th>工程名稱</th>
        <th>承攬商</th>
        <th>部門</th>
        <th>檢驗員</th>
        <th>工地負責人</th>
        <th>職安人員</th>
        <th>工作地址</th>
        <th>施工人數</th>
        <th>主要工作項目</th>
        <th>主要災害類型</th>
      </tr>
    `;
  } else {
    // 已登入：顯示完整欄位
    thead.innerHTML = `
      <tr>
        <th>序號</th>
        <th>工程名稱</th>
        <th>承攬商</th>
        <th>部門</th>
        <th>檢驗員</th>
        <th>工地負責人</th>
        <th>職安人員</th>
        <th>工作地址</th>
        <th>施工人數</th>
        <th>主要工作項目</th>
        <th>主要災害類型</th>
        <th>操作</th>
      </tr>
    `;
  }

  const colspanCount = isGuestMode ? 10 : 12;

  if (summaryData.length === 0) {
    tbody.innerHTML = `<tr><td colspan="${colspanCount}" class="text-muted">查無資料</td></tr>`;
    return;
  }

  tbody.innerHTML = '';

  summaryData.forEach(row => {
    const isClickable = !isGuestMode && (row.hasFilled || row.projectStatus === '施工中');
    const inspectorText = formatInspectorDisplay(row.inspectors, row.inspectorDetails) || '-';
    const workersCountText = row.isHolidayNoWork ? '-' : (row.workersCount || '-');
    const holidayBadge = row.isHolidayWork ?
      '<span class="status-badge-holiday-work" style="margin-left: 0.5rem;">🏗️ 假日施工</span>' : '';

    // 格式化職安人員顯示：姓名(證照) 或 姓名
    const safetyOfficerText = row.safetyLicense ? `${row.safetyOfficer}(${row.safetyLicense})` : row.safetyOfficer;

    // 判斷是否需要分割工作項目
    let workItems = [];

    if (row.isHolidayNoWork) {
      // 假日不施工：單一列
      workItems = [{
        text: '🏖️ 假日不施工',
        disasters: '無',
        isBadge: true
      }];
    } else if (row.workItems && row.workItems.length > 0) {
      // 有工作項目：每個工作項目對應自己的災害類型
      workItems = row.workItems.map(wi => ({
        text: wi.text || '',
        disasters: wi.disasters && wi.disasters.length > 0 ? wi.disasters.join('、') : '無',
        countermeasures: wi.countermeasures || '',
        workLocation: wi.workLocation || ''
      }));
    } else {
      // 未填寫：單一列
      workItems = [{
        text: '未填寫',
        disasters: '未填寫',
        isEmpty: true
      }];
    }

    // 計算需要合併的列數
    const rowspan = workItems.length;

    // 生成每個工作項目的列
    workItems.forEach((workItem, idx) => {
      const tr = document.createElement('tr');

      // 添加樣式類別
      if (row.hasFilled) {
        tr.classList.add('filled-row');
      } else if (row.projectStatus === '施工中') {
        tr.classList.add('empty-row');
      }

      // [新增] 標記最後一個項目，用於 CSS 邊框處理
      if (idx === workItems.length - 1) {
        tr.classList.add('is-last-item');
      }

      // 第一列包含所有合併的欄位
      if (idx === 0) {
        if (isClickable) {
          tr.style.cursor = 'pointer';
          tr.onclick = function () {
            if (row.hasFilled) {
              openEditSummaryLogModal(row);
            } else {
              openLogEntryForProject(row.seqNo, row.fullName);
            }
          };
        }

        if (isGuestMode) {
          // 訪客模式
          tr.innerHTML = `
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);"><strong>${row.fullName}</strong>${holidayBadge}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${row.contractor}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${row.dept}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${inspectorText}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${row.resp}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${safetyOfficerText}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${truncateText(row.address, 20)}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${workersCountText}</td>
            <td style="border-top: 2px solid rgba(0,0,0,0.15); white-space: normal; word-wrap: break-word;">${workItem.isBadge ? '<span class="status-badge-holiday-no">' + workItem.text + '</span>' : (workItem.isEmpty ? '<span class="unfilled-text">' + workItem.text + '</span>' : workItem.text)}</td>
            <td style="border-top: 2px solid rgba(0,0,0,0.15); white-space: normal; word-wrap: break-word;">${workItem.isEmpty ? '<span class="unfilled-text">' + workItem.disasters + '</span>' : workItem.disasters}</td>
          `;
        } else {
          // 已登入模式
          tr.innerHTML = `
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${row.seqNo}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);"><strong>${row.fullName}</strong>${holidayBadge}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${row.contractor}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${row.dept}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${inspectorText}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${row.resp}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${safetyOfficerText}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${truncateText(row.address, 20)}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">${workersCountText}</td>
            <td style="border-top: 2px solid rgba(0,0,0,0.15); white-space: normal; word-wrap: break-word;">${workItem.isBadge ? '<span class="status-badge-holiday-no">' + workItem.text + '</span>' : (workItem.isEmpty ? '<span class="unfilled-text">' + workItem.text + '</span>' : workItem.text)}</td>
            <td style="border-top: 2px solid rgba(0,0,0,0.15); white-space: normal; word-wrap: break-word;">${workItem.isEmpty ? '<span class="unfilled-text">' + workItem.disasters + '</span>' : workItem.disasters}</td>
            <td rowspan="${rowspan}" style="border-right: 1px solid rgba(0,0,0,0.08);">
              ${isClickable ? (row.hasFilled ? '<button class="btn-mini">✏️ 編輯</button>' : '<button class="btn-mini">📝 填報</button>') : '-'}
            </td>
          `;
        }
      } else {
        // 後續列只包含工作項目和災害類型
        tr.innerHTML = `
          <td style="white-space: normal; word-wrap: break-word;">${workItem.isBadge ? '<span class="status-badge-holiday-no">' + workItem.text + '</span>' : (workItem.isEmpty ? '<span class="unfilled-text">' + workItem.text + '</span>' : workItem.text)}</td>
          <td style="white-space: normal; word-wrap: break-word;">${workItem.isEmpty ? '<span class="unfilled-text">' + workItem.disasters + '</span>' : workItem.disasters}</td>
        `;
      }

      tbody.appendChild(tr);
    });
  });
}

/**
 * [新增] 渲染手機版總表卡片視圖
 * 包含欄位：工程名稱, 承攬商, 檢驗員, 工地負責人, 職安人員, 工作地址, 施工人數, 主要工作項目, 主要災害類型
 */
function renderMobileSummary(summaryData) {
  const container = document.getElementById('summaryMobileView');
  if (!container) return;

  if (summaryData.length === 0) {
    container.innerHTML = '<div style="text-align: center; padding: 2rem; color: #666;">查無資料</div>';
    return;
  }

  let html = '';

  summaryData.forEach(row => {
    // 1. 狀態判斷
    const isFilled = row.hasFilled;
    const statusClass = isFilled ? 'status-filled' : 'status-active';
    const badgeClass = isFilled ? 'm-badge-success' : 'm-badge-warning';
    const badgeText = isFilled ? '已填寫' : '未填寫';

    // 2. 資料格式化
    const inspectorText = formatInspectorDisplay(row.inspectors, row.inspectorDetails) || '-';
    // 職安人員若有證照則顯示
    const safetyDisplay = row.safetyLicense ? `${row.safetyOfficer}(${row.safetyLicense})` : (row.safetyOfficer || '-');

    // 3. 處理工項與災害類型
    let workItemsHtml = '';
    if (row.isHolidayNoWork) {
      workItemsHtml = '<div class="m-work-item" style="color: #0891b2; font-weight: bold;">🏖️ 假日不施工</div>';
    } else if (row.workItems && row.workItems.length > 0) {
      row.workItems.forEach((item, idx) => {
        const disasterText = (item.disasters && item.disasters.length > 0) ? item.disasters : '無';
        workItemsHtml += `
          <div class="m-work-item">
            <div class="m-work-desc">${idx + 1}. ${item.text}</div>
            <div class="m-disaster">⚠️ 災害：${disasterText}</div>
          </div>
        `;
      });
    } else {
      workItemsHtml = '<div class="m-work-item" style="color: #9ca3af;">尚無工項資料</div>';
    }

    // 4. 組裝卡片 HTML
    html += `
      <div class="mobile-summary-card ${statusClass}">
        <div class="m-card-header">
          <div class="m-project-name">${row.fullName}</div>
          <div class="m-header-row">
             <div class="m-contractor">🏢 ${row.contractor}</div>
             <div class="m-badge ${badgeClass}">${badgeText}</div>
          </div>
        </div>
        
        <div class="m-body">
          <div class="m-row">
            <span class="m-label">檢驗員</span>
            <span class="m-value">${inspectorText}</span>
          </div>
          <div class="m-row">
            <span class="m-label">工地負責人</span>
            <span class="m-value">${row.resp || '-'} ${row.respPhone ? '📱' : ''}</span>
          </div>
          <div class="m-row">
            <span class="m-label">職安人員</span>
            <span class="m-value">${safetyDisplay} ${row.safetyPhone ? '📱' : ''}</span>
          </div>
          <div class="m-row">
            <span class="m-label">工作地址</span>
            <span class="m-value">${row.address || '-'}</span>
          </div>
          
          ${isFilled ? `
          <div class="m-row">
            <span class="m-label">施工人數</span>
            <span class="m-value" style="font-weight: bold;">${row.workersCount} 人</span>
          </div>
          ` : ''}
        </div>

        <div class="m-work-section">
          <span class="m-section-title">主要工作項目 & 災害類型</span>
          ${workItemsHtml}
        </div>
        
        ${!isGuestMode ? `
          <button class="m-action-btn" onclick="${isFilled ? `openEditSummaryLogModal(${JSON.stringify(row).replace(/"/g, '&quot;')})` : `openLogEntryForProject('${row.seqNo}', '${row.fullName}')`}">
            ${isFilled ? '✏️ 編輯日誌' : '📝 填寫日誌'}
          </button>
        ` : ''}
      </div>
    `;
  });

  container.innerHTML = html;
}







function switchSummaryMode(mode) {
  document.querySelectorAll('.mode-btn').forEach(btn => btn.classList.remove('active'));
  document.querySelector(`.mode-btn[data-mode="${mode}"]`).classList.add('active');

  if (mode === 'calendar') {
    document.getElementById('summaryListView').style.display = 'none';
    document.getElementById('summaryCalendarView').style.display = 'block';
    renderCalendar();
  } else {
    document.getElementById('summaryListView').style.display = 'block';
    document.getElementById('summaryCalendarView').style.display = 'none';
  }
}

// 繼續在第二部分...

// ============================================
// [新增] 批次假日設定功能
// ============================================
function showBatchHolidayModal() {
  const modal = document.getElementById('batchHolidayModal');
  const projectList = document.getElementById('batchProjectList');

  // 設定預設日期 (明日)
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);

  const yyyy = tomorrow.getFullYear();
  const mm = String(tomorrow.getMonth() + 1).padStart(2, '0');
  const dd = String(tomorrow.getDate()).padStart(2, '0');

  // 若無值預設填入
  if (!document.getElementById('batchStartDate').value) {
    document.getElementById('batchStartDate').value = `${yyyy}-${mm}-${dd}`;
  }
  if (!document.getElementById('batchEndDate').value) {
    document.getElementById('batchEndDate').value = `${yyyy}-${mm}-${dd}`;
  }

  projectList.innerHTML = '<div style="padding: 1rem; text-align: center; color: var(--text-muted);">載入中...</div>';
  modal.style.display = 'flex';

  // 載入施工中工程
  window.GAS_API
    .withSuccessHandler(function (projects) {
      projectList.innerHTML = '';
      if (projects.length === 0) {
        projectList.innerHTML = '<div style="padding: 1rem; text-align: center;">無施工中工程</div>';
        return;
      }

      projects.forEach(proj => {
        const div = document.createElement('div');
        div.className = 'checkbox-item';
        const id = `batch_proj_${proj.seqNo}`;
        div.innerHTML = `
            <input type="checkbox" id="${id}" name="batchProject" value="${proj.seqNo}" checked>
            <label for="${id}">${proj.fullName} (${proj.shortName})</label>
          `;
        projectList.appendChild(div);
      });
    })
    .withFailureHandler(function (err) {
      projectList.innerHTML = '<div style="color: red;">載入失敗</div>';
    })
    .getActiveProjects(); // 需確認後端有此函式，或使用 getInitialData 中的 projects 過濾
}

function closeBatchHolidayModal() {
  document.getElementById('batchHolidayModal').style.display = 'none';
}

function toggleBatchAllProjects(checkbox) {
  const checkboxes = document.querySelectorAll('input[name="batchProject"]');
  checkboxes.forEach(cb => cb.checked = checkbox.checked);
}

function submitBatchHoliday() {
  const startDate = document.getElementById('batchStartDate').value;
  const endDate = document.getElementById('batchEndDate').value;

  if (!startDate || !endDate) {
    showToast('請選擇日期範圍', true);
    return;
  }

  if (startDate > endDate) {
    showToast('開始日期不能晚於結束日期', true);
    return;
  }

  const targetDays = [];
  if (document.getElementById('batchCheckSun').checked) targetDays.push(0); // 週日
  if (document.getElementById('batchCheckSat').checked) targetDays.push(6); // 週六

  if (targetDays.length === 0) {
    showToast('請至少選擇一天 (週六或週日)', true);
    return;
  }

  const selectedProjects = [];
  document.querySelectorAll('input[name="batchProject"]:checked').forEach(cb => {
    selectedProjects.push(cb.value);
  });

  if (selectedProjects.length === 0) {
    showToast('請至少選擇一個工程', true);
    return;
  }

  const confirmMessage = `
       <p><strong>📅 批次設定假日不施工</strong></p>
       <p>日期：${startDate} ~ ${endDate}</p>
       <p>對象：${targetDays.includes(6) ? '週六' : ''} ${targetDays.includes(0) ? '週日' : ''}</p>
       <p>工程數：${selectedProjects.length} 個</p>
       <p style="margin-top: 1rem; color: var(--warning);">⚠️ 此操作將覆蓋現有日誌，確認執行？</p>
    `;

  showConfirmModal(confirmMessage, function () {
    showLoading();
    window.GAS_API
      .withSuccessHandler(function (result) {
        hideLoading();
        if (result.success) {
          showToast(`✓ ${result.message}`);
          closeBatchHolidayModal();
          // 重新整理總表
          loadSummaryReport();
        } else {
          showToast('設定失敗：' + result.message, true);
        }
      })
      .withFailureHandler(function (err) {
        hideLoading();
        showToast('系統錯誤：' + err.message, true);
      })
      .batchSubmitHolidayLogs(startDate, endDate, targetDays, selectedProjects);

    closeConfirmModal();
  });
}

// ============================================
// 編輯總表日誌

// ============================================
function openEditSummaryLogModal(rowData) {
  const dateString = document.getElementById('summaryDatePicker').value;

  document.getElementById('editSummaryLogDate').value = dateString;
  document.getElementById('editSummaryLogProjectSeqNo').value = rowData.seqNo;
  document.getElementById('editSummaryLogProjectName').textContent = rowData.fullName;
  document.getElementById('editSummaryIsHolidayNoWork').value = rowData.isHolidayNoWork ? 'yes' : 'no';

  // 顯示/隱藏切換按鈕區域
  const toggleContainer = document.getElementById('editHolidayToggleContainer');
  if (!toggleContainer) {
    // 如果容器不存在，動態創建
    const header = document.getElementById('editSummaryLogModal').querySelector('.modal-header');
    const container = document.createElement('div');
    container.id = 'editHolidayToggleContainer';
    container.style.marginTop = '10px';
    container.style.padding = '10px';
    container.style.background = '#f3f4f6';
    container.style.borderRadius = '8px';

    container.innerHTML = `
        <label class="toggle-switch">
          <input type="checkbox" id="editSwitchHolidayNoWork" onchange="toggleEditHolidayNoWork(this)">
          <span class="toggle-slider"></span>
          <span class="toggle-label">🏖️ 假日不施工</span>
        </label>
      `;
    header.parentNode.insertBefore(container, header.nextSibling);
  }

  // 設定初始狀態
  const isNoWork = rowData.isHolidayNoWork;
  document.getElementById('editSwitchHolidayNoWork').checked = isNoWork;
  document.getElementById('editSummaryIsHolidayNoWork').value = isNoWork ? 'yes' : 'no';

  toggleEditFields(isNoWork);

  // 填入資料（即使隱藏也要填入，以便切換時顯示）
  renderInspectorCheckboxes('editInspectorCheckboxes', rowData.inspectorIds || []);
  document.getElementById('editSummaryWorkersCount').value = rowData.workersCount || 0;
  renderEditWorkItemsList(rowData.workItems || []);
  document.getElementById('editIsHolidayWork').checked = rowData.isHolidayWork || false;
  document.getElementById('editSummaryReason').value = '';
  document.getElementById('editSummaryLogModal').style.display = 'flex';
}

// [新增] 切換編輯模式的假日不施工狀態
function toggleEditHolidayNoWork(checkbox) {
  const isNoWork = checkbox.checked;
  document.getElementById('editSummaryIsHolidayNoWork').value = isNoWork ? 'yes' : 'no';
  toggleEditFields(isNoWork);
}

// [新增] 控制編輯欄位顯示
function toggleEditFields(isNoWork) {
  if (isNoWork) {
    document.getElementById('editHolidayNoWorkBadge').style.display = 'block';
    document.getElementById('editHolidayWorkBadge').style.display = 'none';
    document.getElementById('editHolidayWorkToggle').style.display = 'none';
    document.getElementById('editSummaryInspectorGroup').style.display = 'none';
    document.getElementById('editSummaryWorkersGroup').style.display = 'none';
    document.getElementById('editSummaryWorkItemsGroup').style.display = 'none';
  } else {
    document.getElementById('editHolidayNoWorkBadge').style.display = 'none';
    checkEditHolidayWorkStatus(); // 更新假日施工標記
    document.getElementById('editHolidayWorkToggle').style.display = 'block';
    document.getElementById('editSummaryInspectorGroup').style.display = 'block';
    document.getElementById('editSummaryWorkersGroup').style.display = 'block';
    document.getElementById('editSummaryWorkItemsGroup').style.display = 'block';
  }
}

function checkEditHolidayWorkStatus() {
  const isHolidayWork = document.getElementById('editIsHolidayWork').checked;
  document.getElementById('editHolidayWorkBadge').style.display = isHolidayWork ? 'block' : 'none';
}

function closeEditSummaryLogModal() {
  document.getElementById('editSummaryLogModal').style.display = 'none';
}

function renderEditWorkItemsList(workItems) {
  const container = document.getElementById('editWorkItemsList');
  container.innerHTML = '';

  workItems.forEach((item, index) => {
    const itemDiv = document.createElement('div');
    itemDiv.className = 'edit-work-item';

    const disasters = item.disasters.join('、');

    itemDiv.innerHTML = `
      <div class="edit-work-item-header">工項 ${index + 1}</div>
      
      <div class="form-group">
        <label class="form-label">
          <span class="label-icon">🛠️</span>
          <span>工作工項</span>
        </label>
        <textarea class="form-textarea edit-work-item-text" rows="2">${item.text}</textarea>
      </div>
      
      <div class="form-group">
        <label class="form-label">
          <span class="label-icon">⚠️</span>
          <span>災害類型</span>
        </label>
        <div class="disaster-checkboxes-grid-small">
          ${renderEditDisasterCheckboxes(index, item.disasters)}
        </div>
      </div>
      
      <div class="form-group">
        <label class="form-label">
          <span class="label-icon">🛡️</span>
          <span>危害對策</span>
        </label>
        <textarea class="form-textarea edit-countermeasures-text" rows="2">${item.countermeasures}</textarea>
      </div>
      
      <div class="form-group">
        <label class="form-label">
          <span class="label-icon">📍</span>
          <span>工作地點</span>
        </label>
        <input type="text" class="form-input edit-work-location-text" value="${item.workLocation}">
      </div>
      
      <button type="button" class="btn-remove" onclick="removeEditWorkItem(this)" style="width: 100%;">✕ 移除此工項</button>
    `;

    container.appendChild(itemDiv);
  });

  // 新增工項按鈕
  const addBtn = document.createElement('button');
  addBtn.type = 'button';
  addBtn.className = 'btn btn-secondary';
  addBtn.style.width = '100%';
  addBtn.innerHTML = '<span class="btn-icon">➕</span><span>新增工項</span>';
  addBtn.onclick = addEditWorkItem;
  container.appendChild(addBtn);
}

function renderEditDisasterCheckboxes(itemIndex, selectedDisasters) {
  let html = '';

  disasterOptions.forEach(disaster => {
    if (disaster.type === '其他') return; // 編輯模式暫不支援新增「其他」

    const isChecked = selectedDisasters.includes(disaster.type);
    const checkboxId = `edit_disaster_${itemIndex}_${disaster.type.replace(/[、\/]/g, '_')}`;

    html += `
      <div class="disaster-checkbox-item-small">
        <input type="checkbox" id="${checkboxId}" value="${disaster.type}" ${isChecked ? 'checked' : ''}>
        <label for="${checkboxId}">${disaster.icon} ${disaster.type}</label>
      </div>
    `;
  });

  return html;
}

function addEditWorkItem() {
  const container = document.getElementById('editWorkItemsList');
  const addBtn = container.querySelector('.btn-secondary');

  const itemCount = container.querySelectorAll('.edit-work-item').length + 1;

  const itemDiv = document.createElement('div');
  itemDiv.className = 'edit-work-item';

  itemDiv.innerHTML = `
    <div class="edit-work-item-header">工項 ${itemCount}</div>
    
    <div class="form-group">
      <label class="form-label">
        <span class="label-icon">🛠️</span>
        <span>工作工項</span>
      </label>
      <textarea class="form-textarea edit-work-item-text" rows="2" placeholder="請描述主要工作內容"></textarea>
    </div>
    
    <div class="form-group">
      <label class="form-label">
        <span class="label-icon">⚠️</span>
        <span>災害類型</span>
      </label>
      <div class="disaster-checkboxes-grid-small">
        ${renderEditDisasterCheckboxes(itemCount, [])}
      </div>
    </div>
    
    <div class="form-group">
      <label class="form-label">
        <span class="label-icon">🛡️</span>
        <span>危害對策</span>
      </label>
      <textarea class="form-textarea edit-countermeasures-text" rows="2" placeholder="請描述具體的危害對策"></textarea>
    </div>
    
    <div class="form-group">
      <label class="form-label">
        <span class="label-icon">📍</span>
        <span>工作地點</span>
      </label>
      <input type="text" class="form-input edit-work-location-text" placeholder="請輸入具體工作地點">
    </div>
    
    <button type="button" class="btn-remove" onclick="removeEditWorkItem(this)" style="width: 100%;">✕ 移除此工項</button>
  `;

  container.insertBefore(itemDiv, addBtn);
  updateEditWorkItemNumbers();
}

function removeEditWorkItem(button) {
  button.closest('.edit-work-item').remove();
  updateEditWorkItemNumbers();
}

function updateEditWorkItemNumbers() {
  const items = document.querySelectorAll('.edit-work-item');
  items.forEach((item, index) => {
    item.querySelector('.edit-work-item-header').textContent = `工項 ${index + 1}`;
  });
}

function confirmEditSummaryLog() {
  const dateString = document.getElementById('editSummaryLogDate').value;
  const projectSeqNo = document.getElementById('editSummaryLogProjectSeqNo').value;
  const isHolidayNoWork = document.getElementById('editSummaryIsHolidayNoWork').value === 'yes';
  const reason = document.getElementById('editSummaryReason').value.trim();

  if (!reason) {
    showToast('請填寫修改原因', true);
    return;
  }

  if (isHolidayNoWork) {
    // 假日不施工不需要其他資料
    const confirmMessage = `
      <p><strong>📝 修改日誌</strong></p>
      <p><strong>📅 日期：</strong>${dateString}</p>
      <p><strong>🏗️ 工程：</strong>${document.getElementById('editSummaryLogProjectName').textContent}</p>
      <p><strong>📝 修改原因：</strong>${reason}</p>
      <p style="margin-top: 1rem; color: var(--warning);">⚠️ 確認修改嗎？</p>
    `;

    showConfirmModal(confirmMessage, function () {
      showLoading();
      executeUpdateSummaryLog({
        dateString: dateString,
        projectSeqNo: projectSeqNo,
        isHolidayNoWork: true,
        reason: reason
      });
      closeConfirmModal();
    });
    return;
  }

  const inspectorIds = getSelectedInspectors('editInspectorCheckboxes');
  const workersCount = document.getElementById('editSummaryWorkersCount').value;
  const isHolidayWork = document.getElementById('editIsHolidayWork').checked;

  if (inspectorIds.length === 0) {
    showToast('請至少選擇一位檢驗員', true);
    return;
  }

  if (!workersCount || workersCount <= 0) {
    showToast('請填寫施工人數', true);
    return;
  }

  const workItems = collectEditWorkItems();
  if (workItems.length === 0) {
    showToast('請至少填寫一組工項資料', true);
    return;
  }

  const confirmMessage = `
    <p><strong>📝 修改日誌</strong></p>
    <p><strong>📅 日期：</strong>${dateString}</p>
    <p><strong>🏗️ 工程：</strong>${document.getElementById('editSummaryLogProjectName').textContent}</p>
    <p><strong>📝 修改原因：</strong>${reason}</p>
    <p style="margin-top: 1rem; color: var(--warning);">⚠️ 確認修改嗎？</p>
  `;

  showConfirmModal(confirmMessage, function () {
    showLoading();
    executeUpdateSummaryLog({
      dateString: dateString,
      projectSeqNo: projectSeqNo,
      isHolidayNoWork: false,
      isHolidayWork: isHolidayWork,
      inspectorIds: inspectorIds,
      workersCount: parseInt(workersCount),
      workItems: workItems,
      reason: reason
    });
    closeConfirmModal();
  });
}

function collectEditWorkItems() {
  const workItems = [];
  const items = document.querySelectorAll('.edit-work-item');

  items.forEach(item => {
    const workItemText = item.querySelector('.edit-work-item-text').value.trim();
    const countermeasuresText = item.querySelector('.edit-countermeasures-text').value.trim();
    const workLocationText = item.querySelector('.edit-work-location-text').value.trim();

    const disasterCheckboxes = item.querySelectorAll('.disaster-checkboxes-grid-small input[type="checkbox"]:checked');
    const disasterTypes = Array.from(disasterCheckboxes).map(cb => cb.value);

    if (workItemText && disasterTypes.length > 0 && countermeasuresText && workLocationText) {
      workItems.push({
        workItem: workItemText,
        disasterTypes: disasterTypes,
        countermeasures: countermeasuresText,
        workLocation: workLocationText
      });
    }
  });

  return workItems;
}

function executeUpdateSummaryLog(data) {
  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();
      if (result.success) {
        showToast(`✓ ${result.message}`);
        closeEditSummaryLogModal();
        loadSummaryReport();
      } else {
        showToast('修改失敗：' + result.message, true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('伺服器錯誤：' + error.message, true);
    })
    .updateDailySummaryLog(data);
}

// ============================================
// [新增] 批次假日設定相關功能
// ============================================
function showBatchHolidayModal() {
  const modal = document.getElementById('batchHolidayModal');
  if (!modal) return;

  // 預設日期：明日
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = tomorrow.toISOString().split('T')[0];

  document.getElementById('batchStartDate').value = tomorrowStr;
  document.getElementById('batchEndDate').value = tomorrowStr;

  // 渲染工程列表
  renderBatchProjectList();

  modal.style.display = 'flex';
}

function closeBatchHolidayModal() {
  document.getElementById('batchHolidayModal').style.display = 'none';
}

function renderBatchProjectList() {
  const container = document.getElementById('batchProjectList');
  if (!container) return;

  container.innerHTML = '';

  // 取得所有施工中工程 (全域變數 allProjects)
  const activeProjects = allProjects.filter(p => p.projectStatus === '施工中');

  activeProjects.forEach(project => {
    const div = document.createElement('div');
    div.className = 'checkbox-item';
    div.innerHTML = `
        <input type="checkbox" name="batchProject" value="${project.seqNo}" checked>
        <label>${project.shortName || project.fullName}</label>
      `;
    container.appendChild(div);
  });
}

function toggleBatchAllProjects(checkbox) {
  const checkboxes = document.querySelectorAll('input[name="batchProject"]');
  checkboxes.forEach(cb => cb.checked = checkbox.checked);
}

function submitBatchHoliday() {
  const startDate = document.getElementById('batchStartDate').value;
  const endDate = document.getElementById('batchEndDate').value;

  if (!startDate || !endDate) {
    showToast('請選擇開始與結束日期', true);
    return;
  }

  if (startDate > endDate) {
    showToast('開始日期不能晚於結束日期', true);
    return;
  }

  const targetDays = [];
  if (document.getElementById('batchCheckSat').checked) targetDays.push(6);
  if (document.getElementById('batchCheckSun').checked) targetDays.push(0);

  if (targetDays.length === 0) {
    showToast('請至少選擇一個星期幾 (週六或週日)', true);
    return;
  }

  const selectedProjects = [];
  document.querySelectorAll('input[name="batchProject"]:checked').forEach(cb => {
    selectedProjects.push(cb.value);
  });

  if (selectedProjects.length === 0) {
    showToast('請至少選擇一個工程', true);
    return;
  }

  if (!confirm(`確定要為 ${selectedProjects.length} 個工程設定假日不施工嗎？日期：${startDate} ~ ${endDate}`)) {
    return;
  }

  showLoading();
  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();
      if (result.success) {
        showToast(result.message);
        closeBatchHolidayModal();
        // 如果當前在日誌填報頁面且日期在範圍內，可能需要重新整理
      } else {
        showToast(result.message, true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('批次設定失敗：' + error.message, true);
    })
    .batchSubmitHolidayLogs(startDate, endDate, targetDays, selectedProjects);
}

function copyPreviousDayData() {
  const projectSeqNo = document.getElementById('editSummaryLogProjectSeqNo').value;
  const currentDate = document.getElementById('editSummaryLogDate').value;

  showLoading();
  window.GAS_API
    .withSuccessHandler(function (prevLog) {
      hideLoading();
      if (prevLog) {
        renderInspectorCheckboxes('editInspectorCheckboxes', prevLog.inspectorIds);
        document.getElementById('editSummaryWorkersCount').value = prevLog.workersCount;
        showToast('✓ 已複製前一天資料');
      } else {
        showToast('找不到前一天的日誌資料', true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('複製失敗：' + error.message, true);
    })
    .getPreviousDayLog(projectSeqNo, currentDate);
}

// ============================================
// 工程設定功能
// ============================================
function loadAndRenderProjectCards() {
  // showLoading 已在 switchTab 中處理

  // 如果 allInspectors 還沒載入，先載入檢驗員資料
  if (!allInspectors || allInspectors.length === 0) {
    window.GAS_API
      .withSuccessHandler(function (inspectors) {
        allInspectors = inspectors;
        // 載入檢驗員資料後，再載入工程資料
        loadProjectsData();
      })
      .withFailureHandler(function (error) {
        hideLoading();
        showToast('載入檢驗員資料失敗：' + error.message, true);
      })
      .getAllInspectors();
  } else {
    // allInspectors 已經載入，直接載入工程資料
    loadProjectsData();
  }
}

function loadProjectsData() {
  window.GAS_API
    .withSuccessHandler(function (projects) {
      // 不要在這裡 hideLoading，留給 renderProjectCards 完成後再關閉
      allProjectsData = projects;

      let filteredProjects = projects;

      // 填表人只能看到自己的工程
      if (currentUserInfo && currentUserInfo.role === '填表人') {
        const managedProjects = currentUserInfo.managedProjects || [];
        filteredProjects = filteredProjects.filter(p => managedProjects.includes(p.seqNo));
      }

      const filterStatus = document.querySelector('input[name="projectStatusFilter"]:checked').value;
      const filterDept = document.getElementById('projectDeptFilter').value;

      if (filterStatus === 'active') {
        filteredProjects = filteredProjects.filter(p => p.projectStatus === '施工中');
      }

      if (filterDept && filterDept !== 'all') {
        filteredProjects = filteredProjects.filter(p => p.dept === filterDept);
      }

      // renderProjectCards 會在渲染完成後 hideLoading
      renderProjectCards(filteredProjects);
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('載入工程資料失敗：' + error.message, true);
    })
    .getAllProjects();
}

function formatProjectDefaultInspector(inspectorIds) {
  // 如果沒有預設檢驗員，返回紅色警告
  if (!inspectorIds || inspectorIds.length === 0) {
    return '<span style="color: #dc2626; font-weight: 600;">⚠️ 未設定預設檢驗員</span>';
  }

  // 從 allInspectors 全局變數中查找檢驗員資訊
  if (!allInspectors || allInspectors.length === 0) {
    return inspectorIds.join('、');
  }

  const inspectorNames = inspectorIds.map(id => {
    const inspector = allInspectors.find(ins => ins.id === id);
    if (inspector) {
      const isOutsource = inspector.dept === '委外監造';
      return `${inspector.name}(${inspector.profession})${isOutsource ? '委' : ''}`;
    }
    return id;
  });

  return inspectorNames.join('、');
}

function renderProjectCards(projects) {
  const container = document.getElementById('projectCardsContainer');

  if (projects.length === 0) {
    container.innerHTML = '<div class="text-muted" style="grid-column: 1 / -1; text-align: center; padding: 2rem;">查無工程資料</div>';
    hideLoading();
    return;
  }

  // 先不要清空容器，保持 loading 顯示
  // container.innerHTML = '';

  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = tomorrow.toISOString().substring(0, 10);

  window.GAS_API
    .withSuccessHandler(function (summaryData) {
      // 收到資料後才清空容器並渲染
      container.innerHTML = '';

      const filledSeqNos = new Set();
      summaryData.forEach(row => {
        if (row.hasFilled) {
          filledSeqNos.add(row.seqNo);
        }
      });

      projects.forEach(project => {
        const card = document.createElement('div');
        card.className = 'project-card editable';

        const hasFilled = filledSeqNos.has(project.seqNo);
        const isActive = project.projectStatus === '施工中';

        if (hasFilled) {
          card.classList.add('filled');
        } else if (isActive) {
          card.classList.add('not-filled');
        } else {
          card.classList.add('inactive');
        }

        card.onclick = function () {
          openEditProjectModal(project);
        };

        const statusClass = `status-${project.projectStatus}`;
        const statusIcon = hasFilled ? '✓' : (isActive ? '⚠' : '○');
        const statusText = hasFilled ? '已填寫' : (isActive ? '未填寫' : '非施工中');
        const statusColor = hasFilled ? 'filled' : (isActive ? 'not-filled' : 'inactive');

        // 格式化預設檢驗員顯示
        const defaultInspectorDisplay = formatProjectDefaultInspector(project.defaultInspectors);

        card.innerHTML = `
          <div class="card-edit-hint">點擊編輯</div>

          <div class="project-card-header">
            <div class="project-card-seq-name">
              <div class="project-card-icon">🏗️</div>
              <div class="project-card-seq">${project.seqNo}</div>
            </div>
            <div class="project-status-badge ${statusClass}">${project.projectStatus}</div>
          </div>

          <div class="project-card-body">
            <div class="project-card-title">${project.fullName}</div>

            <div class="project-card-info">
              <div class="info-row">
                <span class="info-label">承攬商：</span>
                <span class="info-value">${project.contractor}</span>
              </div>
              <div class="info-row">
                <span class="info-label">部門：</span>
                <span class="info-value">${project.dept}</span>
              </div>
              <div class="info-row">
                <span class="info-label">預設檢驗員：</span>
                <span class="info-value">${defaultInspectorDisplay}</span>
              </div>
              <div class="info-row">
                <span class="info-label">工地負責人：</span>
                <span class="info-value">${project.resp}</span>
              </div>
              <div class="info-row">
                <span class="info-label">職安人員：</span>
                <span class="info-value">${project.safetyOfficer}</span>
              </div>
            </div>

            ${project.remark ? `<div class="project-remark">📝 ${project.remark}</div>` : ''}
          </div>

          <div class="project-card-footer ${statusColor}">
            <span class="status-icon">${statusIcon}</span>
            <span>${tomorrowStr} - ${statusText}</span>
          </div>
        `;

        container.appendChild(card);
      });

      // 卡片渲染完成後關閉 loading
      hideLoading();
    })
    .withFailureHandler(function (error) {
      console.error('載入填報狀況失敗：', error);

      // 清空容器
      container.innerHTML = '';

      projects.forEach(project => {
        const card = document.createElement('div');
        card.className = 'project-card editable';

        if (project.projectStatus !== '施工中') {
          card.classList.add('inactive');
        }

        card.onclick = function () {
          openEditProjectModal(project);
        };

        const statusClass = `status-${project.projectStatus}`;
        const defaultInspectorDisplay = formatProjectDefaultInspector(project.defaultInspectors);

        card.innerHTML = `
          <div class="card-edit-hint">點擊編輯</div>

          <div class="project-card-header">
            <div class="project-card-seq-name">
              <div class="project-card-icon">🏗️</div>
              <div class="project-card-seq">${project.seqNo}</div>
            </div>
            <div class="project-status-badge ${statusClass}">${project.projectStatus}</div>
          </div>

          <div class="project-card-body">
            <div class="project-card-title">${project.fullName}</div>

            <div class="project-card-info">
              <div class="info-row">
                <span class="info-label">承攬商：</span>
                <span class="info-value">${project.contractor}</span>
              </div>
              <div class="info-row">
                <span class="info-label">部門：</span>
                <span class="info-value">${project.dept}</span>
              </div>
              <div class="info-row">
                <span class="info-label">預設檢驗員：</span>
                <span class="info-value">${defaultInspectorDisplay}</span>
              </div>
              <div class="info-row">
                <span class="info-label">工地負責人：</span>
                <span class="info-value">${project.resp}</span>
              </div>
              <div class="info-row">
                <span class="info-label">職安人員：</span>
                <span class="info-value">${project.safetyOfficer}</span>
              </div>
            </div>

            ${project.remark ? `<div class="project-remark">📝 ${project.remark}</div>` : ''}
          </div>
          
          <div class="project-card-footer inactive">
            <span class="status-icon">○</span>
            <span>工程資料</span>
          </div>
        `;

        container.appendChild(card);
      });

      // 失敗時也要關閉 loading
      hideLoading();
    })
    .getDailySummaryReport(tomorrowStr, 'all', 'all', isGuestMode, currentUserInfo);
}

function openEditProjectModal(project) {
  document.getElementById('editProjectSeqNo').value = project.seqNo;
  document.getElementById('editProjectName').textContent = `${project.seqNo} - ${project.fullName}`;

  document.getElementById('editResp').value = project.resp || '';
  document.getElementById('editRespPhone').value = project.respPhone || '';
  document.getElementById('editSafetyOfficer').value = project.safetyOfficer || '';
  document.getElementById('editSafetyPhone').value = project.safetyPhone || '';
  document.getElementById('editSafetyLicense').value = project.safetyLicense || '';
  document.getElementById('editProjectStatus').value = project.projectStatus || '施工中';
  document.getElementById('editStatusRemark').value = project.remark || '';

  // 觸發狀態變更事件以顯示/隱藏備註欄
  const statusSelect = document.getElementById('editProjectStatus');
  statusSelect.dispatchEvent(new Event('change'));

  // 預設檢驗員
  renderInspectorCheckboxes('editDefaultInspectorCheckboxes', project.defaultInspectors);

  document.getElementById('editProjectReason').value = '';

  document.getElementById('editProjectModal').style.display = 'flex';
}

function closeEditProjectModal() {
  document.getElementById('editProjectModal').style.display = 'none';
}

function confirmEditProject() {
  const projectSeqNo = document.getElementById('editProjectSeqNo').value;
  const resp = document.getElementById('editResp').value.trim();
  const respPhone = document.getElementById('editRespPhone').value.trim();
  const safetyOfficer = document.getElementById('editSafetyOfficer').value.trim();
  const safetyPhone = document.getElementById('editSafetyPhone').value.trim();
  const safetyLicense = document.getElementById('editSafetyLicense').value;
  const projectStatus = document.getElementById('editProjectStatus').value;
  const statusRemark = document.getElementById('editStatusRemark').value.trim();
  const reason = document.getElementById('editProjectReason').value.trim();

  if (!resp || !safetyOfficer || !reason) {
    showToast('請填寫所有必填欄位', true);
    return;
  }

  if (projectStatus !== '施工中' && !statusRemark) {
    showToast('工程狀態非「施工中」時，備註欄為必填', true);
    return;
  }

  const defaultInspectors = getSelectedInspectors('editDefaultInspectorCheckboxes');

  // 取得當前使用者名稱
  const modifierName = currentUserInfo ? currentUserInfo.name : '訪客';

  // 取得預設檢驗員名稱
  const defaultInspectorNames = defaultInspectors.map(id => {
    const inspector = allInspectorsWithStatus.find(ins => ins.id === id);
    return inspector ? `${inspector.name} (${id})` : id;
  }).join('、');
  const inspectorDisplay = defaultInspectorNames || '未設定';

  const safetyDisplay = safetyLicense ? `${safetyOfficer}(${safetyLicense})` : safetyOfficer;

  const confirmMessage = `
    <p><strong>⚙️ 修改工程資料</strong></p>
    <p><strong>🏗️ 工程：</strong>${document.getElementById('editProjectName').textContent}</p>
    <p><strong>👤 修改人員：</strong>${modifierName}</p>
    <p><strong>👷 工地負責人：</strong>${resp}</p>
    <p><strong>🦺 職安人員：</strong>${safetyDisplay}</p>
    <p><strong>📊 工程狀態：</strong>${projectStatus}</p>
    ${statusRemark ? `<p><strong>📝 備註：</strong>${statusRemark}</p>` : ''}
    <p><strong>👨‍🔧 預設檢驗員：</strong>${inspectorDisplay}</p>
    <p><strong>📝 修改原因：</strong>${reason}</p>
    <p style="margin-top: 1rem; color: var(--warning);">⚠️ 確認修改嗎？</p>
  `;

  showConfirmModal(confirmMessage, function () {
    showLoading();
    window.GAS_API
      .withSuccessHandler(function (result) {
        hideLoading();
        if (result.success) {
          showToast(`✓ ${result.message}`);
          closeEditProjectModal();
          loadAndRenderProjectCards();
        } else {
          showToast('修改失敗：' + result.message, true);
        }
      })
      .withFailureHandler(function (error) {
        hideLoading();
        showToast('伺服器錯誤：' + error.message, true);
      })
      .updateProjectInfo({
        projectSeqNo: projectSeqNo,
        resp: resp,
        respPhone: respPhone,
        safetyOfficer: safetyOfficer,
        safetyPhone: safetyPhone,
        safetyLicense: safetyLicense,
        projectStatus: projectStatus,
        statusRemark: statusRemark,
        defaultInspectors: defaultInspectors,
        reason: reason
      });
    closeConfirmModal();
  });
}

// ============================================
// 檢驗員管理功能
// ============================================
function loadInspectorManagement() {
  // showLoading 已在 switchTab 中處理
  window.GAS_API
    .withSuccessHandler(function (inspectors) {
      hideLoading();
      allInspectorsWithStatus = inspectors;

      const filterStatus = document.querySelector('input[name="inspectorStatusFilter"]:checked').value;
      const filterDept = document.getElementById('inspectorDeptFilter').value;

      let filteredInspectors = inspectors;

      if (filterStatus === 'active') {
        filteredInspectors = filteredInspectors.filter(ins => ins.status === '啟用');
      }

      if (filterDept && filterDept !== 'all') {
        filteredInspectors = filteredInspectors.filter(ins => ins.dept === filterDept);
      }

      renderInspectorCards(filteredInspectors);
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('載入檢驗員資料失敗：' + error.message, true);
    })
    .getAllInspectorsWithStatus();
}

function renderInspectorCards(inspectors) {
  const container = document.getElementById('inspectorCardsContainer');

  if (inspectors.length === 0) {
    container.innerHTML = '<div class="text-muted" style="grid-column: 1 / -1; text-align: center; padding: 2rem;">查無檢驗員資料</div>';
    return;
  }

  container.innerHTML = '';

  // 按部門分組
  const inspectorsByDept = {};
  inspectors.forEach(inspector => {
    const dept = inspector.dept || '未分類';
    if (!inspectorsByDept[dept]) {
      inspectorsByDept[dept] = [];
    }
    inspectorsByDept[dept].push(inspector);
  });

  // 渲染每個部門
  Object.keys(inspectorsByDept).sort().forEach(dept => {
    // 部門標題
    const deptHeader = document.createElement('div');
    deptHeader.style.cssText = 'grid-column: 1 / -1; margin-top: 1.5rem; margin-bottom: 1rem; padding: 0.75rem 1.5rem; background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%); color: white; border-radius: 12px; font-weight: 700; font-size: 1.1rem; display: flex; align-items: center; gap: 0.75rem; box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);';
    deptHeader.innerHTML = `
      <span style="font-size: 1.5rem;">🏢</span>
      <span>${dept}</span>
      <span style="margin-left: auto; background: rgba(255,255,255,0.2); padding: 0.25rem 0.75rem; border-radius: 20px; font-size: 0.9rem;">${inspectorsByDept[dept].length} 位</span>
    `;
    container.appendChild(deptHeader);

    // 渲染該部門的檢驗員卡片
    inspectorsByDept[dept].forEach(inspector => {
      const card = document.createElement('div');
      card.className = 'inspector-card';

      const isActive = inspector.status === '啟用';

      if (!isActive) {
        card.classList.add('inactive');
      }

      const statusClass = isActive ? 'active' : 'inactive';

      card.innerHTML = `
      <div class="inspector-card-header">
        <div class="inspector-card-id">${inspector.id}</div>
        <div class="inspector-card-status ${statusClass}">${inspector.status}</div>
      </div>
      
      <div class="inspector-card-body">
        <div class="inspector-card-name">${inspector.name}</div>
        
        <div class="inspector-card-info">
          <div class="inspector-info-row">
            <span class="inspector-info-icon">🏢</span>
            <span class="inspector-info-label">部門：</span>
            <span class="inspector-info-value">${inspector.dept}</span>
          </div>
          <div class="inspector-info-row">
            <span class="inspector-info-icon">🎓</span>
            <span class="inspector-info-label">職稱：</span>
            <span class="inspector-info-value">${inspector.title}</span>
          </div>
          <div class="inspector-info-row">
            <span class="inspector-info-icon">🔧</span>
            <span class="inspector-info-label">專業：</span>
            <span class="inspector-info-value">${inspector.profession}</span>
          </div>
          ${inspector.phone ? `
          <div class="inspector-info-row">
            <span class="inspector-info-icon">📞</span>
            <span class="inspector-info-label">電話：</span>
            <span class="inspector-info-value">${inspector.phone}</span>
          </div>
          ` : ''}
        </div>
      </div>
      
      <div class="inspector-card-footer">
        <button class="btn btn-primary btn-mini" onclick="openEditInspectorModal('${inspector.id}')">
          <span>✏️ 編輯</span>
        </button>
        ${isActive ?
          `<button class="btn btn-danger btn-mini" onclick="confirmDeactivateInspector('${inspector.id}')">
            <span>⏸️ 停用</span>
          </button>` :
          `<button class="btn btn-success btn-mini" onclick="confirmActivateInspector('${inspector.id}')">
            <span>▶️ 啟用</span>
          </button>`
        }
      </div>
    `;

      container.appendChild(card);
    });
  });
}

function generateInspectorId(dept) {
  const prefix = DEPT_CODE_MAP[dept];
  if (!prefix) {
    return null;
  }

  // 找出該部門現有的最大編號
  const deptInspectors = allInspectorsWithStatus.filter(ins => ins.dept === dept);
  let maxNumber = 0;

  deptInspectors.forEach(ins => {
    if (ins.id && ins.id.startsWith(prefix)) {
      const numPart = ins.id.substring(prefix.length);
      const num = parseInt(numPart, 10);
      if (!isNaN(num) && num > maxNumber) {
        maxNumber = num;
      }
    }
  });

  // 生成新編號（兩位數）
  const newNumber = maxNumber + 1;
  return `${prefix}${String(newNumber).padStart(2, '0')}`;
}

function openAddInspectorModal() {
  document.getElementById('addInspectorDept').value = '';
  document.getElementById('addInspectorName').value = '';
  document.getElementById('addInspectorTitle').value = '';
  document.getElementById('addInspectorProfession').value = '';
  document.getElementById('addInspectorPhone').value = '';
  document.getElementById('addInspectorReason').value = '';
  document.getElementById('addInspectorIdPreview').textContent = '請先選擇部門';

  document.getElementById('addInspectorModal').style.display = 'flex';
}

function updateInspectorIdPreview() {
  const dept = document.getElementById('addInspectorDept').value;
  const previewElement = document.getElementById('addInspectorIdPreview');

  if (!dept) {
    previewElement.textContent = '請先選擇部門';
    previewElement.style.color = 'var(--text-muted)';
    return;
  }

  const newId = generateInspectorId(dept);
  if (newId) {
    previewElement.textContent = newId;
    previewElement.style.color = 'var(--primary)';
    previewElement.style.fontWeight = 'bold';
  } else {
    previewElement.textContent = '無法生成編號';
    previewElement.style.color = 'var(--error)';
  }
}

function closeAddInspectorModal() {
  document.getElementById('addInspectorModal').style.display = 'none';
}

function confirmAddInspector() {
  const dept = document.getElementById('addInspectorDept').value;
  const name = document.getElementById('addInspectorName').value.trim();
  const title = document.getElementById('addInspectorTitle').value.trim();
  const profession = document.getElementById('addInspectorProfession').value;
  const phone = document.getElementById('addInspectorPhone').value.trim();
  const reason = document.getElementById('addInspectorReason').value.trim();

  if (!dept || !name || !title || !profession || !reason) {
    showToast('請填寫所有必填欄位', true);
    return;
  }

  // 生成檢驗員編號
  const newId = generateInspectorId(dept);
  if (!newId) {
    showToast('無法生成檢驗員編號，請檢查部門設定', true);
    return;
  }

  const confirmMessage = `
    <p><strong>➕ 新增檢驗員</strong></p>
    <p><strong>🔢 編號：</strong><span style="color: var(--primary); font-weight: bold;">${newId}</span></p>
    <p><strong>🏢 部門：</strong>${dept}</p>
    <p><strong>👤 姓名：</strong>${name}</p>
    <p><strong>🎓 職稱：</strong>${title}</p>
    <p><strong>🔧 專業：</strong>${profession}</p>
    ${phone ? `<p><strong>📞 電話：</strong>${phone}</p>` : ''}
    <p><strong>📝 新增原因：</strong>${reason}</p>
    <p style="margin-top: 1rem; color: var(--warning);">⚠️ 確認新增嗎？</p>
  `;

  showConfirmModal(confirmMessage, function () {
    showLoading();
    window.GAS_API
      .withSuccessHandler(function (result) {
        hideLoading();
        if (result.success) {
          showToast(`✓ ${result.message}`);
          closeAddInspectorModal();
          loadInspectorManagement();
          loadInitialData();
        } else {
          showToast('新增失敗：' + result.message, true);
        }
      })
      .withFailureHandler(function (error) {
        hideLoading();
        showToast('伺服器錯誤：' + error.message, true);
      })
      .addInspector({
        id: newId,
        dept: dept,
        name: name,
        title: title,
        profession: profession,
        phone: phone,
        reason: reason
      });
    closeConfirmModal();
  });
}

function openEditInspectorModal(inspectorId) {
  const inspector = allInspectorsWithStatus.find(ins => ins.id === inspectorId);
  if (!inspector) {
    showToast('找不到檢驗員資料', true);
    return;
  }

  document.getElementById('editInspectorId').value = inspector.id;
  document.getElementById('editInspectorIdDisplay').textContent = inspector.id;
  document.getElementById('editInspectorDept').value = inspector.dept;
  document.getElementById('editInspectorName').value = inspector.name;
  document.getElementById('editInspectorTitle').value = inspector.title;
  document.getElementById('editInspectorProfession').value = inspector.profession;
  document.getElementById('editInspectorPhone').value = inspector.phone || '';
  document.getElementById('editInspectorReason').value = '';

  document.getElementById('editInspectorModal').style.display = 'flex';
}

function closeEditInspectorModal() {
  document.getElementById('editInspectorModal').style.display = 'none';
}

function confirmEditInspector() {
  const id = document.getElementById('editInspectorId').value;
  const dept = document.getElementById('editInspectorDept').value;
  const name = document.getElementById('editInspectorName').value.trim();
  const title = document.getElementById('editInspectorTitle').value.trim();
  const profession = document.getElementById('editInspectorProfession').value;
  const phone = document.getElementById('editInspectorPhone').value.trim();
  const reason = document.getElementById('editInspectorReason').value.trim();

  if (!dept || !name || !title || !profession || !reason) {
    showToast('請填寫所有必填欄位', true);
    return;
  }

  const confirmMessage = `
    <p><strong>✏️ 修改檢驗員資料</strong></p>
    <p><strong>🆔 編號：</strong>${id}</p>
    <p><strong>🏢 部門：</strong>${dept}</p>
    <p><strong>👤 姓名：</strong>${name}</p>
    <p><strong>🎓 職稱：</strong>${title}</p>
    <p><strong>🔧 專業：</strong>${profession}</p>
    ${phone ? `<p><strong>📞 電話：</strong>${phone}</p>` : ''}
    <p><strong>📝 修改原因：</strong>${reason}</p>
    <p style="margin-top: 1rem; color: var(--warning);">⚠️ 確認修改嗎？</p>
  `;

  showConfirmModal(confirmMessage, function () {
    showLoading();
    window.GAS_API
      .withSuccessHandler(function (result) {
        hideLoading();
        if (result.success) {
          showToast(`✓ ${result.message}`);
          closeEditInspectorModal();
          loadInspectorManagement();
          loadInitialData();
        } else {
          showToast('修改失敗：' + result.message, true);
        }
      })
      .withFailureHandler(function (error) {
        hideLoading();
        showToast('伺服器錯誤：' + error.message, true);
      })
      .updateInspector({
        id: id,
        dept: dept,
        name: name,
        title: title,
        profession: profession,
        phone: phone,
        reason: reason
      });
    closeConfirmModal();
  });
}

function confirmDeactivateInspector(inspectorId) {
  const inspector = allInspectorsWithStatus.find(ins => ins.id === inspectorId);
  if (!inspector) return;

  showLoading();
  window.GAS_API
    .withSuccessHandler(function (usage) {
      hideLoading();

      let warningMessage = '';
      if (usage.isUsed) {
        warningMessage = '<div style="background: rgba(245, 158, 11, 0.1); padding: 1rem; border-radius: 0.5rem; margin: 1rem 0;">';
        warningMessage += '<p style="color: var(--warning); font-weight: 600;">⚠️ 使用狀況提醒</p>';

        if (usage.projects.length > 0) {
          warningMessage += '<p style="margin-top: 0.5rem;">此檢驗員為以下工程的預設檢驗員：</p>';
          warningMessage += '<ul style="margin: 0.5rem 0; padding-left: 1.5rem;">';
          usage.projects.forEach(proj => {
            warningMessage += `<li>${proj.seqNo} - ${proj.name}</li>`;
          });
          warningMessage += '</ul>';
        }

        if (usage.logs.length > 0) {
          warningMessage += `<p style="margin-top: 0.5rem;">最近30天內有 ${usage.logs.length} 筆日誌記錄使用此檢驗員</p>`;
        }

        warningMessage += '</div>';
      }

      const reason = prompt(`${warningMessage ? warningMessage + '' : ''}請輸入停用原因：`);

      if (!reason || reason.trim() === '') {
        showToast('未輸入停用原因，已取消操作', false);
        return;
      }

      const confirmMessage = `
        <p><strong>⏸️ 停用檢驗員</strong></p>
        <p><strong>🆔 編號：</strong>${inspector.id}</p>
        <p><strong>👤 姓名：</strong>${inspector.name}</p>
        ${warningMessage}
        <p><strong>📝 停用原因：</strong>${reason}</p>
        <p style="margin-top: 1rem; color: var(--danger);">⚠️ 確認停用嗎？</p>
      `;

      showConfirmModal(confirmMessage, function () {
        showLoading();
        window.GAS_API
          .withSuccessHandler(function (result) {
            hideLoading();
            if (result.success) {
              showToast(`✓ ${result.message}`);
              loadInspectorManagement();
              loadInitialData();
            } else {
              showToast('停用失敗：' + result.message, true);
            }
          })
          .withFailureHandler(function (error) {
            hideLoading();
            showToast('伺服器錯誤：' + error.message, true);
          })
          .deactivateInspector({
            id: inspectorId,
            reason: reason.trim()
          });
        closeConfirmModal();
      });
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('檢查使用狀況失敗：' + error.message, true);
    })
    .checkInspectorUsage(inspectorId);
}

function confirmActivateInspector(inspectorId) {
  const inspector = allInspectorsWithStatus.find(ins => ins.id === inspectorId);
  if (!inspector) return;

  const reason = prompt('請輸入啟用原因：');

  if (!reason || reason.trim() === '') {
    showToast('未輸入啟用原因，已取消操作', false);
    return;
  }

  const confirmMessage = `
    <p><strong>▶️ 啟用檢驗員</strong></p>
    <p><strong>🆔 編號：</strong>${inspector.id}</p>
    <p><strong>👤 姓名：</strong>${inspector.name}</p>
    <p><strong>📝 啟用原因：</strong>${reason}</p>
    <p style="margin-top: 1rem; color: var(--success);">✓ 確認啟用嗎？</p>
  `;

  showConfirmModal(confirmMessage, function () {
    showLoading();
    window.GAS_API
      .withSuccessHandler(function (result) {
        hideLoading();
        if (result.success) {
          showToast(`✓ ${result.message}`);
          loadInspectorManagement();
          loadInitialData();
        } else {
          showToast('啟用失敗：' + result.message, true);
        }
      })
      .withFailureHandler(function (error) {
        hideLoading();
        showToast('伺服器錯誤：' + error.message, true);
      })
      .activateInspector({
        id: inspectorId,
        reason: reason.trim()
      });
    closeConfirmModal();
  });
}

// 繼續在第三部分...

// ============================================
// 日誌填報狀況總覽
// ============================================
function loadLogStatus() {
  // showLoading 已在 switchTab 中處理
  window.GAS_API
    .withSuccessHandler(function (data) {
      hideLoading();
      renderLogStatus(data);
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('載入填報狀況失敗：' + error.message, true);
    })
    .getDailyLogStatus();
}

function renderLogStatus(data) {
  document.getElementById('statusTomorrowDate').textContent = data.tomorrowDate;

  document.getElementById('statTotalProjects').textContent = data.totalProjects;
  document.getElementById('statFilledTomorrow').textContent = data.filledTomorrow;

  const completionRate = data.totalProjects > 0 ?
    Math.round((data.filledTomorrow / data.totalProjects) * 100) : 0;
  document.getElementById('statCompletionRate').textContent = completionRate + '%';

  const container = document.getElementById('deptCardsContainer');
  container.innerHTML = '';

  const deptNames = Object.keys(data.byDept);

  // 修正5：部門排序
  deptNames.sort((a, b) => {
    const aIsTeam = a.includes('隊');
    const bIsTeam = b.includes('隊');

    if (aIsTeam && !bIsTeam) return -1;
    if (!aIsTeam && bIsTeam) return 1;

    if (a === '委外監造' && b !== '委外監造') return 1;
    if (a !== '委外監造' && b === '委外監造') return -1;

    return a.localeCompare(b, 'zh-TW');
  });

  deptNames.forEach(deptName => {
    const deptData = data.byDept[deptName];
    const deptRate = deptData.total > 0 ?
      Math.round((deptData.filled.length / deptData.total) * 100) : 0;

    const card = document.createElement('div');
    card.className = 'dept-card';

    const cardId = `deptCard_${deptName.replace(/\s/g, '_')}`;

    card.innerHTML = `
      <div class="dept-card-header" onclick="toggleDeptCard('${cardId}')">
        <div class="dept-card-title">
          <span class="dept-icon">🏢</span>
          <span class="dept-name">${deptName}</span>
        </div>
        <div class="dept-card-stats">
          <span>${deptData.filled.length} / ${deptData.total}</span>
          ${deptData.missing.length > 0 ?
        `<span class="dept-missing">缺 ${deptData.missing.length}</span>` : ''}
          <span class="dept-rate" style="background: ${deptRate === 100 ? 'var(--success)' : (deptRate >= 50 ? 'var(--warning)' : 'var(--danger)')}">${deptRate}%</span>
          <span class="dept-toggle" id="${cardId}_toggle">▼</span>
        </div>
      </div>
      <div class="dept-card-content" id="${cardId}_content" style="display: none;">
        ${deptData.missing.length > 0 ? `
          <div style="margin-bottom: 1rem;">
            <strong style="color: var(--danger);">⚠️ 未填寫 (${deptData.missing.length})</strong>
            ${deptData.missing.map(proj => `
              <div class="project-list-item">
                <div class="project-info">
                  <div class="project-name">${proj.fullName}</div>
                  <div class="project-meta">序號：${proj.seqNo} | 承攬商：${proj.contractor}</div>
                </div>
                <span class="status-badge">未填寫</span>
              </div>
            `).join('')}
          </div>
        ` : ''}
        ${deptData.filled.length > 0 ? `
          <div>
            <strong style="color: var(--success);">✓ 已填寫 (${deptData.filled.length})</strong>
            ${deptData.filled.map(proj => `
              <div class="project-list-item">
                <div class="project-info">
                  <div class="project-name">${proj.fullName}</div>
                  <div class="project-meta">序號：${proj.seqNo} | 承攬商：${proj.contractor}</div>
                </div>
                <span class="status-badge" style="background: var(--success);">已填寫</span>
              </div>
            `).join('')}
          </div>
        ` : ''}
      </div>
    `;

    container.appendChild(card);
  });
}

function toggleDeptCard(cardId) {
  const content = document.getElementById(`${cardId}_content`);
  const toggle = document.getElementById(`${cardId}_toggle`);

  if (content.style.display === 'none') {
    content.style.display = 'block';
    toggle.textContent = '▲';
  } else {
    content.style.display = 'none';
    toggle.textContent = '▼';
  }
}

// ============================================
// 修正4：日曆功能（月份處理修正）
// ============================================
function loadFilledDates() {
  window.GAS_API
    .withSuccessHandler(function (dates) {
      filledDates = dates;
    })
    .withFailureHandler(function (error) {
      console.error('載入已填寫日期失敗：', error);
    })
    .getFilledDates();
}

function changeMonth(delta) {
  currentCalendarMonth += delta;

  // 修正4：正確處理月份邊界
  if (currentCalendarMonth > 11) {
    currentCalendarMonth = 0;
    currentCalendarYear += 1;
  } else if (currentCalendarMonth < 0) {
    currentCalendarMonth = 11;
    currentCalendarYear -= 1;
  }

  renderCalendar();
}

function renderCalendar() {
  // 修正4：正確使用月份（JavaScript月份是0-11）
  const year = currentCalendarYear;
  const month = currentCalendarMonth + 1; // 轉為1-12供後端使用

  showLoading();

  // 載入該月假日資訊
  window.GAS_API
    .withSuccessHandler(function (holidays) {
      currentMonthHolidays = holidays;
      renderCalendarGrid();
      hideLoading();
    })
    .withFailureHandler(function (error) {
      console.error('載入假日資訊失敗：', error);
      currentMonthHolidays = {};
      renderCalendarGrid();
      hideLoading();
    })
    .getMonthHolidays(year, month);
}

function renderCalendarGrid() {
  const year = currentCalendarYear;
  const month = currentCalendarMonth;

  // 更新標題
  const monthNames = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
  document.getElementById('calendarTitle').textContent = `${year}年 ${monthNames[month]}`;

  const grid = document.getElementById('calendarGrid');
  grid.innerHTML = '';

  // 星期標題
  const weekdays = ['日', '一', '二', '三', '四', '五', '六'];
  weekdays.forEach(day => {
    const dayHeader = document.createElement('div');
    dayHeader.className = 'calendar-weekday';
    dayHeader.textContent = day;
    grid.appendChild(dayHeader);
  });

  // 計算該月第一天是星期幾
  const firstDay = new Date(year, month, 1).getDay();

  // 該月有多少天
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  // 填充空白日期
  for (let i = 0; i < firstDay; i++) {
    const emptyDay = document.createElement('div');
    emptyDay.className = 'calendar-day empty';
    grid.appendChild(emptyDay);
  }

  // 填充日期
  for (let day = 1; day <= daysInMonth; day++) {
    const dayDiv = document.createElement('div');
    dayDiv.className = 'calendar-day';

    const dateString = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

    // 檢查是否為假日
    const isHoliday = currentMonthHolidays[dateString];
    if (isHoliday) {
      dayDiv.classList.add('holiday');
    }

    // 檢查是否有填報資料並計算工程數量
    const filledData = filledDates.filter(fd => fd.dateString === dateString);
    const hasData = filledData.length > 0;
    const projectCount = hasData ? filledData.length : 0;

    if (hasData) {
      dayDiv.classList.add('has-data');
      dayDiv.onclick = function () {
        showCalendarDetailModal(dateString);
      };
    } else {
      // 沒有資料時也可以點擊建置日誌
      dayDiv.style.cursor = 'pointer';
      dayDiv.onclick = function () {
        createLogFromCalendar(dateString);
      };
    }

    dayDiv.innerHTML = `
      <div class="day-number">${day}</div>
      ${isHoliday && isHoliday.remark ? `<div class="day-remark">${isHoliday.remark}</div>` : ''}
      ${hasData ? `<div class="day-indicator" style="color: var(--success); font-weight: 700; font-size: 0.85rem;">${projectCount}筆</div>` : '<div class="day-indicator" style="color: var(--text-muted); font-size: 0.75rem;">+建置</div>'}
    `;

    grid.appendChild(dayDiv);
  }
}

// 從日曆建置日誌
function createLogFromCalendar(dateString) {
  const confirmMessage = `
    <p><strong>📅 建置日誌</strong></p>
    <p><strong>日期：</strong>${dateString}</p>
    <p style="margin-top: 1rem;">確認要切換到日誌填報頁面嗎？</p>
  `;

  showConfirmModal(confirmMessage, function () {
    // 設置日期
    document.getElementById('logDatePicker').value = dateString;

    // 切換到日誌填報頁面
    switchTab('logEntry');

    closeConfirmModal();

    // 滾動到頂部
    window.scrollTo({ top: 0, behavior: 'smooth' });

    showToast(`已切換至日誌填報，日期：${dateString}`);
  });
}

function showCalendarDetailModal(dateString) {
  document.getElementById('detailDateTitle').textContent = dateString;

  showLoading();
  window.GAS_API
    .withSuccessHandler(function (summaryData) {
      hideLoading();

      const filledData = summaryData.filter(row => row.hasFilled);

      const container = document.getElementById('calendarDetailContent');

      if (filledData.length === 0) {
        container.innerHTML = '<div class="text-muted">該日期無填報資料</div>';
      } else {
        container.innerHTML = '<div class="detail-list">' +
          filledData.map(row => `
            <div class="detail-item">
              <div class="detail-header">
                <h4>${row.fullName}</h4>
                <div class="detail-contractor">${row.contractor}</div>
              </div>
              <div class="detail-body">
                <div class="detail-row">
                  <span class="detail-label">部門：</span>
                  <span>${row.dept}</span>
                </div>
                <div class="detail-row">
                  <span class="detail-label">檢驗員：</span>
                  <span>${row.inspectors || '-'}</span>
                </div>
                <div class="detail-row">
                  <span class="detail-label">施工人數：</span>
                  <span>${row.isHolidayNoWork ? '假日不施工' : (row.workersCount + ' 人')}</span>
                </div>
                ${!row.isHolidayNoWork && row.workItems.length > 0 ? `
                  <div class="detail-row">
                    <span class="detail-label">工項：</span>
                    <span>${row.workItems.map(wi => wi.text).join('；')}</span>
                  </div>
                ` : ''}
              </div>
            </div>
          `).join('') +
          '</div>';
      }

      document.getElementById('calendarDetailModal').style.display = 'flex';
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('載入日期詳情失敗：' + error.message, true);
    })
    .getDailySummaryReport(dateString, 'all', 'all', isGuestMode, currentUserInfo);
}

function closeCalendarDetailModal() {
  document.getElementById('calendarDetailModal').style.display = 'none';
}

// ============================================
// TBM-KY 生成功能
// ============================================
function openTBMKYModal() {
  const dateString = document.getElementById('summaryDatePicker').value;

  if (!dateString) {
    showToast('請先選擇日期', true);
    return;
  }

  if (currentSummaryData.length === 0) {
    showToast('請先載入總表資料', true);
    return;
  }

  // 填充工程選單
  const select = document.getElementById('tbmkyProjectSelect');
  select.innerHTML = '<option value="">請選擇工程</option>';

  const filledProjects = currentSummaryData.filter(row => row.hasFilled && !row.isHolidayNoWork);

  if (filledProjects.length === 0) {
    showToast('當日無可生成TBM-KY的工程', true);
    return;
  }

  filledProjects.forEach(proj => {
    const option = document.createElement('option');
    option.value = proj.seqNo;
    option.textContent = `${proj.seqNo} - ${proj.fullName}`;
    select.appendChild(option);
  });

  document.getElementById('tbmkyDatePicker').value = dateString;

  document.getElementById('tbmkyModal').style.display = 'flex';
}

function closeTBMKYModal() {
  document.getElementById('tbmkyModal').style.display = 'none';
}

function confirmGenerateTBMKY() {
  const projectSeqNo = document.getElementById('tbmkyProjectSelect').value;
  const dateString = document.getElementById('tbmkyDatePicker').value;
  const mode = document.querySelector('input[name="tbmkyMode"]:checked').value;

  if (!projectSeqNo || !dateString) {
    showToast('請選擇工程和日期', true);
    return;
  }

  closeTBMKYModal();
  showLoading();

  window.GAS_API
    .withSuccessHandler(function (result) {
      hideLoading();
      if (result.success) {
        showTBMKYResultModal(result);
      } else {
        showToast('生成失敗：' + result.message, true);
      }
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('伺服器錯誤：' + error.message, true);
    })
    .generateTBMKY({
      projectSeqNo: projectSeqNo,
      dateString: dateString,
      mode: mode
    });
}

function showTBMKYResultModal(result) {
  document.getElementById('tbmkyResultMessage').textContent = result.message;

  const filesList = document.getElementById('tbmkyFilesList');
  filesList.innerHTML = '';

  result.files.forEach(file => {
    const fileItem = document.createElement('div');
    fileItem.className = 'tbmky-file-item';

    fileItem.innerHTML = `
      <div class="tbmky-file-icon">📄</div>
      <div class="tbmky-file-info">
        <div class="tbmky-file-name">${file.name}</div>
        <a href="${file.url}" target="_blank" class="tbmky-file-link">🔗 開啟文件</a>
      </div>
    `;

    filesList.appendChild(fileItem);
  });

  document.getElementById('tbmkyResultModal').style.display = 'flex';
}

function closeTBMKYResultModal() {
  document.getElementById('tbmkyResultModal').style.display = 'none';
}

// ============================================
// 工具函數
// ============================================
function showLoading() {
  document.getElementById('loadingOverlay').classList.add('active');
}

function hideLoading() {
  document.getElementById('loadingOverlay').classList.remove('active');
}

function showToast(message, isError = false) {
  const toast = document.getElementById('toast');
  toast.textContent = message;

  if (isError) {
    toast.classList.add('error');
  } else {
    toast.classList.remove('error');
  }

  toast.classList.add('show');

  setTimeout(() => {
    toast.classList.remove('show');
  }, 3000);
}

function showConfirmModal(message, onConfirm) {
  document.getElementById('confirmMessage').innerHTML = message;

  const confirmBtn = document.getElementById('confirmBtn');

  // 移除舊的監聽器
  const newConfirmBtn = confirmBtn.cloneNode(true);
  confirmBtn.parentNode.replaceChild(newConfirmBtn, confirmBtn);

  // 添加新的監聽器
  newConfirmBtn.addEventListener('click', function () {
    onConfirm();
  });

  document.getElementById('confirmModal').style.display = 'flex';
}

function closeConfirmModal() {
  document.getElementById('confirmModal').style.display = 'none';
}

function renderProjectSelect(selectId, projects, activeOnly = false) {
  const select = document.getElementById(selectId);
  if (!select) return;

  select.innerHTML = '<option value="">請選擇工程</option>';

  let filteredProjects = projects;

  // 篩選施工中的工程
  if (activeOnly) {
    filteredProjects = projects.filter(p => p.projectStatus === '施工中');
  }

  // 填表人只能看到被指派的工程
  if (currentUserInfo && currentUserInfo.role === '填表人') {
    const managedProjects = currentUserInfo.managedProjects || [];
    filteredProjects = filteredProjects.filter(p => {
      // 支援兩種格式：
      // 1. 純序號：'1', '2', '3'
      // 2. 序號+名稱：'1明潭-聯鋐', '2協和-泰興-士林'
      return managedProjects.some(managed => {
        // 提取序號（取第一個非數字字元之前的部分）
        const managedSeqNo = managed.match(/^\d+/)?.[0];
        return managedSeqNo && p.seqNo.toString() === managedSeqNo;
      });
    });
  }

  filteredProjects.forEach(project => {
    const option = document.createElement('option');
    option.value = project.seqNo;
    option.textContent = `${project.seqNo} - ${project.shortName}`;
    option.setAttribute('data-short-name', project.shortName);
    select.appendChild(option);
  });
}

function truncateText(text, maxLength) {
  if (!text) return '-';
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength) + '...';
}

function openLogEntryForProject(seqNo, projectName) {
  const confirmMessage = `
    <p><strong>📝 切換至日誌填報</strong></p>
    <p><strong>🏗️ 工程：</strong>${projectName}</p>
    <p>是否要切換到日誌填報頁面並選擇此工程？</p>
  `;

  showConfirmModal(confirmMessage, function () {
    switchTab('logEntry');
    document.getElementById('logProjectSelect').value = seqNo;
    handleProjectChange();
    closeConfirmModal();

    // 滾動到頂部
    window.scrollTo({ top: 0, behavior: 'smooth' });

    showToast('✓ 已切換至日誌填報，請填寫資料');
  });
}

// ============================================
// 鍵盤快捷鍵
// ============================================
document.addEventListener('keydown', function (e) {
  // ESC 關閉彈窗
  if (e.key === 'Escape') {
    // 依序檢查並關閉打開的彈窗
    if (document.getElementById('editSummaryLogModal').style.display === 'flex') {
      closeEditSummaryLogModal();
    } else if (document.getElementById('editProjectModal').style.display === 'flex') {
      closeEditProjectModal();
    } else if (document.getElementById('addInspectorModal').style.display === 'flex') {
      closeAddInspectorModal();
    } else if (document.getElementById('editInspectorModal').style.display === 'flex') {
      closeEditInspectorModal();
    } else if (document.getElementById('tbmkyModal').style.display === 'flex') {
      closeTBMKYModal();
    } else if (document.getElementById('tbmkyResultModal').style.display === 'flex') {
      closeTBMKYResultModal();
    } else if (document.getElementById('calendarDetailModal').style.display === 'flex') {
      closeCalendarDetailModal();
    } else if (document.getElementById('confirmModal').style.display === 'flex') {
      closeConfirmModal();
    }
  }
});

// ============================================
// 彈窗外點擊關閉
// ============================================
function setupModalOutsideClick() {
  const modals = [
    'editSummaryLogModal',
    'editProjectModal',
    'addInspectorModal',
    'editInspectorModal',
    'tbmkyModal',
    'tbmkyResultModal',
    'calendarDetailModal',
    'confirmModal'
  ];

  modals.forEach(modalId => {
    const modal = document.getElementById(modalId);
    if (modal) {
      modal.addEventListener('click', function (e) {
        if (e.target === modal) {
          switch (modalId) {
            case 'editSummaryLogModal':
              closeEditSummaryLogModal();
              break;
            case 'editProjectModal':
              closeEditProjectModal();
              break;
            case 'addInspectorModal':
              closeAddInspectorModal();
              break;
            case 'editInspectorModal':
              closeEditInspectorModal();
              break;
            case 'tbmkyModal':
              closeTBMKYModal();
              break;
            case 'tbmkyResultModal':
              closeTBMKYResultModal();
              break;
            case 'calendarDetailModal':
              closeCalendarDetailModal();
              break;
            case 'confirmModal':
              closeConfirmModal();
              break;
          }
        }
      });
    }
  });
}

// 當DOM載入完成後設置
document.addEventListener('DOMContentLoaded', function () {
  setupModalOutsideClick();
});

// ============================================
// 防止重複提交
// ============================================
let isSubmitting = false;

function preventDoubleSubmit(fn) {
  return function (...args) {
    if (isSubmitting) {
      showToast('請勿重複提交', true);
      return;
    }

    isSubmitting = true;

    try {
      fn.apply(this, args);
    } finally {
      setTimeout(() => {
        isSubmitting = false;
      }, 1000);
    }
  };
}

// ============================================
// 自動儲存草稿（可選功能）
// ============================================
function saveDraft() {
  try {
    const draft = {
      logDate: document.getElementById('logDatePicker').value,
      projectSeqNo: document.getElementById('logProjectSelect').value,
      workersCount: document.getElementById('logWorkersCount').value,
      timestamp: new Date().toISOString()
    };

    localStorage.setItem('dailyLogDraft', JSON.stringify(draft));
  } catch (error) {
    console.error('儲存草稿失敗：', error);
  }
}

function loadDraft() {
  try {
    const draftStr = localStorage.getItem('dailyLogDraft');
    if (!draftStr) return;

    const draft = JSON.parse(draftStr);

    // 檢查草稿是否在24小時內
    const draftTime = new Date(draft.timestamp);
    const now = new Date();
    const hoursDiff = (now - draftTime) / 1000 / 60 / 60;

    if (hoursDiff > 24) {
      localStorage.removeItem('dailyLogDraft');
      return;
    }

    // 詢問是否載入草稿
    if (confirm('發現未提交的草稿，是否要載入？')) {
      if (draft.logDate) document.getElementById('logDatePicker').value = draft.logDate;
      if (draft.projectSeqNo) document.getElementById('logProjectSelect').value = draft.projectSeqNo;
      if (draft.workersCount) document.getElementById('logWorkersCount').value = draft.workersCount;

      showToast('✓ 已載入草稿');
    } else {
      localStorage.removeItem('dailyLogDraft');
    }
  } catch (error) {
    console.error('載入草稿失敗：', error);
  }
}

// ============================================
// 頁面離開前提醒
// ============================================
window.addEventListener('beforeunload', function (e) {
  // 檢查表單是否有未提交的內容
  const form = document.getElementById('dailyLogForm');
  if (!form) return;

  const hasContent =
    document.getElementById('logProjectSelect').value ||
    document.getElementById('logWorkersCount').value ||
    document.querySelectorAll('.work-item-pair').length > 0;

  if (hasContent) {
    // 儲存草稿
    saveDraft();

    // 標準的離開提醒
    e.preventDefault();
    e.returnValue = '';
  }
});

// ============================================
// 效能監控（開發用）
// ============================================
function logPerformance(label) {
  if (window.performance && window.performance.now) {
    const time = window.performance.now();
    console.log(`[Performance] ${label}: ${time.toFixed(2)}ms`);
  }
}

// ============================================
// 錯誤追蹤
// ============================================
window.addEventListener('error', function (e) {
  console.error('Global error:', e.error);

  // 可選：將錯誤記錄到後端
  if (currentUserInfo && currentUserInfo.role === '超級管理員') {
    showToast(`系統錯誤：${e.message}`, true);
  }
});

window.addEventListener('unhandledrejection', function (e) {
  console.error('Unhandled promise rejection:', e.reason);

  if (currentUserInfo && currentUserInfo.role === '超級管理員') {
    showToast(`Promise錯誤：${e.reason}`, true);
  }
});

// ============================================
// 版本資訊
// ============================================
// 使用者管理功能
// ============================================

let allUsersData = [];

function loadUserManagement() {
  // 顯示/隱藏提示文字（只對聯絡員顯示）
  const hintElement = document.getElementById('userManagementHint');
  if (hintElement && currentUserInfo.role === '聯絡員') {
    hintElement.style.display = 'block';
  }

  // showLoading 已在 switchTab 中處理
  window.GAS_API
    .withSuccessHandler(function (users) {
      hideLoading();
      allUsersData = users;
      populateUserDeptFilter();
      applyUserFilters();
    })
    .withFailureHandler(function (error) {
      hideLoading();
      showToast('載入使用者資料失敗：' + error.message, true);
    })
    .getAllUsers();
}

function populateUserDeptFilter() {
  const deptSet = new Set();
  allUsersData.forEach(user => {
    if (user.dept && user.dept.trim()) {
      deptSet.add(user.dept);
    }
  });

  // 聯絡員權限過濾：只顯示自己部門 + 委外監造部門
  let deptArray = Array.from(deptSet).sort();
  if (currentUserInfo.role === '聯絡員') {
    deptArray = deptArray.filter(dept => {
      return dept === currentUserInfo.dept || dept.includes('委外監造');
    });
  }

  const deptFilter = document.getElementById('userDeptFilter');
  deptFilter.innerHTML = '<option value="all">全部部門</option>';

  deptArray.forEach(dept => {
    const option = document.createElement('option');
    option.value = dept;
    option.textContent = dept;
    deptFilter.appendChild(option);
  });
}

function applyUserFilters() {
  const roleFilter = document.getElementById('userRoleFilter').value;
  const deptFilter = document.getElementById('userDeptFilter').value;
  const searchText = document.getElementById('userSearchInput').value.toLowerCase().trim();

  // 根據當前使用者角色篩選可見的使用者
  let visibleUsers = allUsersData;
  if (currentUserInfo.role === '聯絡員') {
    // 聯絡員只能看到自己部門的填表人 + 委外監造部門的填表人
    visibleUsers = allUsersData.filter(u => {
      return u.role === '填表人' &&
        (u.dept === currentUserInfo.dept || (u.dept && u.dept.includes('委外監造')));
    });
  }

  // 應用角色篩選
  if (roleFilter !== 'all') {
    visibleUsers = visibleUsers.filter(u => u.role === roleFilter);
  }

  // 應用部門篩選
  if (deptFilter !== 'all') {
    visibleUsers = visibleUsers.filter(u => u.dept === deptFilter);
  }

  // 應用搜尋
  if (searchText) {
    visibleUsers = visibleUsers.filter(u => {
      return (u.name && u.name.toLowerCase().includes(searchText)) ||
        (u.account && u.account.toLowerCase().includes(searchText)) ||
        (u.email && u.email.toLowerCase().includes(searchText));
    });
  }

  renderUserCards(visibleUsers);
  updateUserStats(visibleUsers.length);
}

function clearUserFilters() {
  document.getElementById('userRoleFilter').value = 'all';
  document.getElementById('userDeptFilter').value = 'all';
  document.getElementById('userSearchInput').value = '';
  applyUserFilters();
}

function updateUserStats(count) {
  document.getElementById('userTotalCount').textContent = count;
}

function renderUserCards(visibleUsers) {
  const container = document.getElementById('userListContainer');

  if (!visibleUsers || visibleUsers.length === 0) {
    container.innerHTML = '<div style="text-align: center; padding: 3rem; color: #9ca3af;"><p style="font-size: 1.25rem; margin-bottom: 0.5rem;">📭</p><p>找不到符合條件的使用者</p></div>';
    return;
  }

  container.innerHTML = visibleUsers.map(user => {
    const managedProjectsText = user.managedProjects.length > 0 ?
      user.managedProjects.join('、') : '<span style="color: #94a3b8;">無</span>';

    // 根據角色設定顏色和樣式
    let roleConfig = {
      color: '#2563eb',
      bgLight: '#eff6ff',
      bgGradient: 'linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%)',
      icon: '✍️',
      label: '填表人'
    };

    if (user.role === '超級管理員') {
      roleConfig = {
        color: '#dc2626',
        bgLight: '#fef2f2',
        bgGradient: 'linear-gradient(135deg, #fee2e2 0%, #fecaca 100%)',
        icon: '👑',
        label: '超級管理員'
      };
    } else if (user.role === '聯絡員') {
      roleConfig = {
        color: '#059669',
        bgLight: '#f0fdf4',
        bgGradient: 'linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%)',
        icon: '📞',
        label: '聯絡員'
      };
    }

    return `
      <div style="
        background: white;
        border-radius: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1), 0 4px 12px rgba(0,0,0,0.05);
        overflow: hidden;
        transition: all 0.3s ease;
        border: 2px solid ${roleConfig.bgLight};
        height: 100%;
        display: flex;
        flex-direction: column;
      " onmouseover="this.style.boxShadow='0 8px 24px rgba(0,0,0,0.12)'; this.style.transform='translateY(-4px)'" onmouseout="this.style.boxShadow='0 1px 3px rgba(0,0,0,0.1), 0 4px 12px rgba(0,0,0,0.05)'; this.style.transform='translateY(0)'">

        <div style="background: ${roleConfig.bgGradient}; padding: 1rem 1.5rem; border-bottom: 3px solid ${roleConfig.color};">
          <div style="display: flex; align-items: center; gap: 0.75rem;">
            <div style="
              width: 48px;
              height: 48px;
              background: ${roleConfig.color};
              border-radius: 12px;
              display: flex;
              align-items: center;
              justify-content: center;
              font-size: 1.5rem;
              box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            ">${roleConfig.icon}</div>
            <div style="flex: 1;">
              <div style="font-size: 1.25rem; font-weight: 700; color: #1e293b; margin-bottom: 0.25rem;">
                ${user.name}
              </div>
              <div style="
                display: inline-block;
                background: ${roleConfig.color};
                color: white;
                padding: 0.25rem 0.75rem;
                border-radius: 6px;
                font-size: 0.8rem;
                font-weight: 600;
                letter-spacing: 0.02em;
              ">${roleConfig.label}</div>
            </div>
          </div>
        </div>

        <div style="padding: 1.5rem; flex: 1; background: #fafafa;">
          <div style="display: flex; flex-direction: column; gap: 1rem;">

            ${user.dept ? `
            <div style="display: flex; align-items: center; gap: 0.75rem; padding: 0.75rem; background: white; border-radius: 8px; border-left: 3px solid ${roleConfig.color};">
              <span style="font-size: 1.25rem;">🏢</span>
              <div style="flex: 1;">
                <div style="font-size: 0.75rem; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.125rem;">部門</div>
                <div style="color: #1e293b; font-weight: 600;">${user.dept}</div>
              </div>
            </div>
            ` : ''}

            <div style="display: flex; align-items: center; gap: 0.75rem; padding: 0.75rem; background: white; border-radius: 8px;">
              <span style="font-size: 1.25rem;">🔑</span>
              <div style="flex: 1;">
                <div style="font-size: 0.75rem; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.125rem;">帳號</div>
                <div style="color: #1e293b; font-family: 'Courier New', monospace; font-weight: 600; font-size: 0.95rem;">${user.account}</div>
              </div>
            </div>

            <div style="display: flex; align-items: center; gap: 0.75rem; padding: 0.75rem; background: white; border-radius: 8px;">
              <span style="font-size: 1.25rem;">📧</span>
              <div style="flex: 1; min-width: 0;">
                <div style="font-size: 0.75rem; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.125rem;">Email</div>
                <div style="color: #1e293b; font-size: 0.9rem; word-break: break-all;">${user.email}</div>
              </div>
            </div>

            ${user.managedProjects.length > 0 ? `
            <div style="display: flex; align-items: start; gap: 0.75rem; padding: 0.75rem; background: white; border-radius: 8px;">
              <span style="font-size: 1.25rem;">🏗️</span>
              <div style="flex: 1;">
                <div style="font-size: 0.75rem; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.25rem;">管理工程</div>
                <div style="color: #1e293b; line-height: 1.6;">${managedProjectsText}</div>
              </div>
            </div>
            ` : ''}

            ${user.supervisorEmail ? `
            <div style="display: flex; align-items: center; gap: 0.75rem; padding: 0.75rem; background: white; border-radius: 8px;">
              <span style="font-size: 1.25rem;">👔</span>
              <div style="flex: 1; min-width: 0;">
                <div style="font-size: 0.75rem; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.125rem;">主管 Email</div>
                <div style="color: #1e293b; font-size: 0.9rem; word-break: break-all;">${user.supervisorEmail}</div>
              </div>
            </div>
            ` : ''}
          </div>
        </div>

        <div style="padding: 1rem 1.5rem; background: white; border-top: 2px solid #f1f5f9; display: flex; gap: 0.75rem;">
          <button
            onclick="editUser(${user.rowIndex})"
            style="
              flex: 1;
              padding: 0.75rem 1.5rem;
              background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
              color: white;
              border: none;
              border-radius: 10px;
              font-weight: 600;
              font-size: 0.95rem;
              cursor: pointer;
              transition: all 0.2s;
              box-shadow: 0 2px 8px rgba(59, 130, 246, 0.3);
            "
            onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 4px 12px rgba(59, 130, 246, 0.4)'"
            onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 8px rgba(59, 130, 246, 0.3)'"
          >
            ✏️ 編輯
          </button>
          <button
            onclick="deleteUserConfirm(${user.rowIndex}, '${user.name}')"
            style="
              flex: 1;
              padding: 0.75rem 1.5rem;
              background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
              color: white;
              border: none;
              border-radius: 10px;
              font-weight: 600;
              font-size: 0.95rem;
              cursor: pointer;
              transition: all 0.2s;
              box-shadow: 0 2px 8px rgba(239, 68, 68, 0.3);
            "
            onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 4px 12px rgba(239, 68, 68, 0.4)'"
            onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 8px rgba(239, 68, 68, 0.3)'"
          >
            🗑️ 刪除
          </button>
        </div>
      </div>
    `;
  }).join('');
}

function openAddUserModal() {
  document.getElementById('addUserModalTitle').textContent = '新增使用者';
  document.getElementById('editUserRowIndex').value = '';
  document.getElementById('userDept').value = '';
  document.getElementById('userName').value = '';
  document.getElementById('userAccount').value = '';
  document.getElementById('userEmail').value = '';
  document.getElementById('userRole').value = '填表人';
  document.getElementById('userPassword').value = '';
  document.getElementById('userSupervisorEmail').value = '';

  document.getElementById('passwordRequired').style.display = '';
  document.getElementById('passwordHint').style.display = 'none';

  // 根據當前使用者角色設定可選身分
  const roleSelect = document.getElementById('userRole');
  if (currentUserInfo.role === '聯絡員') {
    roleSelect.innerHTML = '<option value="填表人">填表人</option>';
    roleSelect.value = '填表人';
    roleSelect.disabled = true;
  } else {
    roleSelect.innerHTML = `
      <option value="填表人">填表人</option>
      <option value="聯絡員">聯絡員</option>
      <option value="超級管理員">超級管理員</option>
    `;
    roleSelect.disabled = false;
  }

  // 載入部門列表
  loadDepartmentsForUser();

  // 載入工程列表
  loadProjectsForUser();

  handleRoleChange();
  document.getElementById('addUserModal').style.display = 'flex';
}

function loadDepartmentsForUser() {
  window.GAS_API
    .withSuccessHandler(function (departments) {
      const deptSelect = document.getElementById('userDept');
      deptSelect.innerHTML = '<option value="">請選擇部門</option>';

      // 聯絡員權限過濾
      let filteredDepts = departments;
      if (currentUserInfo.role === '聯絡員') {
        filteredDepts = departments.filter(dept => {
          // 聯絡員可以看到自己的部門 + 包含「委外監造」的部門
          return dept === currentUserInfo.dept || dept.includes('委外監造');
        });
      }

      // 預設選擇聯絡員自己的部門
      const defaultDept = currentUserInfo.role === '聯絡員' ? currentUserInfo.dept : '';

      filteredDepts.forEach(dept => {
        const option = document.createElement('option');
        option.value = dept;
        option.textContent = dept;
        deptSelect.appendChild(option);
      });

      // 設定預設值
      if (defaultDept) {
        deptSelect.value = defaultDept;
      }
    })
    .withFailureHandler(function (error) {
      showToast('載入部門列表失敗：' + error.message, true);
    })
    .getDepartmentsList();
}

function closeAddUserModal() {
  document.getElementById('addUserModal').style.display = 'none';
}

function handleRoleChange() {
  const role = document.getElementById('userRole').value;
  const managedProjectsGroup = document.getElementById('managedProjectsGroup');

  // 只有填表人需要選擇管理工程
  if (role === '填表人') {
    managedProjectsGroup.style.display = 'block';
  } else {
    managedProjectsGroup.style.display = 'none';
  }
}

function loadProjectsForUser() {
  window.GAS_API
    .withSuccessHandler(function (projects) {
      const container = document.getElementById('managedProjectsCheckboxes');
      container.innerHTML = '';

      // 按主辦部門分組
      const projectsByDept = {};
      projects.forEach(project => {
        const dept = project.dept || '未分類';
        if (!projectsByDept[dept]) {
          projectsByDept[dept] = [];
        }
        projectsByDept[dept].push(project);
      });

      // 渲染分組
      Object.keys(projectsByDept).sort().forEach(dept => {
        const deptDiv = document.createElement('div');
        deptDiv.style.cssText = 'background: white; padding: 1rem; border-radius: 8px; border: 1px solid #e5e7eb;';

        const deptHeader = document.createElement('div');
        deptHeader.style.cssText = 'display: flex; align-items: center; justify-content: space-between; margin-bottom: 0.75rem; padding-bottom: 0.75rem; border-bottom: 2px solid #e5e7eb;';
        deptHeader.innerHTML = `
          <strong style="font-size: 1rem; color: #1f2937; display: flex; align-items: center; gap: 0.5rem;">
            <span style="font-size: 1.25rem;">🏢</span>
            ${dept} <span style="color: #6b7280; font-weight: normal; font-size: 0.875rem;">(${projectsByDept[dept].length})</span>
          </strong>
          <button
            type="button"
            class="btn-select-all"
            data-dept="${dept}"
            onclick="toggleSelectAllInDept('${dept}')"
            style="padding: 0.375rem 0.75rem; background: #3b82f6; color: white; border: none; border-radius: 6px; font-size: 0.875rem; font-weight: 600; cursor: pointer; transition: all 0.2s;"
            onmouseover="this.style.background='#2563eb'"
            onmouseout="this.style.background='#3b82f6'"
          >
            全選
          </button>
        `;

        const deptProjects = document.createElement('div');
        deptProjects.id = `dept_${dept.replace(/\s+/g, '_')}`;
        deptProjects.style.cssText = 'display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 0.5rem;';

        projectsByDept[dept].forEach(project => {
          const checkboxDiv = document.createElement('div');
          checkboxDiv.style.cssText = 'display: flex; align-items: center; gap: 0.5rem; padding: 0.5rem; border-radius: 6px; transition: background 0.2s;';
          checkboxDiv.onmouseover = function () { this.style.background = '#f3f4f6'; };
          checkboxDiv.onmouseout = function () { this.style.background = 'transparent'; };

          checkboxDiv.innerHTML = `
            <input
              type="checkbox"
              id="project_${project.seqNo}"
              value="${project.seqNo}"
              data-dept="${dept}"
              style="width: 18px; height: 18px; cursor: pointer;"
            >
            <label for="project_${project.seqNo}" style="cursor: pointer; font-size: 0.9rem; color: #374151; flex: 1;">
              <strong>${project.seqNo}</strong> - ${project.shortName}
            </label>
          `;
          deptProjects.appendChild(checkboxDiv);
        });

        deptDiv.appendChild(deptHeader);
        deptDiv.appendChild(deptProjects);
        container.appendChild(deptDiv);
      });
    })
    .withFailureHandler(function (error) {
      showToast('載入工程列表失敗：' + error.message, true);
    })
    .getAllProjects();
}

function toggleSelectAllInDept(dept) {
  const checkboxes = document.querySelectorAll(`#managedProjectsCheckboxes input[data-dept="${dept}"]`);
  const allChecked = Array.from(checkboxes).every(cb => cb.checked);

  // 切換全選/取消全選
  checkboxes.forEach(cb => {
    cb.checked = !allChecked;
  });

  // 更新按鈕文字
  const btn = document.querySelector(`.btn-select-all[data-dept="${dept}"]`);
  if (btn) {
    btn.textContent = allChecked ? '全選' : '取消全選';
  }
}

function editUser(rowIndex) {
  const user = allUsersData.find(u => u.rowIndex === rowIndex);
  if (!user) {
    showToast('找不到使用者資料', true);
    return;
  }

  document.getElementById('addUserModalTitle').textContent = '編輯使用者';
  document.getElementById('editUserRowIndex').value = user.rowIndex;
  document.getElementById('userDept').value = user.dept;
  document.getElementById('userName').value = user.name;
  document.getElementById('userAccount').value = user.account;
  document.getElementById('userEmail').value = user.email;
  document.getElementById('userRole').value = user.role;
  document.getElementById('userPassword').value = '';
  document.getElementById('userSupervisorEmail').value = user.supervisorEmail;

  document.getElementById('passwordRequired').style.display = 'none';
  document.getElementById('passwordHint').style.display = 'block';

  // 根據當前使用者角色設定可選身分
  const roleSelect = document.getElementById('userRole');
  if (currentUserInfo.role === '聯絡員') {
    roleSelect.innerHTML = '<option value="填表人">填表人</option>';
    roleSelect.disabled = true;
  } else {
    roleSelect.innerHTML = `
      <option value="填表人">填表人</option>
      <option value="聯絡員">聯絡員</option>
      <option value="超級管理員">超級管理員</option>
    `;
    roleSelect.disabled = false;
  }

  // 載入工程列表並勾選已選工程
  window.GAS_API
    .withSuccessHandler(function (projects) {
      const container = document.getElementById('managedProjectsCheckboxes');
      container.innerHTML = '';

      // 按主辦部門分組 (編輯模式下也需分組渲染以保持一致性)
      const projectsByDept = {};
      projects.forEach(project => {
        const dept = project.dept || '未分類';
        if (!projectsByDept[dept]) {
          projectsByDept[dept] = [];
        }
        projectsByDept[dept].push(project);
      });

      Object.keys(projectsByDept).sort().forEach(dept => {
        const deptDiv = document.createElement('div');
        deptDiv.style.cssText = 'background: white; padding: 1rem; border-radius: 8px; border: 1px solid #e5e7eb; margin-bottom: 0.5rem;';

        const deptHeader = document.createElement('div');
        deptHeader.style.cssText = 'display: flex; align-items: center; justify-content: space-between; margin-bottom: 0.75rem; padding-bottom: 0.75rem; border-bottom: 2px solid #e5e7eb;';
        deptHeader.innerHTML = `
          <strong style="font-size: 1rem; color: #1f2937;">${dept}</strong>
          <button type="button" class="btn-select-all" data-dept="${dept}" onclick="toggleSelectAllInDept('${dept}')" style="padding: 0.25rem 0.5rem; font-size: 0.8rem; background: #eee; border: none; border-radius: 4px; cursor: pointer;">全選</button>
        `;

        const deptProjects = document.createElement('div');
        deptProjects.style.cssText = 'display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 0.5rem;';

        projectsByDept[dept].forEach(project => {
          const isChecked = user.managedProjects.includes(project.seqNo);
          const checkboxDiv = document.createElement('div');
          checkboxDiv.style.cssText = 'display: flex; align-items: center; gap: 0.5rem; padding: 0.5rem; border-radius: 6px;';

          checkboxDiv.innerHTML = `
            <input type="checkbox" id="project_${project.seqNo}" value="${project.seqNo}" data-dept="${dept}" ${isChecked ? 'checked' : ''} style="cursor: pointer;">
            <label for="project_${project.seqNo}" style="cursor: pointer; font-size: 0.9rem;">${project.seqNo} - ${project.shortName}</label>
          `;
          deptProjects.appendChild(checkboxDiv);
        });

        deptDiv.appendChild(deptHeader);
        deptDiv.appendChild(deptProjects);
        container.appendChild(deptDiv);
      });
    })
    .getAllProjects();

  handleRoleChange();
  document.getElementById('addUserModal').style.display = 'flex';
}

function confirmAddUser() {
  const rowIndex = document.getElementById('editUserRowIndex').value;
  const dept = document.getElementById('userDept').value.trim();
  const name = document.getElementById('userName').value.trim();
  const account = document.getElementById('userAccount').value.trim();
  const email = document.getElementById('userEmail').value.trim();
  const role = document.getElementById('userRole').value;
  const password = document.getElementById('userPassword').value;
  const supervisorEmail = document.getElementById('userSupervisorEmail').value.trim();

  // 驗證必填欄位
  if (!name || !account || !email) {
    showToast('請填寫所有必填欄位', true);
    return;
  }

  // 新增時密碼為必填
  if (!rowIndex && !password) {
    showToast('請輸入密碼', true);
    return;
  }

  // 取得選中的工程
  let managedProjects = [];
  if (role === '填表人') {
    const checkboxes = document.querySelectorAll('#managedProjectsCheckboxes input[type="checkbox"]:checked');
    managedProjects = Array.from(checkboxes).map(cb => cb.value);

    if (managedProjects.length === 0) {
      showToast('填表人至少需要選擇一個可管理的工程', true);
      return;
    }
  }

  const userData = {
    dept,
    name,
    account,
    email,
    role,
    password,
    managedProjects,
    supervisorEmail
  };

  if (rowIndex) {
    // 編輯：顯示確認視窗
    userData.rowIndex = parseInt(rowIndex);

    // 找到原始使用者資料
    const originalUser = allUsersData.find(u => u.rowIndex === userData.rowIndex);

    showEditUserConfirmModal(originalUser, userData);
  } else {
    // 新增
    showLoading();
    window.GAS_API
      .withSuccessHandler(function (result) {
        hideLoading();
        if (result.success) {
          showToast('✓ ' + result.message);
          closeAddUserModal();
          loadUserManagement();
        } else {
          showToast('新增失敗：' + result.message, true);
        }
      })
      .withFailureHandler(function (error) {
        hideLoading();
        showToast('伺服器錯誤：' + error.message, true);
      })
      .addUser(userData);
  }
}

function deleteUserConfirm(rowIndex, userName) {
  const confirmMessage = `
    <p><strong>🗑️ 刪除使用者</strong></p>
    <p><strong>使用者：</strong>${userName}</p>
    <p style="margin-top: 1rem; color: var(--danger);">⚠️ 確認刪除嗎？此操作無法復原！</p>
  `;

  showConfirmModal(confirmMessage, function () {
    showLoading();
    window.GAS_API
      .withSuccessHandler(function (result) {
        hideLoading();
        if (result.success) {
          showToast('✓ ' + result.message);
          loadUserManagement();
        } else {
          showToast('刪除失敗：' + result.message, true);
        }
      })
      .withFailureHandler(function (error) {
        hideLoading();
        showToast('伺服器錯誤：' + error.message, true);
      })
      .deleteUser(rowIndex);
  });
}

function showEditUserConfirmModal(originalUser, newUserData) {
  // 建立對比 HTML
  const changes = [];

  // 比對部門
  if (originalUser.dept !== newUserData.dept) {
    changes.push(`<tr>
      <td>部門</td>
      <td style="color: #9ca3af;">${originalUser.dept || '（空白）'}</td>
      <td style="color: var(--primary); font-weight: 600;">${newUserData.dept || '（空白）'}</td>
    </tr>`);
  } else {
    changes.push(`<tr>
      <td>部門</td>
      <td colspan="2">${newUserData.dept}</td>
    </tr>`);
  }

  // 比對姓名
  if (originalUser.name !== newUserData.name) {
    changes.push(`<tr>
      <td>姓名</td>
      <td style="color: #9ca3af;">${originalUser.name}</td>
      <td style="color: var(--primary); font-weight: 600;">${newUserData.name}</td>
    </tr>`);
  } else {
    changes.push(`<tr>
      <td>姓名</td>
      <td colspan="2">${newUserData.name}</td>
    </tr>`);
  }

  // 比對帳號
  if (originalUser.account !== newUserData.account) {
    changes.push(`<tr>
      <td>帳號</td>
      <td style="color: #9ca3af;">${originalUser.account}</td>
      <td style="color: var(--primary); font-weight: 600;">${newUserData.account}</td>
    </tr>`);
  } else {
    changes.push(`<tr>
      <td>帳號</td>
      <td colspan="2">${newUserData.account}</td>
    </tr>`);
  }

  // 比對信箱
  if (originalUser.email !== newUserData.email) {
    changes.push(`<tr>
      <td>信箱</td>
      <td style="color: #9ca3af;">${originalUser.email}</td>
      <td style="color: var(--primary); font-weight: 600;">${newUserData.email}</td>
    </tr>`);
  } else {
    changes.push(`<tr>
      <td>信箱</td>
      <td colspan="2">${newUserData.email}</td>
    </tr>`);
  }

  // 比對身分
  if (originalUser.role !== newUserData.role) {
    changes.push(`<tr>
      <td>身分</td>
      <td style="color: #9ca3af;">${originalUser.role}</td>
      <td style="color: var(--primary); font-weight: 600;">${newUserData.role}</td>
    </tr>`);
  } else {
    changes.push(`<tr>
      <td>身分</td>
      <td colspan="2">${newUserData.role}</td>
    </tr>`);
  }

  // 比對管理工程（僅填表人）
  if (newUserData.role === '填表人') {
    const originalProjects = originalUser.managedProjects ? originalUser.managedProjects.join(', ') : '';
    const newProjects = newUserData.managedProjects ? newUserData.managedProjects.join(', ') : '';

    if (originalProjects !== newProjects) {
      changes.push(`<tr>
        <td>管理工程</td>
        <td style="color: #9ca3af;">${originalProjects || '（無）'}</td>
        <td style="color: var(--primary); font-weight: 600;">${newProjects || '（無）'}</td>
      </tr>`);
    } else {
      changes.push(`<tr>
        <td>管理工程</td>
        <td colspan="2">${newProjects || '（無）'}</td>
      </tr>`);
    }
  }

  const confirmMessage = `
    <div style="max-width: 600px;">
      <h3 style="margin-top: 0; color: #1f2937; display: flex; align-items: center; gap: 0.5rem;">
        <span style="font-size: 1.5rem;">🔍</span>
        <span>確認修改使用者資料</span>
      </h3>
      <p style="color: #6b7280; margin-bottom: 1.5rem;">
        請確認以下資料是否正確：
      </p>
      <table style="width: 100%; border-collapse: collapse; margin-bottom: 1.5rem;">
        <thead>
          <tr style="background: #f3f4f6;">
            <th style="padding: 0.75rem; text-align: left; border-bottom: 2px solid #e5e7eb; width: 120px;">欄位</th>
            <th style="padding: 0.75rem; text-align: left; border-bottom: 2px solid #e5e7eb;">修改前</th>
            <th style="padding: 0.75rem; text-align: left; border-bottom: 2px solid #e5e7eb;">修改後</th>
          </tr>
        </thead>
        <tbody>
          ${changes.join('')}
        </tbody>
      </table>
      <p style="color: #6b7280; font-size: 0.875rem; background: #fef3c7; padding: 1rem; border-radius: 8px; margin: 0;">
        💡 提示：修改後的內容會以<span style="color: var(--primary); font-weight: 600;">藍色</span>顯示
      </p>
    </div>
  `;

  showConfirmModal(confirmMessage, function () {
    // 確認後執行更新
    showLoading();
    window.GAS_API
      .withSuccessHandler(function (result) {
        hideLoading();
        closeConfirmModal();
        if (result.success) {
          showToast('✓ ' + result.message);
          closeAddUserModal();
          loadUserManagement();
        } else {
          showToast('更新失敗：' + result.message, true);
        }
      })
      .withFailureHandler(function (error) {
        hideLoading();
        closeConfirmModal();
        showToast('伺服器錯誤：' + error.message, true);
      })
      .updateUser(newUserData);
  });
}

// ============================================
console.log('%c綜合施工處 每日工程日誌系統 v2.1', 'color: #2563eb; font-size: 16px; font-weight: bold;');
console.log('%c修正日期：2025-01-18', 'color: #64748b; font-size: 12px;');
console.log('%c所有12項修正已完成實作', 'color: #10b981; font-size: 12px;');

// ============================================
// 初始化完成
// ============================================
console.log('%c✓ JavaScript 載入完成', 'color: #10b981; font-weight: bold;');
