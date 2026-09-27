/**
 * 企业作息与评价查询平台 - 核心业务逻辑 (App Core Logic)
 * 适配 GitHub Pages 静态站点架构：
 * 1. 自动异步拉取仓库中的 ./data.xlsx 并渲染（修改表格后推送即可全网自动更新）
 * 2. 访客录入即时本地预览 + 一键打通 GitHub Issue 申请永久收录
 */

(function () {
  'use strict';

  // ==========================================
  // 全局配置 (可根据你的实际 GitHub 仓库修改)
  // ==========================================
  const CONFIG = {
    // 自动加载的 Excel 文件相对路径 (推送到 GitHub 仓库即可自动拉取)
    excelUrl: './data.xlsx',
    // 你的 GitHub 仓库地址 (访客点击“提交收录”时将自动跳转到此处创建 Issue)
    // 部署到你的 GitHub Pages 后，请将此处改为你的真实仓库地址，例如：'https://github.com/myname/shuangxiu'
    githubRepoUrl: 'https://github.com/qx2807/shuangxiu'
  };

  // 默认离线后备演示数据
  const DEFAULT_SAMPLE_DATA = [
    {
      id: "comp_1",
      name: "微软中国 (Microsoft)",
      scheduleType: "严格双休",
      workingHours: "09:30 - 18:00 (弹性制)",
      industry: "外企 / 软件科技",
      reviews: [
        "真正的WLB天花板，严格双休，支持每周2-3天远程办公(WFH)。",
        "年假15天起步，带薪病假充足，下班后Teams绝无催命消息。",
        "同事之间人际关系简单，没有无效内卷和强制加班表演。"
      ]
    },
    {
      id: "comp_2",
      name: "腾讯科技 (TEG技术工程事业群)",
      scheduleType: "弹性双休",
      workingHours: "10:00 - 20:30",
      industry: "互联网大厂",
      reviews: [
        "默认周末双休，除非大促或线上事故极少周六加班。",
        "平时工作日晚上有加班餐和打车报销，一般在8点半至9点左右陆续离开。",
        "福利很好，房补、企鹅币、中秋礼盒给力，部门氛围看组。"
      ]
    },
    {
      id: "comp_3",
      name: "某大型国有银行软件研发中心",
      scheduleType: "严格双休",
      workingHours: "08:30 - 17:30",
      industry: "国企 / 金融科技",
      reviews: [
        "严格遵守国家法定节假日，周六日雷打不动双休。",
        "中午有两个小时吃饭休息时间，食堂物美价廉，有下午茶点心。",
        "偶尔投产上线周五晚上需要值班支持，但次周周一可以申请调休。"
      ]
    },
    {
      id: "comp_4",
      name: "米哈游 (miHoYo)",
      scheduleType: "弹性双休",
      workingHours: "10:00 - 19:30",
      industry: "游戏 / 文娱",
      reviews: [
        "不搞大小周，周末双休，新版本前可能会有少量版本加班但给调休。",
        "公司福利极好，包三餐，下午茶点心咖啡不断，年终奖丰厚。",
        "团队年轻有活力，二次元浓度高，做自己喜欢的游戏有成就感。"
      ]
    },
    {
      id: "comp_5",
      name: "某知名电商大厂 (核心业务部)",
      scheduleType: "单休 / 11-11-6",
      workingHours: "11:00 - 23:00",
      industry: "电商 / 互联网",
      reviews: [
        "业务节奏极快，基本单休，周日通常只能躺平补觉。",
        "薪资期权给得确实有竞争力，但对身心耐力要求极高。",
        "厕所排队严重，离职率偏高，适合拼1-2年攒经验和钱的同学。"
      ]
    },
    {
      id: "comp_6",
      name: "大疆创新 (DJI)",
      scheduleType: "大小周",
      workingHours: "09:00 - 21:00",
      industry: "智能硬件 / 机器人",
      reviews: [
        "实行大小周制度，单休那一周周六需全天出勤，有相应加班补偿。",
        "研发技术底蕴深厚，硬核工程能力强，技术人员成长飞速。",
        "考核严格，KPI压力不小，追求卓越的氛围很浓。"
      ]
    },
    {
      id: "comp_7",
      name: "SAP中国研究院",
      scheduleType: "严格双休",
      workingHours: "09:00 - 17:30",
      industry: "外企 / 企服",
      reviews: [
        "典型欧洲外企风格，不打卡，准点下班，周末完全是个人生活。",
        "孕妇产假、陪产假非常规范，对家庭生活极为友好。",
        "节奏较慢，技术栈以企业级稳定为主，适合长期沉淀养老。"
      ]
    },
    {
      id: "comp_8",
      name: "某头部造车新势力",
      scheduleType: "大小周",
      workingHours: "09:30 - 21:30",
      industry: "新能源汽车 / 制造",
      reviews: [
        "大小周，新车型交付阶段周六全员加班是常态。",
        "扁平化沟通，业务发展极快，能接触到整车与智能驾驶最前沿技术。",
        "节奏紧凑，会议较多，需要较强的心理抗压和跨部门协同能力。"
      ]
    },
    {
      id: "comp_9",
      name: "网易互娱 (雷火/盘古等)",
      scheduleType: "弹性双休",
      workingHours: "10:00 - 19:30",
      industry: "互联网 / 游戏",
      reviews: [
        "正常情况下周末双休，上下班不打卡，弹性工时比较自由。",
        "网易食堂全国知名，每日四餐免费，健身房和园区环境非常优美。",
        "节点上线期有冲刺加班，但平时只要工作完成即可自行支配时间。"
      ]
    },
    {
      id: "comp_10",
      name: "某跨境电商中小型出海公司",
      scheduleType: "单休",
      workingHours: "09:00 - 18:30 (旺季常态加班)",
      industry: "跨境贸易 / 电商",
      reviews: [
        "固定单休，周六必须上班，黑五和圣诞大促期间全天待命。",
        "底薪一般主要靠提成，爆款出来后提成确实可观。",
        "制度不够健全，人员流动大，没有加班费只有象征性零食补贴。"
      ]
    },
    {
      id: "comp_11",
      name: "亚马逊中国 (Amazon)",
      scheduleType: "严格双休",
      workingHours: "09:30 - 18:30",
      industry: "外企 / 电商云计算",
      reviews: [
        "严格双休，实行On-call轮值机制，轮到当值周会有专门调休。",
        "强调文档文化(6-pagers)，开会效率高，不提倡加班。",
        "股票期权机制透明，英语工作环境，全球转岗机会多。"
      ]
    },
    {
      id: "comp_12",
      name: "某省属设计研究院",
      scheduleType: "严格双休",
      workingHours: "08:30 - 17:00",
      industry: "国企 / 建筑市政",
      reviews: [
        "正常项目周期下严格双休，节假日按国家标准一律放假。",
        "工作环境稳定，有独立院区，食堂一日三餐象征性收费。",
        "赶投标节点偶尔通宵一两天，但事后有大段调休补偿。"
      ]
    },
    {
      id: "comp_13",
      name: "苏州智绿科技有限公司",
      scheduleType: "单休",
      workingHours: "08:30 - 17:00",
      industry: "汽车零部件/制造",
      reviews: [
        "外包和劳务派遣极多，随产线需求即招即走",
        "中午半小时吃饭时间，饭菜极差"
      ]
    },
    {
      id: "comp_14",
      name: "苏州ab科技有限公司",
      scheduleType: "单休",
      workingHours: "08:00 - 20:00",
      industry: "示例",
      reviews: [
        "示例1",
        "示例2"
      ]
    }
  ];

  const STORAGE_KEY = 'company_schedule_data_v1';

  // 应用状态
  const state = {
    companies: [],
    searchQuery: '',
    currentFilter: 'all',
    currentSort: 'default',
    currentView: 'card', // 'card' | 'table'
    selectedCompanyId: null,
    lastAddedCompany: null
  };

  // DOM 元素缓存
  const elements = {
    // 自动同步状态条
    syncDot: document.getElementById('syncDot'),
    syncStatusText: document.getElementById('syncStatusText'),
    btnReloadRemote: document.getElementById('btnReloadRemote'),
    btnToggleDropzone: document.getElementById('btnToggleDropzone'),

    // 统计
    statTotalCompanies: document.getElementById('statTotalCompanies'),
    statShuangxiuCount: document.getElementById('statShuangxiuCount'),
    statShuangxiuRatio: document.getElementById('statShuangxiuRatio'),
    statWarningCount: document.getElementById('statWarningCount'),
    statTotalReviews: document.getElementById('statTotalReviews'),

    // 控制栏
    searchInput: document.getElementById('searchInput'),
    searchClearBtn: document.getElementById('searchClearBtn'),
    filterChips: document.querySelectorAll('.filter-chip'),
    sortSelect: document.getElementById('sortSelect'),
    btnViewCard: document.getElementById('btnViewCard'),
    btnViewTable: document.getElementById('btnViewTable'),
    resultsCount: document.getElementById('resultsCount'),

    // 展示容器
    cardsContainer: document.getElementById('cardsContainer'),
    tableViewContainer: document.getElementById('tableViewContainer'),
    tableBody: document.getElementById('tableBody'),
    emptyState: document.getElementById('emptyState'),

    // 拖拽与上传
    dropzone: document.getElementById('dropzone'),
    fileInput: document.getElementById('fileInput'),

    // 顶部操作按钮
    btnDownloadTemplate: document.getElementById('btnDownloadTemplate'),
    btnExportExcel: document.getElementById('btnExportExcel'),
    btnAddCompany: document.getElementById('btnAddCompany'),
    btnResetData: document.getElementById('btnResetData'),

    // 企业详情弹窗
    detailModal: document.getElementById('detailModal'),
    detailModalClose: document.getElementById('detailModalClose'),
    detailCompanyName: document.getElementById('detailCompanyName'),
    detailBadge: document.getElementById('detailBadge'),
    detailHours: document.getElementById('detailHours'),
    detailIndustry: document.getElementById('detailIndustry'),
    detailReviewCount: document.getElementById('detailReviewCount'),
    detailReviewsList: document.getElementById('detailReviewsList'),
    detailNewReviewText: document.getElementById('detailNewReviewText'),
    btnSubmitDetailReview: document.getElementById('btnSubmitDetailReview'),

    // 新增企业弹窗
    addModal: document.getElementById('addModal'),
    addModalClose: document.getElementById('addModalClose'),
    addCompanyForm: document.getElementById('addCompanyForm'),
    btnCancelAdd: document.getElementById('btnCancelAdd'),

    // 永久存档与开源收录弹窗
    archiveModal: document.getElementById('archiveModal'),
    archiveModalClose: document.getElementById('archiveModalClose'),
    archivePreviewCode: document.getElementById('archivePreviewCode'),
    btnSubmitGithubIssue: document.getElementById('btnSubmitGithubIssue'),
    btnCopyArchiveText: document.getElementById('btnCopyArchiveText'),

    // Toast
    toastContainer: document.getElementById('toastContainer')
  };

  /**
   * 初始化应用
   */
  async function init() {
    setupEventListeners();
    setupDialogDismissal(elements.detailModal);
    setupDialogDismissal(elements.addModal);
    setupDialogDismissal(elements.archiveModal);

    // 优先尝试自动从仓库拉取最新的 data.xlsx
    await autoSyncExcelData();
  }

  /**
   * 自动从远程/本地仓库拉取 data.xlsx
   */
  async function autoSyncExcelData() {
    updateSyncStatus('loading', '正在自动同步 data.xlsx 最新数据...');

    try {
      // 只有在 http: 或 https: 协议下尝试 fetch (如 GitHub Pages 或 本地 HTTP 服务)
      const fetchUrl = `${CONFIG.excelUrl}?_t=${Date.now()}`;
      const response = await fetch(fetchUrl);

      if (!response.ok) {
        throw new Error(`无法获取文件 (HTTP ${response.status})`);
      }

      const arrayBuffer = await response.arrayBuffer();
      const parsedList = parseExcelBuffer(arrayBuffer);

      if (parsedList && parsedList.length > 0) {
        state.companies = parsedList;
        saveData();
        render();

        const timeStr = new Date().toLocaleTimeString('zh-CN', { hour12: false });
        updateSyncStatus('success', `🟢 数据已与 data.xlsx 自动同步 (${parsedList.length} 家企业 · ${timeStr})`);
        return;
      } else {
        throw new Error('未能从 data.xlsx 解析到有效企业数据');
      }
    } catch (err) {
      console.warn('自动同步远程 data.xlsx 提示/失败:', err.message);

      // 回退至本地数据
      loadFallbackData();

      // 判断是否是本地 file:// 协议
      if (window.location.protocol === 'file:') {
        updateSyncStatus('warning', '📁 当前使用本地 file:// 协议浏览；上传至 GitHub Pages (https://) 后将全自动实时同步 data.xlsx！');
      } else {
        updateSyncStatus('warning', `⚠️ 自动同步提示：${err.message || '已载入本地数据'}`);
      }
    }
  }

  /**
   * 更新同步横幅状态
   */
  function updateSyncStatus(type, text) {
    if (!elements.syncStatusText || !elements.syncDot) return;

    elements.syncStatusText.textContent = text;
    if (type === 'loading') {
      elements.syncDot.className = 'sync-dot warning';
    } else if (type === 'success') {
      elements.syncDot.className = 'sync-dot';
    } else {
      elements.syncDot.className = 'sync-dot warning';
    }
  }

  /**
   * 回退数据装载逻辑 (读取 localStorage 或默认示例数据)
   */
  function loadFallbackData() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) {
          state.companies = parsed;
          render();
          return;
        }
      }
    } catch (e) {
      console.warn('读取本地数据失败，将载入默认示例数据', e);
    }

    state.companies = JSON.parse(JSON.stringify(DEFAULT_SAMPLE_DATA));
    saveData();
    render();
  }

  /**
   * 保存当前数据至 localStorage
   */
  function saveData() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state.companies));
    } catch (e) {
      console.error('保存数据至 localStorage 失败', e);
    }
  }

  /**
   * 规范化企业作息类型与视觉样式 (智能识别 双休/单休/大小周/是/否/965/996等)
   */
  function getScheduleMeta(scheduleStr) {
    const raw = (scheduleStr || '').trim();
    const s = raw.toLowerCase();

    // 大小周
    if (s.includes('大小周') || s.includes('隔周')) {
      return {
        category: 'daxiaozhou',
        label: raw || '大小周',
        badgeClass: 'badge-yellow',
        icon: '⚠️'
      };
    }

    // 单休 / 996 / 否
    if (
      s === '否' || s === '无' || s === 'no' || s === 'n' ||
      s.includes('单休') || s.includes('996') || s.includes('007') ||
      s.includes('986') || s.includes('11-11-6') || s.includes('无休') ||
      s.includes('严重加班') || s.includes('不休')
    ) {
      return {
        category: 'danxiu',
        label: raw === '否' ? '非双休(单休/加班)' : (raw || '单休/996'),
        badgeClass: 'badge-red',
        icon: '🔥'
      };
    }

    // 双休 / 965 / 是
    if (
      s === '是' || s === '有' || s === 'yes' || s === 'y' ||
      s.includes('双休') || s.includes('965') || s.includes('不加班')
    ) {
      if (s.includes('弹性') || s.includes('自由') || s.includes('wfh') || s.includes('远程')) {
        return {
          category: 'shuangxiu',
          label: raw === '是' ? '严格双休' : (raw || '弹性双休'),
          badgeClass: 'badge-green',
          icon: '✨'
        };
      }
      return {
        category: 'shuangxiu',
        label: raw === '是' ? '严格双休' : (raw || '严格双休'),
        badgeClass: 'badge-green',
        icon: '✅'
      };
    }

    // 默认或视部门而定
    return {
      category: 'other',
      label: raw || '视部门而定',
      badgeClass: 'badge-info',
      icon: 'ℹ️'
    };
  }

  /**
   * 检查是否符合弹性制
   */
  function isFlexibleSchedule(company) {
    const text = (company.scheduleType + ' ' + company.workingHours).toLowerCase();
    return text.includes('弹性') || text.includes('不打卡') || text.includes('wfh') || text.includes('远程');
  }

  /**
   * 过滤与排序企业列表
   */
  function getFilteredAndSortedCompanies() {
    let result = state.companies.filter(c => {
      // 1. 搜索关键词匹配
      if (state.searchQuery) {
        const q = state.searchQuery.toLowerCase().trim();
        const nameMatch = (c.name || '').toLowerCase().includes(q);
        const industryMatch = (c.industry || '').toLowerCase().includes(q);
        const hoursMatch = (c.workingHours || '').toLowerCase().includes(q);
        const scheduleMatch = (c.scheduleType || '').toLowerCase().includes(q);
        const reviewMatch = (c.reviews || []).some(r => r.toLowerCase().includes(q));

        if (!nameMatch && !industryMatch && !hoursMatch && !scheduleMatch && !reviewMatch) {
          return false;
        }
      }

      // 2. 状态标签筛选
      const meta = getScheduleMeta(c.scheduleType);
      if (state.currentFilter === 'shuangxiu') {
        return meta.category === 'shuangxiu';
      }
      if (state.currentFilter === 'daxiaozhou') {
        return meta.category === 'daxiaozhou';
      }
      if (state.currentFilter === 'danxiu') {
        return meta.category === 'danxiu';
      }
      if (state.currentFilter === 'flexible') {
        return isFlexibleSchedule(c);
      }

      return true;
    });

    // 排序
    if (state.currentSort === 'reviews-desc') {
      result.sort((a, b) => (b.reviews ? b.reviews.length : 0) - (a.reviews ? a.reviews.length : 0));
    } else if (state.currentSort === 'name-asc') {
      result.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'zh-CN'));
    }

    return result;
  }

  /**
   * 更新统计数据看板
   */
  function updateStats() {
    const total = state.companies.length;
    let shuangxiuCount = 0;
    let warningCount = 0;
    let totalReviews = 0;

    state.companies.forEach(c => {
      const meta = getScheduleMeta(c.scheduleType);
      if (meta.category === 'shuangxiu') shuangxiuCount++;
      if (meta.category === 'daxiaozhou' || meta.category === 'danxiu') warningCount++;
      totalReviews += (c.reviews ? c.reviews.length : 0);
    });

    elements.statTotalCompanies.textContent = total;
    elements.statShuangxiuCount.textContent = shuangxiuCount;
    const ratio = total > 0 ? Math.round((shuangxiuCount / total) * 100) : 0;
    elements.statShuangxiuRatio.textContent = `双休占比约 ${ratio}%`;
    elements.statWarningCount.textContent = warningCount;
    elements.statTotalReviews.textContent = totalReviews;

    // 更新筛选按钮上的数量
    elements.filterChips.forEach(chip => {
      const filter = chip.dataset.filter;
      const countEl = chip.querySelector('.chip-count');
      if (!countEl) return;

      if (filter === 'all') {
        countEl.textContent = total;
      } else if (filter === 'shuangxiu') {
        countEl.textContent = shuangxiuCount;
      } else if (filter === 'daxiaozhou') {
        const count = state.companies.filter(c => getScheduleMeta(c.scheduleType).category === 'daxiaozhou').length;
        countEl.textContent = count;
      } else if (filter === 'danxiu') {
        const count = state.companies.filter(c => getScheduleMeta(c.scheduleType).category === 'danxiu').length;
        countEl.textContent = count;
      } else if (filter === 'flexible') {
        const count = state.companies.filter(c => isFlexibleSchedule(c)).length;
        countEl.textContent = count;
      }
    });
  }

  /**
   * 渲染主页面内容
   */
  function render() {
    updateStats();
    const list = getFilteredAndSortedCompanies();

    elements.resultsCount.textContent = list.length;

    if (list.length === 0) {
      elements.cardsContainer.style.display = 'none';
      elements.tableViewContainer.style.display = 'none';
      elements.emptyState.style.display = 'block';
      return;
    }

    elements.emptyState.style.display = 'none';

    if (state.currentView === 'card') {
      elements.cardsContainer.style.display = 'grid';
      elements.tableViewContainer.style.display = 'none';
      renderCards(list);
    } else {
      elements.cardsContainer.style.display = 'none';
      elements.tableViewContainer.style.display = 'block';
      renderTable(list);
    }
  }

  /**
   * 渲染卡片网格
   */
  function renderCards(companies) {
    elements.cardsContainer.innerHTML = '';
    const fragment = document.createDocumentFragment();

    companies.forEach(company => {
      const card = document.createElement('article');
      card.className = 'company-card';

      const meta = getScheduleMeta(company.scheduleType);
      const reviews = company.reviews || [];
      const initial = (company.name || '企').trim().charAt(0);

      // 提取前2条精选评价
      const previewReviews = reviews.slice(0, 2);
      const remainingReviewsCount = Math.max(0, reviews.length - previewReviews.length);

      let reviewsHtml = '';
      if (previewReviews.length > 0) {
        reviewsHtml = previewReviews.map(r => `
          <div class="review-item-snippet" title="${escapeHtml(r)}">
            ${escapeHtml(r)}
          </div>
        `).join('');

        if (remainingReviewsCount > 0) {
          reviewsHtml += `
            <div class="reviews-more-badge" data-action="view-details" data-id="${company.id}">
              + 查看剩余 ${remainingReviewsCount} 条员工评价 »
            </div>
          `;
        }
      } else {
        reviewsHtml = `
          <div class="review-item-snippet" style="color: #94a3b8; font-style: italic;">
            暂无具体评价记录，点击下方按钮添加第一条评价
          </div>
        `;
      }

      card.innerHTML = `
        <div class="card-top">
          <div class="card-header-line">
            <div class="company-title-area">
              <div class="company-avatar">${escapeHtml(initial)}</div>
              <div>
                <h3 class="company-name">${escapeHtml(company.name)}</h3>
                <span class="industry-tag">${escapeHtml(company.industry || '综合领域')}</span>
              </div>
            </div>
            <span class="badge ${meta.badgeClass}">${meta.icon} ${escapeHtml(meta.label)}</span>
          </div>

          <div class="hours-strip">
            <svg class="hours-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <circle cx="12" cy="12" r="10"></circle>
              <polyline points="12 6 12 12 16 14"></polyline>
            </svg>
            <span class="hours-text">${escapeHtml(company.workingHours || '未标注作息时间')}</span>
          </div>

          <div class="reviews-preview">
            ${reviewsHtml}
          </div>
        </div>

        <div class="card-footer">
          <span class="review-count-indicator">共 ${reviews.length} 条真实评价</span>
          <div class="card-actions-wrapper">
            <button class="btn btn-primary btn-sm" data-action="view-details" data-id="${company.id}">
              查看全部评价
            </button>
          </div>
        </div>
      `;

      fragment.appendChild(card);
    });

    elements.cardsContainer.appendChild(fragment);
  }

  /**
   * 渲染表格视图
   */
  function renderTable(companies) {
    elements.tableBody.innerHTML = '';
    const fragment = document.createDocumentFragment();

    companies.forEach((company, index) => {
      const tr = document.createElement('tr');
      const meta = getScheduleMeta(company.scheduleType);
      const reviews = company.reviews || [];
      const firstReview = reviews.length > 0 ? reviews[0] : '暂无评价';

      tr.innerHTML = `
        <td>${index + 1}</td>
        <td class="table-company-cell">${escapeHtml(company.name)}</td>
        <td>
          <span class="badge ${meta.badgeClass}">${meta.icon} ${escapeHtml(meta.label)}</span>
        </td>
        <td>${escapeHtml(company.workingHours || '-')}</td>
        <td>${escapeHtml(company.industry || '-')}</td>
        <td class="table-reviews-cell" title="${escapeHtml(firstReview)}">
          ${escapeHtml(firstReview)} ${reviews.length > 1 ? `<b style="color: var(--primary);">[+${reviews.length - 1}条]</b>` : ''}
        </td>
        <td>
          <button class="btn btn-outline btn-sm" data-action="view-details" data-id="${company.id}">
            详情 (${reviews.length})
          </button>
        </td>
      `;

      fragment.appendChild(tr);
    });

    elements.tableBody.appendChild(fragment);
  }

  /**
   * 打开企业详情弹窗
   */
  function openCompanyDetail(companyId) {
    const company = state.companies.find(c => c.id === companyId);
    if (!company) return;

    state.selectedCompanyId = companyId;
    const meta = getScheduleMeta(company.scheduleType);
    const reviews = company.reviews || [];

    elements.detailCompanyName.textContent = company.name;
    elements.detailBadge.className = `badge ${meta.badgeClass}`;
    elements.detailBadge.textContent = `${meta.icon} ${meta.label}`;
    elements.detailHours.textContent = company.workingHours || '未标明';
    elements.detailIndustry.textContent = company.industry || '未归类';
    elements.detailReviewCount.textContent = `${reviews.length} 条`;

    renderDetailReviewsList(reviews);

    // 清空追加输入框
    elements.detailNewReviewText.value = '';

    if (typeof elements.detailModal.showModal === 'function') {
      elements.detailModal.showModal();
    }
  }

  /**
   * 渲染详情弹窗内的评价条目列表
   */
  function renderDetailReviewsList(reviews) {
    if (reviews.length === 0) {
      elements.detailReviewsList.innerHTML = `
        <div style="text-align:center; padding: 24px; color: var(--text-muted); background: #f8fafc; border-radius: var(--radius-md);">
          该企业暂时还没有收录详细评价，欢迎在下方填写提交！
        </div>
      `;
      return;
    }

    elements.detailReviewsList.innerHTML = reviews.map((rev, idx) => `
      <div class="modal-review-card">
        <div class="modal-review-header">
          <span class="modal-review-badge">评价 #${idx + 1}</span>
          <button class="copy-review-btn" data-action="copy-review" data-text="${escapeHtml(rev)}">
            📋 复制评价
          </button>
        </div>
        <div class="modal-review-body">${escapeHtml(rev)}</div>
      </div>
    `).join('');
  }

  /**
   * 弹窗关闭回落支持 (兼容旧浏览器没有 closedBy 特性的情况)
   */
  function setupDialogDismissal(dialog) {
    if (!dialog) return;
    if (!('closedBy' in HTMLDialogElement.prototype)) {
      dialog.addEventListener('click', (event) => {
        if (event.target !== dialog) return;
        const rect = dialog.getBoundingClientRect();
        const isDialogContent = (
          rect.top <= event.clientY &&
          event.clientY <= rect.top + rect.height &&
          rect.left <= event.clientX &&
          event.clientX <= rect.left + rect.width
        );
        if (isDialogContent) return;
        dialog.close();
      });
    }
  }

  /**
   * 通用 Excel 二进制 Buffer 解析器 (返回企业对象数组)
   */
  function parseExcelBuffer(arrayBuffer) {
    if (typeof XLSX === 'undefined') {
      throw new Error('XLSX 解析库尚未加载');
    }

    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array' });

    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error('Excel 文件中没有找到有效工作表 (Sheet)');
    }

    // 智能遍历所有工作表，优先找到包含“公司”或“企业”表头的数据表
    let rawJson = null;
    let colIndexes = null;

    for (const sName of workbook.SheetNames) {
      const ws = workbook.Sheets[sName];
      const testJson = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (testJson && testJson.length > 1) {
        const testHeaders = testJson[0].map(h => (h ? String(h).trim() : ''));
        const testCols = findColumnIndexes(testHeaders);
        if (testCols.name !== -1) {
          rawJson = testJson;
          colIndexes = testCols;
          break;
        }
      }
    }

    // 回退到首个工作表
    if (!rawJson) {
      const ws = workbook.Sheets[workbook.SheetNames[0]];
      rawJson = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (!rawJson || rawJson.length < 2) {
        throw new Error('工作表内容为空或缺少数据行');
      }
      const headers = rawJson[0].map(h => (h ? String(h).trim() : ''));
      colIndexes = findColumnIndexes(headers);
      if (colIndexes.name === -1) {
        throw new Error('未在Excel首行找到“公司名称”或“企业名称”列，请核对表头');
      }
    }

    const parsedCompaniesMap = new Map();

    for (let i = 1; i < rawJson.length; i++) {
      const row = rawJson[i];
      if (!row || row.length === 0) continue;

      const companyName = colIndexes.name !== -1 && row[colIndexes.name] ? String(row[colIndexes.name]).trim() : '';
      if (!companyName) continue;

      const scheduleType = colIndexes.schedule !== -1 && row[colIndexes.schedule] ? String(row[colIndexes.schedule]).trim() : '未标明';
      const workingHours = colIndexes.hours !== -1 && row[colIndexes.hours] ? String(row[colIndexes.hours]).trim() : '未标注时间';
      const industry = colIndexes.industry !== -1 && row[colIndexes.industry] ? String(row[colIndexes.industry]).trim() : '未分类';

      let newReviews = [];
      if (colIndexes.reviews.length > 0) {
        colIndexes.reviews.forEach(idx => {
          if (row[idx]) {
            const cellText = String(row[idx]).trim();
            const splitReviews = splitMultiLineReviews(cellText);
            newReviews.push(...splitReviews);
          }
        });
      }

      const key = companyName.toLowerCase();
      if (parsedCompaniesMap.has(key)) {
        const existing = parsedCompaniesMap.get(key);
        newReviews.forEach(r => {
          if (!existing.reviews.includes(r)) {
            existing.reviews.push(r);
          }
        });
        if (existing.workingHours === '未标注时间' && workingHours !== '未标注时间') {
          existing.workingHours = workingHours;
        }
        if (existing.industry === '未分类' && industry !== '未分类') {
          existing.industry = industry;
        }
      } else {
        parsedCompaniesMap.set(key, {
          id: 'comp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6),
          name: companyName,
          scheduleType: scheduleType,
          workingHours: workingHours,
          industry: industry,
          reviews: newReviews
        });
      }
    }

    return Array.from(parsedCompaniesMap.values());
  }

  /**
   * 手动解析上传的 Excel 文件
   */
  function parseUploadedFile(file) {
    showToast(`正在解析文件：${file.name}...`, 'info');

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const parsedList = parseExcelBuffer(e.target.result);
        if (parsedList.length === 0) {
          throw new Error('未从表格中解析到有效的公司数据');
        }

        state.companies = parsedList;
        saveData();
        render();

        const totalReviews = parsedList.reduce((sum, c) => sum + c.reviews.length, 0);
        updateSyncStatus('success', `📁 本地导入生效：${file.name} (${parsedList.length} 家企业)`);
        showToast(`🎉 成功导入！共解析到 ${parsedList.length} 家企业、${totalReviews} 条评价`, 'success');
      } catch (err) {
        console.error(err);
        showToast('导入失败：' + (err.message || '文件格式不正确'), 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  /**
   * 自动探测与模糊匹配列索引
   */
  function findColumnIndexes(headers) {
    const indexes = {
      name: -1,
      schedule: -1,
      hours: -1,
      industry: -1,
      reviews: []
    };

    headers.forEach((header, idx) => {
      const h = header.toLowerCase().replace(/[\s\r\n]/g, '');

      // 公司名
      if (indexes.name === -1 && (h.includes('公司') || h.includes('企业') || h.includes('单位') || h === 'company' || h === 'name')) {
        indexes.name = idx;
        return;
      }

      // 作息制度 / 是否双休
      if (indexes.schedule === -1 && (h.includes('作息') || h.includes('双休') || h.includes('制度') || h.includes('加班') || h === 'schedule')) {
        indexes.schedule = idx;
        return;
      }

      // 工作时间 / 工时
      if (indexes.hours === -1 && (h.includes('时间') || h.includes('上下班') || h.includes('工时') || h === 'hours' || h === 'time')) {
        indexes.hours = idx;
        return;
      }

      // 行业类别
      if (indexes.industry === -1 && (h.includes('行业') || h.includes('类别') || h.includes('领域') || h === 'industry')) {
        indexes.industry = idx;
        return;
      }

      // 员工评价 / 评论
      if (h.includes('评价') || h.includes('评论') || h.includes('反馈') || h.includes('口碑') || h.includes('备注') || h.includes('review') || h.includes('comment')) {
        indexes.reviews.push(idx);
        return;
      }
    });

    if (indexes.reviews.length === 0) {
      headers.forEach((h, idx) => {
        if (idx !== indexes.name && idx !== indexes.schedule && idx !== indexes.hours && idx !== indexes.industry) {
          indexes.reviews.push(idx);
        }
      });
    }

    return indexes;
  }

  /**
   * 智能拆分多行或符号分隔的评价内容 (支持换行符、分号、以及行内连续数字编号如 1.xxx。2.yyy)
   */
  function splitMultiLineReviews(text) {
    if (!text) return [];
    let raw = String(text).trim();

    // 先按换行符拆分
    let lines = raw.split(/\r?\n/);

    // 如果没有换行，且含有分号
    if (lines.length === 1 && (raw.includes('；') || raw.includes(';'))) {
      lines = raw.split(/[；;]/);
    }

    // 如果仍然只有1行，但含有连续的数字编号如 '1. xxx 2. yyy' 或 '1.xxx。2.yyy'
    if (lines.length === 1 && /(?:[。；;\s]|^)\s*[1-9]\d*[\.、]/.test(raw)) {
      const parts = raw.split(/(?:[。；;\s]+|^)(?=[1-9]\d*[\.、\)])/);
      if (parts.length > 1) {
        lines = parts;
      }
    }

    const cleanLines = [];
    lines.forEach(line => {
      let trimmed = line.trim();
      if (!trimmed) return;
      trimmed = trimmed.replace(/^(\d+[\.、\)]|\(\d+\)|[-*•·]\s*)/, '').trim();
      trimmed = trimmed.replace(/^[。；;,\s]+|[。；;,\s]+$/g, '').trim();
      if (trimmed) cleanLines.push(trimmed);
    });

    return cleanLines;
  }

  /**
   * 导出当前数据为 Excel
   */
  function exportToExcel() {
    if (typeof XLSX === 'undefined') {
      showToast('XLSX导出库未就绪', 'error');
      return;
    }

    const exportRows = state.companies.map(c => {
      const formattedReviews = (c.reviews || [])
        .map((r, i) => `${i + 1}. ${r}`)
        .join('\n');

      return {
        "公司名称": c.name,
        "作息制度": c.scheduleType,
        "工作时间": c.workingHours,
        "行业类别": c.industry,
        "员工评价": formattedReviews
      };
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportRows, {
      header: ["公司名称", "作息制度", "工作时间", "行业类别", "员工评价"]
    });

    ws['!cols'] = [
      { wch: 25 },
      { wch: 15 },
      { wch: 22 },
      { wch: 18 },
      { wch: 70 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, "企业作息与评价");

    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `企业作息与评价数据_${dateStr}.xlsx`);
    showToast('已成功导出当前最新数据为 Excel 表格！', 'success');
  }

  /**
   * 下载标准示例模板
   */
  function downloadTemplate() {
    if (typeof XLSX === 'undefined') {
      showToast('XLSX未就绪', 'error');
      return;
    }

    const templateRows = [
      {
        "公司名称": "示例公司A (某知名科技)",
        "作息制度": "严格双休",
        "工作时间": "09:30 - 18:30",
        "行业类别": "互联网 / 软件",
        "员工评价": "1. 严格双休，周五准点下班，周末不发工作消息。\n2. 年假充裕，请假审批快。\n3. 食堂三餐免费，氛围较轻松。"
      },
      {
        "公司名称": "示例公司B (某电商大厂)",
        "作息制度": "大小周",
        "工作时间": "10:00 - 21:00",
        "行业类别": "电商运营",
        "员工评价": "1. 大小周工作制，隔周周六加班，有加班费。\n2. 业务指标考核严格，KPI压力大。\n3. 薪资待遇在行业内算中上水平。"
      },
      {
        "公司名称": "示例公司C (某外企研发)",
        "作息制度": "弹性双休",
        "工作时间": "09:00 - 17:30 (不打卡)",
        "行业类别": "外资外企",
        "员工评价": "1. 不打卡，支持每周2天居家办公(WFH)。\n2. 节假日严格遵照法定，绝不调休补班。\n3. 人际关系扁平，无繁文缛节。"
      }
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(templateRows, {
      header: ["公司名称", "作息制度", "工作时间", "行业类别", "员工评价"]
    });

    ws['!cols'] = [
      { wch: 25 },
      { wch: 15 },
      { wch: 22 },
      { wch: 18 },
      { wch: 65 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, "填写模板与示例");
    XLSX.writeFile(wb, "企业作息与评价数据表_标准模板.xlsx");
    showToast('已开始下载标准 Excel 填写模板！', 'success');
  }

  /**
   * 打开永久存档与 GitHub 提交指引弹窗
   */
  function showArchiveGuidanceModal(company) {
    state.lastAddedCompany = company;

    // 格式化为 Markdown 提交模板
    const reviewsMd = (company.reviews || []).map((r, i) => `${i + 1}. ${r}`).join('\n');
    const issueBody = `### 🏢 企业基本信息
- **公司名称**：${company.name}
- **作息制度**：${company.scheduleType}
- **工作时间**：${company.workingHours}
- **所属行业**：${company.industry}

### 💬 员工真实评价 / 口碑反馈
${reviewsMd || '暂无评价'}

---
*来自网页访客录入提交，请管理员核实后并入 data.xlsx 永久归档*`;

    elements.archivePreviewCode.textContent = issueBody;

    // 设置 GitHub Issue 快捷链接
    const issueTitle = `【企业录入申请】${company.name} - ${company.scheduleType}`;
    const targetRepo = CONFIG.githubRepoUrl.replace(/\/$/, '');
    const issueUrl = `${targetRepo}/issues/new?title=${encodeURIComponent(issueTitle)}&body=${encodeURIComponent(issueBody)}`;
    elements.btnSubmitGithubIssue.href = issueUrl;

    if (typeof elements.archiveModal.showModal === 'function') {
      elements.archiveModal.showModal();
    }
  }

  /**
   * 统一事件监听绑定
   */
  function setupEventListeners() {
    // 重新从仓库同步
    elements.btnReloadRemote.addEventListener('click', () => {
      autoSyncExcelData();
    });

    // 切换折叠本地上传拖拽区域
    elements.btnToggleDropzone.addEventListener('click', () => {
      const isHidden = elements.dropzone.classList.toggle('hidden');
      elements.btnToggleDropzone.textContent = isHidden ? '📁 本地测试其他表格 ▾' : '📁 收起本地导入测试 ▴';
    });

    // 搜索输入
    elements.searchInput.addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      render();
    });

    // 搜索清空
    elements.searchClearBtn.addEventListener('click', () => {
      elements.searchInput.value = '';
      state.searchQuery = '';
      render();
      elements.searchInput.focus();
    });

    // 状态过滤标签点击
    elements.filterChips.forEach(chip => {
      chip.addEventListener('click', () => {
        elements.filterChips.forEach(c => c.classList.remove('active'));
        chip.classList.add('active');
        state.currentFilter = chip.dataset.filter;
        render();
      });
    });

    // 排序变化
    elements.sortSelect.addEventListener('change', (e) => {
      state.currentSort = e.target.value;
      render();
    });

    // 视图切换：卡片
    elements.btnViewCard.addEventListener('click', () => {
      elements.btnViewCard.classList.add('active');
      elements.btnViewTable.classList.remove('active');
      state.currentView = 'card';
      render();
    });

    // 视图切换：表格
    elements.btnViewTable.addEventListener('click', () => {
      elements.btnViewTable.classList.add('active');
      elements.btnViewCard.classList.remove('active');
      state.currentView = 'table';
      render();
    });

    // 卡片与表格委托事件 (查看详情)
    document.addEventListener('click', (e) => {
      const target = e.target.closest('[data-action]');
      if (!target) return;

      const action = target.dataset.action;
      if (action === 'view-details') {
        const id = target.dataset.id;
        openCompanyDetail(id);
      } else if (action === 'copy-review') {
        const text = target.dataset.text;
        if (navigator.clipboard) {
          navigator.clipboard.writeText(text).then(() => {
            showToast('已将评价复制到剪贴板！', 'success');
          });
        }
      }
    });

    // 详情弹窗关闭
    elements.detailModalClose.addEventListener('click', () => {
      elements.detailModal.close();
    });

    // 详情弹窗中提交新评价
    elements.btnSubmitDetailReview.addEventListener('click', () => {
      const text = elements.detailNewReviewText.value.trim();
      if (!text) {
        showToast('请输入评价内容后再提交', 'error');
        return;
      }
      if (!state.selectedCompanyId) return;

      const comp = state.companies.find(c => c.id === state.selectedCompanyId);
      if (comp) {
        if (!comp.reviews) comp.reviews = [];
        comp.reviews.unshift(text);
        saveData();
        render();
        renderDetailReviewsList(comp.reviews);
        elements.detailReviewCount.textContent = `${comp.reviews.length} 条`;
        elements.detailNewReviewText.value = '';
        showToast('新评价已在本地成功追加！', 'success');
      }
    });

    // 拖拽与文件上传
    elements.dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      elements.dropzone.classList.add('drag-over');
    });

    elements.dropzone.addEventListener('dragleave', () => {
      elements.dropzone.classList.remove('drag-over');
    });

    elements.dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      elements.dropzone.classList.remove('drag-over');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        parseUploadedFile(e.dataTransfer.files[0]);
      }
    });

    elements.fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        parseUploadedFile(e.target.files[0]);
        e.target.value = '';
      }
    });

    // 顶部操作：模板下载
    elements.btnDownloadTemplate.addEventListener('click', downloadTemplate);

    // 顶部操作：导出Excel
    elements.btnExportExcel.addEventListener('click', exportToExcel);

    // 顶部操作：录入企业
    elements.btnAddCompany.addEventListener('click', () => {
      elements.addCompanyForm.reset();
      elements.addModal.showModal();
    });

    // 新增企业弹窗关闭
    elements.addModalClose.addEventListener('click', () => {
      elements.addModal.close();
    });

    elements.btnCancelAdd.addEventListener('click', () => {
      elements.addModal.close();
    });

    // 新增企业提交表单
    elements.addCompanyForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const name = document.getElementById('addCompanyName').value.trim();
      const schedule = document.getElementById('addCompanySchedule').value.trim();
      const hours = document.getElementById('addCompanyHours').value.trim();
      const industry = document.getElementById('addCompanyIndustry').value.trim();
      const reviewText = document.getElementById('addCompanyReviews').value.trim();

      if (!name) {
        showToast('请填写公司名称', 'error');
        return;
      }

      const reviews = splitMultiLineReviews(reviewText);

      const newCompany = {
        id: 'comp_' + Date.now(),
        name: name,
        scheduleType: schedule || '严格双休',
        workingHours: hours || '09:00 - 18:00',
        industry: industry || '综合领域',
        reviews: reviews
      };

      // 1. 本地立即生效
      state.companies.unshift(newCompany);
      saveData();
      render();
      elements.addModal.close();

      showToast(`已在当前浏览器录入：${name}`, 'success');

      // 2. 引导提交 GitHub Issue 永久存档
      showArchiveGuidanceModal(newCompany);
    });

    // 归档弹窗关闭
    elements.archiveModalClose.addEventListener('click', () => {
      elements.archiveModal.close();
    });

    // 复制归档文本
    elements.btnCopyArchiveText.addEventListener('click', () => {
      const code = elements.archivePreviewCode.textContent;
      if (navigator.clipboard) {
        navigator.clipboard.writeText(code).then(() => {
          showToast('已复制申请内容，可发送给管理员！', 'success');
        });
      }
    });

    // 恢复默认示例数据
    elements.btnResetData.addEventListener('click', () => {
      if (confirm('确定要恢复为内置的12家企业精选示例数据吗？当前修改的数据将被替换。')) {
        state.companies = JSON.parse(JSON.stringify(DEFAULT_SAMPLE_DATA));
        saveData();
        render();
        showToast('已成功重置为精选示例数据！', 'info');
      }
    });
  }

  /**
   * 轻量 Toast 提示框
   */
  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '❌';

    toast.innerHTML = `<span>${icon}</span> <span>${escapeHtml(message)}</span>`;
    elements.toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
      }, 200);
    }, 3200);
  }

  /**
   * HTML 转义防 XSS
   */
  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // 页面加载完成后启动
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
