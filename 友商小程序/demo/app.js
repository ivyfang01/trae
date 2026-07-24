const customers = {
  '1': {
    customerId: '6677766',
    name: '泰州市红羽禽业有限公司',
    phone: '18019249090',
    status: '意向客户',
    breedType: '猪',
    scale: '1000',
    scaleUnit: '头',
    region: '河南省驻马店市',
    address: '航鹤路1号',
  },
  '2': {
    customerId: '6677767',
    name: '上海浦耀贸易有限公司',
    phone: '18019999090',
    status: '已失效',
    breedType: '鸡',
    scale: '300000',
    scaleUnit: '羽',
    region: '河南省驻马店市',
    address: '县镇AA村AA道',
  },
};

const breedUnits = {
  鸡: '羽',
  猪: '头',
  鸭: '羽',
};

const REGION_OPTIONS = ['河南省驻马店市', '上海市', '江苏省泰州市'];

const MOCK_LOCATION = {
  region: '河南省驻马店市',
  address: '县镇AA村AA道',
};

const selectOptions = {
  status: ['意向客户', '已成交', '已失效'],
  breedType: ['鸡', '猪', '鸭'],
  addStatus: ['未成交客户', '意向客户', '已成交', '已失效'],
  addScale: ['1-5w', '5-10w', '10-50w', '50w以上'],
  region: REGION_OPTIONS,
  addRegion: REGION_OPTIONS,
  dupRegion: REGION_OPTIONS,
  empBrand: ['嘉吉', '粮巴巴', '正大', '双胞胎'],
};

let activeSelectEl = null;

function getCurrentCustomerId() {
  const params = new URLSearchParams(window.location.search);
  return params.get('id') || '1';
}

function showOverlay() {
  const overlay = document.getElementById('overlay');
  if (overlay) overlay.hidden = false;
}

function hideOverlay() {
  const overlay = document.getElementById('overlay');
  if (overlay) overlay.hidden = true;
  const alert = document.getElementById('customAlert');
  const toast = document.getElementById('toast');
  if (alert) alert.hidden = true;
  if (toast) toast.hidden = true;
}

function showAlert(title, message) {
  showOverlay();
  const alert = document.getElementById('customAlert');
  const toast = document.getElementById('toast');
  if (!alert) return;
  document.getElementById('alertTitle').textContent = title;
  document.getElementById('alertMessage').textContent = message;
  alert.hidden = false;
  if (toast) toast.hidden = true;
}

function showToast(message, duration = 2000, onDone) {
  showOverlay();
  const alert = document.getElementById('customAlert');
  const toast = document.getElementById('toast');
  if (alert) alert.hidden = true;
  if (!toast) return;
  document.getElementById('toastMessage').textContent = message;
  toast.hidden = false;

  setTimeout(() => {
    hideOverlay();
    if (onDone) onDone();
  }, duration);
}

function isPhoneChanged(currentId, phone) {
  const originalPhone = (customers[currentId] || customers['1']).phone;
  return phone !== originalPhone;
}

function setCustomSelectValue(fieldId, value) {
  const input = document.getElementById(fieldId);
  const display = document.getElementById(`${fieldId}Display`);
  input.value = value;
  display.textContent = value || '请选择';
  display.classList.toggle('form-value-placeholder', !value);
}

function closePicker() {
  document.getElementById('pickerOverlay').hidden = true;
  activeSelectEl = null;
}

function openPicker(selectEl) {
  const fieldId = selectEl.dataset.target;
  const options = selectOptions[fieldId];
  const currentValue = document.getElementById(fieldId).value;
  const list = document.getElementById('pickerList');

  activeSelectEl = selectEl;
  list.innerHTML = options.map((option) => `
    <li>
      <button
        type="button"
        class="picker-option${option === currentValue ? ' active' : ''}"
        data-value="${option}"
      >${option}</button>
    </li>
  `).join('');

  list.querySelectorAll('.picker-option').forEach((btn) => {
    btn.addEventListener('click', () => {
      const value = btn.dataset.value;
      const input = document.getElementById(fieldId);
      setCustomSelectValue(fieldId, value);
      closePicker();
      input.dispatchEvent(new Event('change', { bubbles: true }));
    });
  });

  document.getElementById('pickerOverlay').hidden = false;
}

function initCustomSelects() {
  document.querySelectorAll('.custom-select').forEach((selectEl) => {
    selectEl.querySelector('.custom-select-trigger').addEventListener('click', (e) => {
      openPicker(selectEl);
      e.currentTarget.blur();
    });
  });

  document.getElementById('pickerMask').addEventListener('click', closePicker);
  document.getElementById('pickerCancel').addEventListener('click', closePicker);
}

function loadCustomerForm() {
  const id = getCurrentCustomerId();
  const data = customers[id] || customers['1'];

  document.getElementById('name').value = data.name;
  document.getElementById('phone').value = data.phone;
  setCustomSelectValue('status', data.status);
  setCustomSelectValue('breedType', data.breedType);
  document.getElementById('scale').value = data.scale;
  document.getElementById('scaleUnit').textContent = data.scaleUnit;
  setCustomSelectValue('region', data.region);
  setLocationDisplay('editLocationText', 'editLocationBtn', data.address);
  document.getElementById('editForm').dataset.address = data.address;
}

function initEditPage() {
  initCustomSelects();
  loadCustomerForm();
  bindLocationPicker({
    btnId: 'editLocationBtn',
    textId: 'editLocationText',
    regionFieldId: 'region',
    formId: 'editForm',
  });

  const breedType = document.getElementById('breedType');
  const scaleUnit = document.getElementById('scaleUnit');

  breedType.addEventListener('change', () => {
    scaleUnit.textContent = breedUnits[breedType.value] || '只';
  });

  document.getElementById('alertConfirm').addEventListener('click', hideOverlay);

  document.getElementById('editForm').addEventListener('submit', (e) => {
    e.preventDefault();

    const currentId = getCurrentCustomerId();
    const phone = document.getElementById('phone').value.trim();

    if (isPhoneChanged(currentId, phone)) {
      showAlert('修改提醒', '客户已存在，修改失败');
      return;
    }

    showToast('修改成功', 2000, () => {
      window.location.href = 'index.html';
    });
  });
}

if (document.getElementById('editForm')) {
  initEditPage();
}

function findDuplicateCustomer(name, phone) {
  return Object.values(customers).find(
    (customer) => customer.phone === phone || (name && customer.name === name)
  );
}

function getDuplicateReasons(name, phone, duplicate) {
  const reasons = [];
  if (name && duplicate.name === name) {
    reasons.push('已有类似的客户名字备案:**慧耀**。');
  }
  if (phone && duplicate.phone === phone) {
    reasons.push('该手机号已备案。');
  }
  reasons.push('100米范围内已有客户备案。');
  return reasons;
}

function closeResultModal() {
  document.getElementById('resultOverlay').hidden = true;
}

function renderResultRow(label, value) {
  return `
    <div class="result-row">
      <span class="result-label">${label}</span>
      <span class="result-value">${value}</span>
    </div>
  `;
}

function showDuplicateResult(duplicate, reasons) {
  const body = document.getElementById('resultBody');
  const footer = document.getElementById('resultFooter');

  body.innerHTML = `
    <div class="result-info">
      ${renderResultRow('客户名称', duplicate.name)}
      
    </div>
    <div class="result-section">
      <h3 class="result-section-title">结果</h3>
      <ol class="result-list">
        ${reasons.map((reason) => `<li>${reason}</li>`).join('')}
      </ol>
    </div>
  `;

  footer.innerHTML = `
    <button type="button" class="btn btn-cancel btn-block result-back-btn">返回</button>
  `;

  footer.querySelector('.result-back-btn').addEventListener('click', closeResultModal);
  document.getElementById('resultOverlay').hidden = false;
}

function showNoDuplicateResult({ name, phone, region, address }) {
  const body = document.getElementById('resultBody');
  const footer = document.getElementById('resultFooter');

  body.innerHTML = `
    <div class="result-info">
      ${renderResultRow('客户名称', name)}
      ${renderResultRow('客户手机号', phone)}
      ${renderResultRow('客户地址', address || '未选择')}
    </div>
    <div class="result-section">
      <h3 class="result-section-title">结果</h3>
      <p class="result-success">无相似客户备案</p>
    </div>
  `;

  footer.innerHTML = `
    <button type="button" class="btn btn-cancel result-back-btn">返回</button>
    <button type="button" class="btn btn-submit result-add-btn">添加客户</button>
  `;

  footer.querySelector('.result-back-btn').addEventListener('click', closeResultModal);
  footer.querySelector('.result-add-btn').addEventListener('click', () => {
    sessionStorage.setItem('addCustomerDraft', JSON.stringify({ name, phone, region, address }));
    window.location.href = 'add-customer.html';
  });

  document.getElementById('resultOverlay').hidden = false;
}

function inferRegionFromAddress(address) {
  return REGION_OPTIONS.find((item) => address.includes(item)) || '';
}

function setLocationDisplay(textId, btnId, address) {
  const textEl = document.getElementById(textId);
  const btnEl = document.getElementById(btnId);
  if (!textEl || !btnEl) return;
  textEl.textContent = address;
  textEl.classList.toggle('form-value-placeholder', !address);
  btnEl.classList.toggle('is-selected', !!address);
}

function bindLocationPicker({ btnId, textId, regionFieldId, formId }) {
  const btn = document.getElementById(btnId);
  if (!btn) return;

  btn.addEventListener('click', (e) => {
    const { region, address } = MOCK_LOCATION;
    setLocationDisplay(textId, btnId, address);
    if (regionFieldId) {
      setCustomSelectValue(regionFieldId, region);
    }
    const form = formId ? document.getElementById(formId) : btn.closest('form');
    if (form) {
      form.dataset.address = address;
    }
    e.currentTarget.blur();
  });
}

function getAddressFromForm(formId) {
  const form = document.getElementById(formId);
  if (!form) return { region: '', address: '' };
  const regionInput = form.querySelector('[name="region"]');
  return {
    region: regionInput ? regionInput.value.trim() : '',
    address: form.dataset.address || '',
  };
}

function initDuplicateCheckPage() {
  initCustomSelects();
  bindLocationPicker({
    btnId: 'dupLocationBtn',
    textId: 'dupLocationText',
    regionFieldId: 'dupRegion',
    formId: 'duplicateForm',
  });

  document.getElementById('alertConfirm').addEventListener('click', hideOverlay);
  document.getElementById('resultClose').addEventListener('click', closeResultModal);

  document.getElementById('duplicateForm').addEventListener('submit', (e) => {
    e.preventDefault();

    const name = document.getElementById('dupName').value.trim();
    const phone = document.getElementById('dupPhone').value.trim();
    const { region, address } = getAddressFromForm('duplicateForm');

    if (!name) {
      showAlert('查重提醒', '请填写客户名字');
      return;
    }

    const duplicate = findDuplicateCustomer(name, phone);
    if (duplicate) {
      const reasons = getDuplicateReasons(name, phone, duplicate);
      showDuplicateResult(duplicate, reasons);
      return;
    }

    showNoDuplicateResult({ name, phone, region, address });
  });
}

if (document.getElementById('duplicateForm')) {
  initDuplicateCheckPage();
}

function loadAddCustomerForm() {
  const draftRaw = sessionStorage.getItem('addCustomerDraft');
  if (!draftRaw) return null;

  sessionStorage.removeItem('addCustomerDraft');
  return JSON.parse(draftRaw);
}

function applyAddCustomerDraft(draft) {
  if (!draft) return;

  if (draft.name) {
    document.getElementById('addName').value = draft.name;
  }
  if (draft.phone) {
    document.getElementById('addPhone').value = draft.phone;
  }
  if (draft.region) {
    setCustomSelectValue('addRegion', draft.region);
  }
  if (draft.address) {
    document.getElementById('addForm').dataset.address = draft.address;
    setLocationDisplay('addLocationText', 'addLocationBtn', draft.address);
    if (!draft.region) {
      const region = inferRegionFromAddress(draft.address);
      if (region) {
        setCustomSelectValue('addRegion', region);
      }
    }
  }
}

function initAddCustomerPage() {
  initCustomSelects();
  applyAddCustomerDraft(loadAddCustomerForm());
  bindLocationPicker({
    btnId: 'addLocationBtn',
    textId: 'addLocationText',
    regionFieldId: 'addRegion',
    formId: 'addForm',
  });

  document.getElementById('alertConfirm').addEventListener('click', hideOverlay);

  document.getElementById('addForm').addEventListener('submit', (e) => {
    e.preventDefault();

    const name = document.getElementById('addName').value.trim();
    const phone = document.getElementById('addPhone').value.trim();
    const region = document.getElementById('addRegion').value;
    const selectedAddress = document.getElementById('addForm').dataset.address || '';

    if (!name) {
      showAlert('添加提醒', '请填写客户名字');
      return;
    }

    if (!region) {
      showAlert('添加提醒', '请选择地址区域');
      return;
    }

    if (!selectedAddress) {
      showAlert('添加提醒', '请选择定位地址');
      return;
    }

    const duplicate = findDuplicateCustomer(name, phone);
    if (duplicate) {
      showAlert('添加提醒', '客户已存在，添加失败');
      return;
    }

    showToast('添加成功', 2000, () => {
      window.location.href = 'index.html';
    });
  });
}

if (document.getElementById('addForm')) {
  initAddCustomerPage();
}

const duplicateCheckRecords = [
  {
    id: '1',
    companyName: 'A有限公司',
    time: '2026-05-01 12:00',
    operatedAt: '2026-05-01 12:00:00',
    operator: '方伟萍',
    hasDuplicate: false,
    name: 'A有限公司',
    phone: '13800001111',
    address: '河南省驻马店市县镇AA村AA道',
  },
  {
    id: '2',
    companyName: 'B有限公司',
    time: '2026-06-01 09:00',
    operatedAt: '2026-06-01 09:00:00',
    operator: '方伟萍',
    hasDuplicate: true,
    name: '泰州市红羽禽业有限公司',
    phone: '18019249090',
    address: '河南省驻马店市县镇AA村AA道',
    reasons: [
      '已有类似的名字备案。',
      '该手机号已备案。',
      '100米范围内已有客户备案。',
    ],
  },
];

const addCustomerRecords = [
  {
    id: '1',
    companyName: 'C有限公司',
    time: '2026-06-02 10:00',
    operatedAt: '2026-06-02 10:00:00',
    operator: '方伟萍',
    success: true,
    name: 'C有限公司',
    phone: '13900002222',
    status: '未成交客户',
    scale: '5-10w',
    region: '河南省驻马店市',
    address: '河南省驻马店市县镇BB村BB道',
  },
  {
    id: '2',
    companyName: '泰州市红羽禽业有限公司',
    time: '2026-06-02 11:00',
    operatedAt: '2026-06-02 11:00:00',
    operator: '方伟萍',
    success: false,
    name: '泰州市红羽禽业有限公司',
    phone: '18019249090',
    status: '未成交客户',
    scale: '5-10w',
    region: '河南省驻马店市',
    address: '河南省驻马店市航鹤路1号',
    failReason: '客户已存在，添加失败',
    
  },
];

const invalidCustomerRecords = [
  {
    id: '1',
    companyName: '泰州市红羽禽业有限公司',
    time: '2026-06-01 09:00',
    operatedAt: '2026-06-01 09:00:00',
    name: '泰州市红羽禽业有限公司',
    phone: '',
    address: '',
    customerId: '8899888',
    result: '客户已失效',
    notification: '已短信通知方伟萍',
  },
  {
    id: '2',
    companyName: '上海浦耀贸易有限公司',
    time: '2026-05-28 15:30',
    operatedAt: '2026-05-28 15:30:00',
    name: '上海浦耀贸易有限公司',
    phone: '18019999090',
    address: '河南省驻马店市县镇AA村AA道',
    customerId: '6677767',
    result: '客户已失效',
    notification: '已短信通知方伟萍',
  },
];

const editCustomerRecords = [
  {
    id: '1',
    companyName: '郑州市红羽禽业有限公司',
    time: '2026-06-01 09:00',
    operatedAt: '2026-06-01 09:00:00',
    success: true,
    result: '修改成功',
    changes: [
      {
        field: '客户名称',
        before: '泰州市红羽禽业有限公司',
        after: '郑州市红羽禽业有限公司',
      },
      {
        field: '交易状态',
        before: '意向客户',
        after: '成交客户',
      },
      {
        field: '手机号',
        before: '1802949900',
        after: '8999999999',
      },
    ],
  },
  {
    id: '2',
    companyName: '上海浦耀贸易有限公司',
    time: '2026-05-30 14:00',
    operatedAt: '2026-05-30 14:00:00',
    success: false,
    result: '客户已存在，修改失败',
    changes: [
      {
        field: '客户名称',
        before: '上海浦耀贸易有限公司',
        after: '上海浦耀贸易有限公司',
      },
      {
        field: '手机号',
        before: '18019999090',
        after: '18019249090',
      },
    ],
  },
];

function getDuplicateCheckRecord(id) {
  return duplicateCheckRecords.find((record) => record.id === id);
}

function getAddCustomerRecord(id) {
  return addCustomerRecords.find((record) => record.id === id);
}

function getInvalidCustomerRecord(id) {
  return invalidCustomerRecords.find((record) => record.id === id);
}

function getEditCustomerRecord(id) {
  return editCustomerRecords.find((record) => record.id === id);
}

function renderChangeItem(change) {
  return `
    <div class="change-item">
      <p class="change-label">${change.field}</p>
      <p class="change-row">修改前：${change.before}</p>
      <p class="change-row">修改后：${change.after}</p>
    </div>
  `;
}

function renderEditDetailContent(record) {
  if (!record) {
    return '<p class="record-empty">记录不存在</p>';
  }

  const resultClass = record.success ? 'result-success' : 'result-fail';

  return `
    <div class="detail-card">
      <div class="edit-changes">
        <h3 class="result-section-title">修改内容</h3>
        ${record.changes.map(renderChangeItem).join('')}
      </div>
      <div class="detail-meta">
        <div class="result-row">
          <span class="result-label">结果</span>
          <span class="result-value ${resultClass}">${record.result}</span>
        </div>
        ${renderResultRow('操作时间', record.operatedAt)}
        ${renderResultRow('操作类型', '客户修改')}
      </div>
    </div>
  `;
}

function renderInvalidDetailContent(record) {
  if (!record) {
    return '<p class="record-empty">记录不存在</p>';
  }

  return `
    <div class="detail-card">
      <div class="result-info">
        ${renderResultRow('客户名称', record.name)}
        ${renderResultRow('客户手机号', record.phone)}
        ${renderResultRow('客户地址', record.address)}
        ${renderResultRow('客户ID', record.customerId)}
      </div>
      <div class="detail-meta">
        ${renderResultRow('结果', record.result)}
        ${renderResultRow('通知', record.notification)}
        ${renderResultRow('操作时间', record.operatedAt)}
        ${renderResultRow('操作类型', '客户失效')}
      </div>
    </div>
  `;
}

function renderAddDetailContent(record) {
  if (!record) {
    return '<p class="record-empty">记录不存在</p>';
  }

  const resultText = record.success ? '添加成功' : record.failReason;
  const resultClass = record.success ? 'result-success' : 'result-fail';

  return `
    <div class="detail-card">
      <div class="result-info">
        ${renderResultRow('客户名称', record.name)}
        ${renderResultRow('客户手机号', record.phone)}
        ${renderResultRow('交易状态', record.status)}
        ${renderResultRow('养殖规模', record.scale)}
        ${renderResultRow('地址区域', record.region)}
        ${renderResultRow('详细地址', record.address)}
      </div>
      <div class="result-section">
        <h3 class="result-section-title">结果</h3>
        <p class="${resultClass}">${resultText}</p>
      </div>
      <div class="detail-meta">
        ${renderResultRow('操作时间', record.operatedAt)}
        ${renderResultRow('操作人', record.operator)}
        ${renderResultRow('操作类型', '添加客户')}
      </div>
    </div>
  `;
}

function renderDuplicateDetailContent(record) {
  if (!record) {
    return '<p class="record-empty">记录不存在</p>';
  }

  const resultContent = record.hasDuplicate
    ? `<ol class="result-list">${record.reasons.map((reason) => `<li>${reason}</li>`).join('')}</ol>`
    : '<p class="result-success">无相似客户备案</p>';

  return `
    <div class="detail-card">
      <div class="result-info">
        ${renderResultRow('客户名称', record.name)}
        ${renderResultRow('客户手机号', record.phone)}
        ${renderResultRow('客户地址', record.address || '未选择')}
      </div>
      <div class="result-section">
        <h3 class="result-section-title">结果</h3>
        ${resultContent}
      </div>
      <div class="detail-meta">
        ${renderResultRow('操作时间', record.operatedAt)}
        ${renderResultRow('操作人', record.operator)}
        ${renderResultRow('操作类型', '客户查重')}
      </div>
    </div>
  `;
}

function initOperationRecordsPage() {
  const duplicateList = document.getElementById('duplicateRecordList');
  duplicateList.innerHTML = duplicateCheckRecords.map((record) => `
    <li>
      <a href="duplicate-check-detail.html?id=${record.id}" class="record-card">
        <p class="record-card-company">${record.companyName}</p>
        <p class="record-card-time">${record.time}</p>
        <p class="record-card-result">结果：${record.hasDuplicate ? '有相似客户备案' : '无相似客户备案'}</p>
      </a>
    </li>
  `).join('');

  const addList = document.getElementById('addRecordList');
  addList.innerHTML = addCustomerRecords.map((record) => `
    <li>
      <a href="add-record-detail.html?id=${record.id}" class="record-card">
        <p class="record-card-company">${record.companyName}</p>
        <p class="record-card-time">${record.time}</p>
        <p class="record-card-result">结果：${record.success ? '添加成功' : record.failReason}</p>
      </a>
    </li>
  `).join('');

  const invalidList = document.getElementById('invalidRecordList');
  invalidList.innerHTML = invalidCustomerRecords.map((record) => `
    <li>
      <a href="invalid-record-detail.html?id=${record.id}" class="record-card">
        <p class="record-card-company">${record.companyName}</p>
        <p class="record-card-time">${record.time}</p>
        <p class="record-card-result">结果：${record.result}</p>
      </a>
    </li>
  `).join('');

  const editList = document.getElementById('editRecordList');
  editList.innerHTML = editCustomerRecords.map((record) => `
    <li>
      <a href="edit-record-detail.html?id=${record.id}" class="record-card">
        <p class="record-card-company">${record.companyName}</p>
        <p class="record-card-time">${record.time}</p>
        <p class="record-card-result">结果：${record.result}</p>
      </a>
    </li>
  `).join('');

  const params = new URLSearchParams(window.location.search);
  const activeTab = params.get('tab') || 'duplicate';

  document.querySelectorAll('.record-tab').forEach((tab) => {
    if (tab.dataset.tab === activeTab) {
      tab.classList.add('active');
      document.getElementById(`panel-${activeTab}`).classList.add('active');
    } else {
      tab.classList.remove('active');
      document.getElementById(`panel-${tab.dataset.tab}`).classList.remove('active');
    }

    tab.addEventListener('click', () => {
      document.querySelectorAll('.record-tab').forEach((item) => item.classList.remove('active'));
      document.querySelectorAll('.record-panel').forEach((panel) => panel.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById(`panel-${tab.dataset.tab}`).classList.add('active');
    });
  });

  if (activeTab !== 'duplicate') {
    document.querySelector('.record-tab[data-tab="duplicate"]').classList.remove('active');
    document.getElementById('panel-duplicate').classList.remove('active');
  }
}

function initDuplicateCheckDetailPage() {
  const params = new URLSearchParams(window.location.search);
  const record = getDuplicateCheckRecord(params.get('id') || '1');
  document.getElementById('duplicateDetailContent').innerHTML = renderDuplicateDetailContent(record);
}

function initAddRecordDetailPage() {
  const params = new URLSearchParams(window.location.search);
  const record = getAddCustomerRecord(params.get('id') || '1');
  document.getElementById('addDetailContent').innerHTML = renderAddDetailContent(record);
}

function initInvalidRecordDetailPage() {
  const params = new URLSearchParams(window.location.search);
  const record = getInvalidCustomerRecord(params.get('id') || '1');
  document.getElementById('invalidDetailContent').innerHTML = renderInvalidDetailContent(record);
}

function initEditRecordDetailPage() {
  const params = new URLSearchParams(window.location.search);
  const record = getEditCustomerRecord(params.get('id') || '1');
  document.getElementById('editDetailContent').innerHTML = renderEditDetailContent(record);
}

if (document.getElementById('duplicateRecordList')) {
  initOperationRecordsPage();
}

if (document.getElementById('duplicateDetailContent')) {
  initDuplicateCheckDetailPage();
}

if (document.getElementById('addDetailContent')) {
  initAddRecordDetailPage();
}

if (document.getElementById('invalidDetailContent')) {
  initInvalidRecordDetailPage();
}

if (document.getElementById('editDetailContent')) {
  initEditRecordDetailPage();
}

function initEmployeeInfoPage() {
  initCustomSelects();

  document.getElementById('sendCodeBtn').addEventListener('click', () => {
    showToast('验证码已发送', 1500);
  });

  document.getElementById('employeeForm').addEventListener('submit', (e) => {
    e.preventDefault();
    window.location.href = 'index.html';
  });
}

if (document.getElementById('employeeForm')) {
  initEmployeeInfoPage();
}

// 编辑模式功能
function initEditMode() {
  const editModeToggle = document.getElementById('editModeToggle');
  const editIndicator = document.getElementById('editIndicator');
  const editToolbar = document.getElementById('editToolbar');
  const app = document.querySelector('.app');
  const btnSave = document.getElementById('btnSave');
  const btnExport = document.getElementById('btnExport');
  const btnCancel = document.getElementById('btnCancelEdit');
  const btnAddField = document.getElementById('btnAddField');
  
  // 存储原始内容
  const originalContent = new Map();
  const editedElements = new Set();
  const addedFields = [];
  
  // 开启编辑模式
  editModeToggle.addEventListener('change', function() {
    if (this.checked) {
      enterEditMode();
    } else {
      exitEditMode(false);
    }
  });
  
  // 保存修改
  btnSave.addEventListener('click', function() {
    saveChanges();
  });
  
  // 导出HTML
  btnExport.addEventListener('click', function() {
    exportHTML();
  });
  
  // 取消编辑
  btnCancel.addEventListener('click', function() {
    exitEditMode(true);
  });
  
  // 添加字段
  btnAddField.addEventListener('click', function() {
    showAddFieldModal();
  });
  
  // 添加字段弹窗相关元素
  const addFieldOverlay = document.getElementById('addFieldOverlay');
  const addFieldClose = document.getElementById('addFieldClose');
  const addFieldCancel = document.getElementById('addFieldCancel');
  const addFieldConfirm = document.getElementById('addFieldConfirm');
  const fieldType = document.getElementById('fieldType');
  const selectOptionsContainer = document.getElementById('selectOptionsContainer');
  
  // 字段类型变化时显示/隐藏下拉选项
  fieldType.addEventListener('change', function() {
    selectOptionsContainer.hidden = this.value !== 'select';
  });
  
  addFieldClose.addEventListener('click', hideAddFieldModal);
  addFieldCancel.addEventListener('click', hideAddFieldModal);
  
  addFieldConfirm.addEventListener('click', function() {
    addNewField();
  });
  
  // 显示添加字段弹窗
  function showAddFieldModal() {
    addFieldOverlay.hidden = false;
  }
  
  // 隐藏添加字段弹窗
  function hideAddFieldModal() {
    addFieldOverlay.hidden = true;
    // 重置表单
    document.getElementById('fieldLabel').value = '';
    document.getElementById('fieldType').value = 'text';
    document.getElementById('fieldRequired').checked = false;
    document.getElementById('selectOptions').value = '';
    selectOptionsContainer.hidden = true;
  }
  
  // 添加新字段
  function addNewField() {
    const label = document.getElementById('fieldLabel').value.trim();
    const type = document.getElementById('fieldType').value;
    const required = document.getElementById('fieldRequired').checked;
    const options = document.getElementById('selectOptions').value.trim();
    
    if (!label) {
      showToast('请输入字段标签', 2000);
      return;
    }
    
    const form = document.getElementById('duplicateForm');
    const formActions = form.querySelector('.form-actions');
    
    // 创建字段HTML
    let fieldHTML = createFieldHTML(label, type, required, options);
    
    // 插入到提交按钮之前
    formActions.insertAdjacentHTML('beforebegin', fieldHTML);
    
    // 记录添加的字段
    const newField = formActions.previousElementSibling;
    addedFields.push(newField);
    
    // 添加删除按钮事件
    const deleteBtn = newField.querySelector('.field-delete-btn');
    deleteBtn.addEventListener('click', function() {
      newField.remove();
      const index = addedFields.indexOf(newField);
      if (index > -1) {
        addedFields.splice(index, 1);
      }
    });
    
    hideAddFieldModal();
    showToast(`字段"${label}"已添加`, 2000);
  }
  
  // 创建字段HTML
  function createFieldHTML(label, type, required, options) {
    const requiredMark = required ? '<span class="required">*</span>' : '';
    const fieldId = 'customField_' + Date.now();
    
    let fieldContent = '';
    
    switch(type) {
      case 'text':
        fieldContent = `<input type="text" class="form-input" name="${fieldId}" placeholder="请填写" />`;
        break;
      case 'tel':
        fieldContent = `<input type="tel" class="form-input" name="${fieldId}" placeholder="请填写手机号" />`;
        break;
      case 'select':
        const optionList = options.split(',').map(opt => 
          `<option value="${opt.trim()}">${opt.trim()}</option>`
        ).join('');
        fieldContent = `
          <div class="custom-select" data-target="${fieldId}">
            <button type="button" class="form-field form-field-picker custom-select-trigger">
              <span class="form-value form-value-placeholder" id="${fieldId}Display">请选择</span>
              <svg class="field-icon field-icon-right" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <polyline points="6 9 12 15 18 9"/>
              </svg>
            </button>
            <input type="hidden" name="${fieldId}" id="${fieldId}" value="" />
          </div>
        `;
        break;
      case 'textarea':
        fieldContent = `<textarea class="form-input" name="${fieldId}" rows="3" placeholder="请填写"></textarea>`;
        break;
      case 'location':
        fieldContent = `
          <button type="button" class="form-field form-field-location">
            <svg class="location-pin" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/>
              <circle cx="12" cy="10" r="3"/>
            </svg>
            <span class="location-text form-value-placeholder">选择定位地址</span>
            <svg class="field-icon field-icon-right" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polyline points="9 18 15 12 9 6"/>
            </svg>
          </button>
        `;
        break;
    }
    
    return `
      <div class="form-item custom-field" style="position: relative;">
        <label class="form-label">${requiredMark}${label}</label>
        <div class="form-control">
          <div class="form-field">
            ${fieldContent}
          </div>
        </div>
        <button type="button" class="field-delete-btn" title="删除此字段">×</button>
      </div>
    `;
  }
  
  // 进入编辑模式
  function enterEditMode() {
    app.classList.add('edit-mode-active');
    editIndicator.hidden = false;
    editToolbar.classList.add('show');
    
    // 为所有非输入框的文字元素添加可编辑属性
    const textElements = document.querySelectorAll('h1, h2, h3, h4, h5, h6, p, span, label, .form-label, .form-value, .nav-text, .page-title');
    
    textElements.forEach(function(el) {
      // 跳过输入框、按钮和已有可编辑属性的元素
      if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || 
          el.tagName === 'BUTTON' || el.isContentEditable) {
        return;
      }
      
      // 存储原始内容
      originalContent.set(el, el.textContent);
      el.classList.add('editable-text');
      
      // 添加点击事件
      el.addEventListener('click', handleTextClick);
    });
  }
  
  // 处理文字点击
  function handleTextClick(e) {
    if (!editModeToggle.checked) return;
    
    const el = e.target;
    if (!el.classList.contains('editable-text')) return;
    
    const originalText = el.textContent;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'edit-input';
    input.value = originalText;
    input.style.width = (el.offsetWidth + 20) + 'px';
    input.style.fontSize = window.getComputedStyle(el).fontSize;
    
    // 替换为输入框
    el.innerHTML = '';
    el.appendChild(input);
    input.focus();
    input.select();
    
    // 失去焦点时保存
    input.addEventListener('blur', function() {
      const newValue = input.value.trim();
      if (newValue !== originalText) {
        el.textContent = newValue;
        editedElements.add(el);
      } else {
        el.textContent = originalText;
      }
    });
    
    // 回车保存
    input.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') {
        input.blur();
      }
    });
  }
  
  // 退出编辑模式
  function exitEditMode(reset) {
    app.classList.remove('edit-mode-active');
    editIndicator.hidden = true;
    editToolbar.classList.remove('show');
    editModeToggle.checked = false;
    
    // 重置所有修改
    if (reset) {
      originalContent.forEach(function(text, el) {
        el.textContent = text;
      });
      editedElements.clear();
      // 删除所有添加的字段
      addedFields.forEach(function(field) {
        field.remove();
      });
      addedFields.length = 0;
    }
    
    // 移除可编辑属性和事件
    document.querySelectorAll('.editable-text').forEach(function(el) {
      el.classList.remove('editable-text');
      el.removeEventListener('click', handleTextClick);
    });
  }
  
  // 保存修改（保存到 localStorage）
  function saveChanges() {
    if (editedElements.size === 0) {
      showToast('没有修改内容', 2000);
      return;
    }
    
    // 收集所有修改
    const changes = [];
    editedElements.forEach(function(el) {
      const original = originalContent.get(el);
      const current = el.textContent;
      if (original !== current) {
        changes.push({
          element: el.tagName + '.' + (el.className || 'no-class'),
          original: original,
          current: current
        });
      }
    });
    
    // 保存到 localStorage
    localStorage.setItem('pageChanges', JSON.stringify(changes));
    
    showToast('修改已保存', 2000);
    exitEditMode(false);
  }
  
  // 导出HTML
  function exportHTML() {
    // 获取当前页面的HTML
    let html = document.documentElement.outerHTML;
    
    // 创建下载链接
    const blob = new Blob([html], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'duplicate-check-edited.html';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    showToast('HTML已导出', 2000);
    exitEditMode(false);
  }
}

// 页面加载完成后初始化编辑模式
if (document.getElementById('editModeToggle')) {
  initEditMode();
}
