// ========== 設定 ==========
const API_BASE = 'https://script.google.com/macros/s/AKfycbzYLGWStQhyLRFot7AyGNL8QQCCwqCCS1NcLREzp8gQ3Z3ySmR_py885q9_Th4HmLg/exec';

// ========== 状態 ==========
let currentUser = null;
let cart = [];
let categories = [];
let exchangeRate = 0;
let currentProduct = null;

// ========== 初期化 ==========
document.addEventListener('DOMContentLoaded', () => {
  const params = new URLSearchParams(window.location.search);
  const key = params.get('key');

  if (key) {
    document.getElementById('key-input').value = key;
    authenticate();
  }
});

// ========== 認証 ==========
async function authenticate() {
  const key = document.getElementById('key-input').value.trim();
  if (!key) {
    showError('auth-error', 'Please enter your access key');
    return;
  }

  try {
    const userRes = await apiGet('user', { key });
    if (userRes.error) throw new Error(userRes.error);
    currentUser = userRes;

    const catRes = await apiGet('categories');
    categories = catRes.categories || [];

    const rateRes = await apiGet('rate', { currency: currentUser.currency });
    exchangeRate = rateRes.rate || 150;

    const cartRes = await apiGet('cart', { key });
    cart = cartRes.items || [];

    showMainScreen();

  } catch (err) {
    showError('auth-error', err.message);
  }
}

function showMainScreen() {
  document.getElementById('auth-screen').classList.add('hidden');
  document.getElementById('main-screen').classList.remove('hidden');

  document.getElementById('user-name').textContent = currentUser.name;
  document.getElementById('user-currency').textContent = currentUser.currency;

  const select = document.getElementById('category-select');
  select.innerHTML = categories.map(c =>
    `<option value="${c.name}">${c.name}</option>`
  ).join('');

  document.getElementById('rate-currency').textContent = currentUser.currency;
  document.getElementById('rate-value').textContent = exchangeRate.toFixed(2);
  document.getElementById('currency-label').textContent = currentUser.currency;

  renderCart();
  loadOrders();
}

// ========== 商品取得 ==========
async function fetchProduct() {
  const url = document.getElementById('product-url').value.trim();
  if (!url) {
    showError('fetch-error', 'Please enter a URL');
    return;
  }

  const btn = document.getElementById('fetch-btn');
  btn.disabled = true;
  btn.textContent = 'Loading...';
  clearError('fetch-error');

  try {
    const res = await apiGet('product', { url });
    if (res.error) throw new Error(res.error);

    currentProduct = res;

    document.getElementById('preview-name').textContent = res.nameEN || res.nameJP || 'Unknown';
    document.getElementById('preview-price').textContent = res.price.toLocaleString();
    document.getElementById('preview-category').textContent = res.category || 'Unknown';

    // カテゴリプルダウンを自動判定結果に設定
    const select = document.getElementById('category-select');
    if (res.category) {
      select.value = res.category;
    }

    const stockEl = document.getElementById('preview-stock');
    if (res.stock === 'IN') {
      stockEl.textContent = '✓ In Stock';
      stockEl.className = 'product-stock in-stock';
    } else if (res.stock === 'OUT') {
      stockEl.textContent = '✗ Out of Stock';
      stockEl.className = 'product-stock out-of-stock';
    } else {
      stockEl.textContent = '? Stock Unknown';
      stockEl.className = 'product-stock';
    }

    document.getElementById('product-preview').classList.remove('hidden');

  } catch (err) {
    showError('fetch-error', err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Fetch';
  }
}

// ========== カート操作 ==========
async function addToCart() {
  if (!currentProduct) return;

  // ユーザーが変更した可能性があるのでプルダウンから取得
  const category = document.getElementById('category-select').value;

  const item = {
    url: currentProduct.url,
    productId: currentProduct.productId,
    nameEN: currentProduct.nameEN || currentProduct.nameJP,
    price: currentProduct.price,
    category: category
  };

  try {
    await apiPost({
      action: 'addToCart',
      userKey: currentUser.key,
      item: item
    });

    cart.push(item);
    renderCart();

    document.getElementById('product-url').value = '';
    document.getElementById('product-preview').classList.add('hidden');
    currentProduct = null;

  } catch (err) {
    alert('Error adding to cart: ' + err.message);
  }
}

async function removeFromCart(index) {
  try {
    await apiPost({
      action: 'removeFromCart',
      userKey: currentUser.key,
      itemIndex: index
    });

    cart.splice(index, 1);
    renderCart();

  } catch (err) {
    alert('Error removing item: ' + err.message);
  }
}

function renderCart() {
  const container = document.getElementById('cart-items');
  const summary = document.getElementById('cart-summary');

  document.getElementById('cart-count').textContent = `(${cart.length})`;

  if (cart.length === 0) {
    container.innerHTML = '<p style="color: #666; text-align: center; padding: 20px;">Cart is empty</p>';
    summary.classList.add('hidden');
    return;
  }

  container.innerHTML = cart.map((item, i) => `
    <div class="cart-item">
      <div class="cart-item-info">
        <div class="cart-item-name">${item.nameEN}</div>
        <div class="cart-item-details">${item.category} | ID: ${item.productId}</div>
      </div>
      <div class="cart-item-price">¥${item.price.toLocaleString()}</div>
      <button class="remove-btn" onclick="removeFromCart(${i})">Remove</button>
    </div>
  `).join('');

  const subtotal = cart.reduce((sum, item) => sum + (item.price * currentUser.margin), 0);
  const shipping = calculateShipping();
  const totalJPY = subtotal + shipping;
  const totalFX = totalJPY / exchangeRate;

  document.getElementById('subtotal-jpy').textContent = Math.round(subtotal).toLocaleString();
  document.getElementById('shipping-jpy').textContent = shipping.toLocaleString();
  document.getElementById('total-jpy').textContent = Math.round(totalJPY).toLocaleString();
  document.getElementById('total-fx').textContent = totalFX.toFixed(2);

  summary.classList.remove('hidden');
}

function calculateShipping() {
  const counts = {};

  cart.forEach(item => {
    const cat = categories.find(c => c.name === item.category);
    if (!cat) return;

    const groupKey = cat.groupWith || cat.name;
    if (!counts[groupKey]) counts[groupKey] = 0;
    counts[groupKey] += (cat.multiplier || 1);
  });

  let total = 0;
  Object.keys(counts).forEach(groupKey => {
    const cat = categories.find(c => c.name === groupKey);
    if (!cat) return;

    const boxes = Math.ceil(counts[groupKey] / cat.perBox);
    total += boxes * cat.boxPrice;
  });

  return total;
}

// ========== 注文 ==========
async function submitOrder() {
  if (cart.length === 0) return;

  if (!confirm('Submit this order?')) return;

  try {
    const res = await apiPost({
      action: 'submitOrder',
      userKey: currentUser.key
    });

    if (res.error) throw new Error(res.error);

    alert(`Order submitted!\n\nOrder ID: ${res.orderId}\nTotal: ${res.totalFX} ${res.currency}\n\nPlease transfer to Wise account.`);

    cart = [];
    renderCart();
    loadOrders();

  } catch (err) {
    alert('Error submitting order: ' + err.message);
  }
}

async function loadOrders() {
  try {
    const res = await apiGet('orders', { key: currentUser.key });
    const orders = res.orders || [];

    const container = document.getElementById('orders-list');

    if (orders.length === 0) {
      container.innerHTML = '<p style="color: #666;">No orders yet</p>';
      return;
    }

    container.innerHTML = orders.reverse().map(order => `
      <div class="order-item">
        <div class="order-header">
          <span class="order-id">${order.orderId}</span>
          <span class="order-status ${order.status}">${order.status}</span>
        </div>
        <div class="order-date">${new Date(order.orderDate).toLocaleDateString()}</div>
        <div class="order-total">${order.totalFX}</div>
      </div>
    `).join('');

  } catch (err) {
    console.error('Error loading orders:', err);
  }
}

// ========== API ヘルパー ==========
async function apiGet(action, params = {}) {
  const url = new URL(API_BASE);
  url.searchParams.set('action', action);
  Object.keys(params).forEach(k => url.searchParams.set(k, params[k]));

  const res = await fetch(url.toString());
  return res.json();
}

async function apiPost(data) {
  const res = await fetch(API_BASE, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data)
  });
  return res.json();
}

// ========== ユーティリティ ==========
function showError(id, msg) {
  document.getElementById(id).textContent = msg;
}

function clearError(id) {
  document.getElementById(id).textContent = '';
}
