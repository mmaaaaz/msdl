// ==== Constants & DOM Refs ====
const API_BASE_URL = "https://api.gravesoft.dev/msdl/";
const CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours

const ELEMENTS = {
  sessionId: document.getElementById("msdl-session-id"),
  msContent: document.getElementById("msdl-ms-content"),
  pleaseWait: document.getElementById("msdl-please-wait"),
  processingError: document.getElementById("msdl-processing-error"),
  productsList: document.getElementById("products-list"),
  backToProductsDiv: document.getElementById("back-to-products"),
  productsTableBody: document.getElementById("products-table-body"),
  searchInput: document.getElementById("search-products"),
};

// ==== Smart Cache ====

class SmartCache {
  constructor(prefix = "msdl", ttl = CACHE_TTL_MS) {
    this.prefix = prefix;
    this.ttl = ttl;
    this.memory = {};
  }

  _now() {
    return Date.now();
  }

  _getKey(key) {
    return `${this.prefix}:${key}`;
  }

  _isValid(item) {
    return item && item.expiry > this._now();
  }

  get(key) {
    const fullKey = this._getKey(key);

    // In-memory
    if (this.memory[key] && this._isValid(this.memory[key])) {
      return this.memory[key].value;
    }

    // From localStorage
    try {
      const raw = localStorage.getItem(fullKey);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!this._isValid(parsed)) {
        localStorage.removeItem(fullKey);
        return null;
      }
      this.memory[key] = parsed;
      return parsed.value;
    } catch {
      return null;
    }
  }

  set(key, value) {
    const fullKey = this._getKey(key);
    const item = {
      value,
      expiry: this._now() + this.ttl,
    };
    this.memory[key] = item;
    try {
      localStorage.setItem(fullKey, JSON.stringify(item));
    } catch {
      // Ignore storage errors
    }
  }
}

const smartCache = new SmartCache();

// ==== State ====
let availableProducts = {};
let skuId = null;

// ==== Utility Functions ====

const uuidv4 = () =>
  ([1e7] + -1e3 + -4e3 + -8e3 + -1e11).replace(/[018]/g, (c) =>
    (c ^ (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (c / 4)))).toString(16)
  );

const setVisibility = (el, display) => {
  if (el) el.style.display = display;
};

const parseJSONSafe = (str) => {
  try {
    return JSON.parse(str);
  } catch {
    return null;
  }
};

// ==== Language Selection UI ====

const langJsonStrToHTML = (jsonStr) => {
  const json = parseJSONSafe(jsonStr);
  if (!json) return document.createTextNode("Failed to load languages.");

  const container = document.createElement("div");

  container.innerHTML = `
    <h2>Select the product language</h2>
    <p>
      You'll need to choose the same language when you install Windows. To see what language you're currently using,
      go to <strong>Time and language</strong> in PC settings or <strong>Region</strong> in Control Panel.
    </p>
  `;

  const select = document.createElement("select");
  select.id = "product-languages";

  select.innerHTML =
    `<option value="" selected>Choose one</option>` +
    json.Skus.map((sku) => `<option value='${JSON.stringify({ id: sku.Id })}'>${sku.LocalizedLanguage}</option>`).join(
      ""
    );

  const button = document.createElement("button");
  button.id = "submit-sku";
  button.textContent = "Submit";
  button.disabled = true;

  container.append(select, button);
  return container;
};

// ==== Event Handlers ====

const updateSkuId = () => {
  const prodLang = document.getElementById("product-languages");
  const submitBtn = document.getElementById("submit-sku");
  const selected = prodLang?.value ?? "";

  if (!selected) {
    if (submitBtn) submitBtn.disabled = true;
    return null;
  }

  if (submitBtn) submitBtn.disabled = false;
  return parseJSONSafe(selected)?.id ?? null;
};

const onLanguageResponse = (responseText) => {
  setVisibility(ELEMENTS.pleaseWait, "none");
  setVisibility(ELEMENTS.msContent, "block");
  ELEMENTS.msContent.innerHTML = "";

  const langNode = langJsonStrToHTML(responseText);
  ELEMENTS.msContent.appendChild(langNode);

  document.getElementById("submit-sku")?.addEventListener("click", getDownload, { once: true });
  document.getElementById("product-languages")?.addEventListener("change", updateSkuId);
  updateSkuId();
};

const onDownloadResponse = (responseText) => {
  const response = parseJSONSafe(responseText);
  setVisibility(ELEMENTS.pleaseWait, "none");
  setVisibility(ELEMENTS.msContent, "block");
  ELEMENTS.msContent.innerHTML = "";

  if (!response?.ProductDownloadOptions?.length) {
    ELEMENTS.msContent.textContent = "No download options available.";
    return;
  }

  const header = document.createElement("h2");
  const { ProductDisplayName, LocalizedLanguage } = response.ProductDownloadOptions[0];
  header.textContent = `${ProductDisplayName} ${LocalizedLanguage}`;
  ELEMENTS.msContent.appendChild(header);

  response.ProductDownloadOptions.forEach(({ Uri }) => {
    const link = document.createElement("a");
    link.href = Uri;
    link.target = "_blank";
    link.textContent = Uri.split("?")[0].split("/").pop();
    ELEMENTS.msContent.appendChild(link);
    ELEMENTS.msContent.appendChild(document.createElement("br"));
  });
};

// ==== API Requests with Smart Cache ====

const getLanguages = async (productId) => {
  const cacheKey = `languages:${productId}`;
  const cached = smartCache.get(cacheKey);
  if (cached) return onLanguageResponse(cached);

  try {
    setVisibility(ELEMENTS.pleaseWait, "block");
    const res = await fetch(`${API_BASE_URL}skuinfo?product_id=${productId}`);
    const text = await res.text();
    if (!res.ok) throw new Error();
    smartCache.set(cacheKey, text);
    onLanguageResponse(text);
  } catch {
    setVisibility(ELEMENTS.processingError, "block");
  }
};

const getDownload = async () => {
  setVisibility(ELEMENTS.msContent, "none");
  setVisibility(ELEMENTS.pleaseWait, "block");

  skuId = updateSkuId();
  const productId = location.hash.substring(1);
  const cacheKey = `downloads:${productId}|${skuId}`;
  const cached = smartCache.get(cacheKey);

  if (cached) return onDownloadResponse(cached);

  try {
    const res = await fetch(`${API_BASE_URL}proxy?product_id=${productId}&sku_id=${skuId}`);
    const text = await res.text();
    if (!res.ok) throw new Error();
    smartCache.set(cacheKey, text);
    onDownloadResponse(text);
  } catch {
    setVisibility(ELEMENTS.processingError, "block");
  }
};

// ==== UI Actions ====

const backToProducts = () => {
  setVisibility(ELEMENTS.backToProductsDiv, "none");
  setVisibility(ELEMENTS.productsList, "block");
  setVisibility(ELEMENTS.msContent, "none");
  setVisibility(ELEMENTS.pleaseWait, "none");
  setVisibility(ELEMENTS.processingError, "none");
  location.hash = "";
  skuId = null;
};

const prepareDownload = (productId) => {
  setVisibility(ELEMENTS.productsList, "none");
  setVisibility(ELEMENTS.backToProductsDiv, "block");
  setVisibility(ELEMENTS.pleaseWait, "block");
  getLanguages(productId);
};

const addProductToTable = (productId, data) => {
  const row = ELEMENTS.productsTableBody.insertRow();
  const nameCell = row.insertCell();
  const idCell = row.insertCell();

  const link = document.createElement("a");
  link.href = `#${productId}`;
  link.textContent = data[productId];
  link.addEventListener("click", (e) => {
    e.preventDefault();
    prepareDownload(productId);
  });

  nameCell.appendChild(link);
  idCell.textContent = productId;
};

const createProductTable = (products, searchTerm) => {
  const regex = new RegExp(searchTerm, "i");
  ELEMENTS.productsTableBody.innerHTML = "";
  Object.entries(products).forEach(([id, name]) => {
    if (name.match(regex)) addProductToTable(id, products);
  });
};

const updateSearchResults = () => {
  createProductTable(availableProducts, ELEMENTS.searchInput.value);
};

const setSearch = (query) => {
  ELEMENTS.searchInput.value = ELEMENTS.searchInput.value === query ? "" : query;
  updateSearchResults();
};

const checkHashOnLoad = () => {
  const productId = location.hash?.substring(1);
  if (productId) prepareDownload(productId);
};

const preparePage = (dataText) => {
  const data = parseJSONSafe(dataText);
  if (!data) {
    setVisibility(ELEMENTS.pleaseWait, "none");
    setVisibility(ELEMENTS.processingError, "block");
    return;
  }
  availableProducts = data;
  setVisibility(ELEMENTS.pleaseWait, "none");
  setVisibility(ELEMENTS.productsList, "block");
  updateSearchResults();
  checkHashOnLoad();
};

const loadProducts = async () => {
  const cacheKey = "products";
  const cached = smartCache.get(cacheKey);
  if (cached) return preparePage(cached);

  try {
    const res = await fetch("data/products.json", { cache: "force-cache" });
    const text = await res.text();
    if (!res.ok) throw new Error();
    smartCache.set(cacheKey, text);
    preparePage(text);
  } catch {
    setVisibility(ELEMENTS.pleaseWait, "none");
    setVisibility(ELEMENTS.processingError, "block");
  }
};

// ==== Init ====
ELEMENTS.sessionId.value = uuidv4();
setVisibility(ELEMENTS.pleaseWait, "block");
loadProducts();
