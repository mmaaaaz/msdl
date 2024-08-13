const config = {
    langsUrl: "https://www.microsoft.com/en-us/api/controls/contentinclude/html?pageId=cd06bda8-ff9c-4a6e-912a-b92a21f42526&host=www.microsoft.com&segments=software-download%2cwindows11&query=&action=getskuinformationbyproductedition&sdVersion=2",
    downUrl: "https://www.microsoft.com/en-us/api/controls/contentinclude/html?pageId=cfa9e580-a81e-4a4b-a846-7b21bf4e2e5b&host=www.microsoft.com&segments=software-download%2Cwindows11&query=&action=GetProductDownloadLinksBySku&sdVersion=2",
    sessionUrl: "https://vlscppe.microsoft.com/fp/tags?org_id=y6jn8c31&session_id=",
    apiUrl: "https://api.gravesoft.dev/msdl/",
    sharedSessionGUID: "47cbc254-4a79-4be6-9866-9c625eb20911"
};

const elements = {
    sessionId: document.getElementById('msdl-session-id'),
    msContent: document.getElementById('msdl-ms-content'),
    pleaseWait: document.getElementById('msdl-please-wait'),
    processingError: document.getElementById('msdl-processing-error'),
    productsList: document.getElementById('products-list'),
    backToProductsDiv: document.getElementById('back-to-products')
};

let state = {
    availableProducts: {},
    sharedSession: false,
    shouldUseSharedSession: true,
    skuId: null
};

// 4,800 ops/s
const uuidv4 = () => ([1e7] + -1e3 + -4e3 + -8e3 + -1e11).replace(/[018]/g, c =>
    (c ^ crypto.getRandomValues(new Uint8Array(1))[0] & 15 >> c / 4).toString(16)
);

// 32,100 ops/s
const uuidv4Performant = () => crypto.getRandomValues(new Uint8Array(16)).reduce((a, b, i) => a + (b & (i === 6 ? 0x0f | 0x40 : i === 8 ? 0x3f | 0x80 : 0xff)).toString(16).padStart(2, '0') + (i === 3 || i === 5 || i === 7 || i === 9 ? '-' : ''), '');

const updateVars = () => {
    const id = document.getElementById('product-languages').value;
    document.getElementById('submit-sku').disabled = !id;
    return id ? JSON.parse(id)['id'] : null;
};

const updateContent = (content, response) => {
    content.innerHTML = response;
    if (document.getElementById('errorModalMessage')) {
        elements.processingError.style.display = "block";
        return false;
    }
    return true;
};

const fetchAndHandle = async (url, handler) => {
    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error('Network response was not ok');
        handler(await response.text());
    } catch (error) {
        console.error('Fetch error:', error);
        elements.processingError.style.display = "block";
    }
};

const onLanguageFetch = async (response) => {
    if (elements.pleaseWait.style.display !== "block") return;
    elements.pleaseWait.style.display = "none";
    elements.msContent.style.display = "block";
    if (!updateContent(elements.msContent, response)) return;
    document.getElementById('submit-sku').onclick = getDownload;
    document.getElementById('product-languages').onchange = updateVars;
    updateVars();
};

const onDownloadsFetch = async (response) => {
    if (elements.pleaseWait.style.display !== "block") return;
    elements.msContent.style.display = "block";
    if (updateContent(elements.msContent, response)) {
        elements.pleaseWait.style.display = "none";
        if (!state.sharedSession) {
            await fetch(config.sessionUrl + config.sharedSessionGUID);
            await fetch(config.sessionUrl + "de40cb69-50a5-415e-a0e8-3cf1eed1b7cd");
            await fetch(config.apiUrl + 'add_session?session_id=' + elements.sessionId.value);
        }
    } else if (!state.sharedSession && state.shouldUseSharedSession) {
        useSharedSession();
    } else {
        getFromServer();
    }
};

const getFromServer = async () => {
    elements.processingError.style.display = "none";
    const url = `${config.apiUrl}proxy?product_id=${window.location.hash.substring(1)}&sku_id=${state.skuId}`;
    await fetchAndHandle(url, (text) => {
        elements.pleaseWait.style.display = "none";
        if (!response.ok) {
            elements.processingError.style.display = "block";
            alert(JSON.parse(text)["Error"]);
            return;
        }
        elements.msContent.innerHTML = text;
    });
};

const getLanguages = async (productId) => {
    const url = `${config.langsUrl}&productEditionId=${productId}&sessionId=${state.sharedSession ? config.sharedSessionGUID : elements.sessionId.value}`;
    await fetchAndHandle(url, onLanguageFetch);
};

const getDownload = async () => {
    elements.msContent.style.display = "none";
    elements.pleaseWait.style.display = "block";
    state.skuId = state.skuId || updateVars();
    const url = `${config.downUrl}&skuId=${state.skuId}&sessionId=${state.sharedSession ? config.sharedSessionGUID : elements.sessionId.value}`;
    await fetchAndHandle(url, onDownloadsFetch);
};

const backToProducts = () => {
    elements.backToProductsDiv.style.display = 'none';
    elements.productsList.style.display = 'block';
    elements.msContent.style.display = 'none';
    elements.pleaseWait.style.display = 'none';
    elements.processingError.style.display = 'none';
    window.location.hash = "";
    state.skuId = null;
};

const useSharedSession = () => {
    state.sharedSession = true;
    retryDownload();
};

const retryDownload = async () => {
    elements.pleaseWait.style.display = "block";
    elements.processingError.style.display = 'none';
    const url = `${config.langsUrl}&productEditionId=${window.location.hash.substring(1)}&sessionId=${config.sharedSessionGUID}`;
    await fetchAndHandle(url, getDownload);
};

const prepareDownload = async (id) => {
    elements.productsList.style.display = 'none';
    elements.backToProductsDiv.style.display = 'block';
    elements.pleaseWait.style.display = "block";
    try {
        await fetch(config.sessionUrl + elements.sessionId.value);
    } catch {
        getLanguages(id);
    }
};

const addTableElement = (table, value, data) => {
    const a = document.createElement('a');
    a.href = "#" + value;
    a.onclick = () => prepareDownload(value);
    a.appendChild(document.createTextNode(data[value]));
    const tr = table.insertRow();
    tr.insertCell().appendChild(a);
    tr.insertCell().appendChild(document.createTextNode(value));
};

const createTable = (data, search) => {
    const table = document.getElementById('products-table-body');
    const regex = new RegExp(search, 'ig');
    table.innerHTML = "";
    Object.keys(data).forEach(value => {
        if (data[value].match(regex)) addTableElement(table, value, data);
    });
};

const updateResults = () => {
    createTable(state.availableProducts, document.getElementById('search-products').value);
};

const setSearch = (query) => {
    document.getElementById('search-products').value = query;
    updateResults();
};

const checkHash = () => {
    const hash = window.location.hash;
    if (hash) prepareDownload(hash.substring(1));
};

const preparePage = (resp) => {
    state.availableProducts = JSON.parse(resp);
    if (!state.availableProducts) {
        elements.pleaseWait.style.display = 'none';
        elements.processingError.style.display = 'block';
        return;
    }
    elements.pleaseWait.style.display = 'none';
    elements.productsList.style.display = 'block';
    updateResults();
    checkHash();
};

const initialize = async () => {
    elements.sessionId.value = uuidv4Performant();
    elements.pleaseWait.style.display = 'block';
    await fetchAndHandle('data/products.json', preparePage);
    try {
        const response = await fetch(`${config.apiUrl}use_shared_session`);
        if (!response.ok) state.shouldUseSharedSession = false;
    } catch {
        state.shouldUseSharedSession = false;
    }
};

initialize();