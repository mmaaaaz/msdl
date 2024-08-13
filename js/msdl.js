const urls = {
  langsUrl:
    "https://www.microsoft.com/en-us/api/controls/contentinclude/html?pageId=cd06bda8-ff9c-4a6e-912a-b92a21f42526&host=www.microsoft.com&segments=software-download%2cwindows11&query=&action=getskuinformationbyproductedition&sdVersion=2",
  downUrl:
    "https://www.microsoft.com/en-us/api/controls/contentinclude/html?pageId=cfa9e580-a81e-4a4b-a846-7b21bf4e2e5b&host=www.microsoft.com&segments=software-download%2Cwindows11&query=&action=GetProductDownloadLinksBySku&sdVersion=2",
  sessionUrl:
    "https://vlscppe.microsoft.com/fp/tags?org_id=y6jn8c31&session_id=",
  apiUrl: "https://api.gravesoft.dev/msdl/",
};

const elements = {
  sessionId: document.getElementById("msdl-session-id"),
  msContent: document.getElementById("msdl-ms-content"),
  pleaseWait: document.getElementById("msdl-please-wait"),
  processingError: document.getElementById("msdl-processing-error"),
  productsList: document.getElementById("products-list"),
  backToProductsDiv: document.getElementById("back-to-products"),
};

const sharedSessionGUID = "47cbc254-4a79-4be6-9866-9c625eb20911";

let state = {
  availableProducts: {},
  sharedSession: false,
  shouldUseSharedSession: true,
  skuId: null,
};

const uuidv4 = () =>
  ([1e7] + -1e3 + -4e3 + -8e3 + -1e11).replace(/[018]/g, (c) =>
    (
      c ^
      (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (c / 4)))
    ).toString(16)
  );

const updateVars = () => {
  const id = document.getElementById("product-languages").value;
  document.getElementById("submit-sku").disabled = id === "";
  return id ? JSON.parse(id)["id"] : null;
};

const updateContent = (content, response) => {
  content.innerHTML = response;
  const errorMessage = document.getElementById("errorModalMessage");
  if (errorMessage) {
    elements.processingError.style.display = "block";
    return false;
  }
  return true;
};

const handleFetchResponse = async (response, onSuccess, onError) => {
  if (response.ok) {
    const text = await response.text();
    onSuccess(text);
  } else {
    onError();
  }
};

const fetchContent = async (url, onSuccess, onError = () => {}) => {
  try {
        const response = await fetch(url);
        return await handleFetchResponse(response, onSuccess, onError);
    } catch (onError) {
        return onError(onError);
    }
};

const getFromServer = () => {
  elements.processingError.style.display = "none";
  const url = `${urls.apiUrl}proxy?product_id=${window.location.hash.substring(
    1
  )}&sku_id=${state.skuId}`;
  fetchContent(url, displayResponseFromServer, () => {
    elements.processingError.style.display = "block";
  });
};

const displayResponseFromServer = (responseText) => {
  elements.pleaseWait.style.display = "none";
  elements.msContent.innerHTML = responseText;
};

const getLanguages = (productId) => {
  const url = `${urls.langsUrl}&productEditionId=${productId}&sessionId=${
    state.sharedSession ? sharedSessionGUID : elements.sessionId.value
  }`;
  fetchContent(url, (responseText) => {
    if (updateContent(elements.msContent, responseText)) {
      document
        .getElementById("submit-sku")
        .setAttribute("onClick", "getDownload();");
      document
        .getElementById("product-languages")
        .setAttribute("onChange", "updateVars();");
      updateVars();
    }
  });
};

const getDownload = () => {
  elements.msContent.style.display = "none";
  elements.pleaseWait.style.display = "block";
  state.skuId = state.skuId || updateVars();
  const url = `${urls.downUrl}&skuId=${state.skuId}&sessionId=${
    state.sharedSession ? sharedSessionGUID : elements.sessionId.value
  }`;
  fetchContent(url, (responseText) => {
    const wasSuccessful = updateContent(elements.msContent, responseText);
    if (wasSuccessful) {
      elements.pleaseWait.style.display = "none";
      if (!state.sharedSession) {
        fetch(`${urls.sessionUrl}${sharedSessionGUID}`);
        fetch(`${urls.sessionUrl}de40cb69-50a5-415e-a0e8-3cf1eed1b7cd`);
        fetch(
          `${urls.apiUrl}add_session?session_id=${elements.sessionId.value}`
        );
      }
    } else if (!state.sharedSession && state.shouldUseSharedSession) {
      useSharedSession();
    } else {
      getFromServer();
    }
  });
};

const backToProducts = () => {
  elements.backToProductsDiv.style.display = "none";
  elements.productsList.style.display = "block";
  elements.msContent.style.display = "none";
  elements.pleaseWait.style.display = "none";
  elements.processingError.style.display = "none";
  window.location.hash = "";
  state.skuId = null;
};

const useSharedSession = () => {
  state.sharedSession = true;
  retryDownload();
};

const retryDownload = () => {
  elements.pleaseWait.style.display = "block";
  elements.processingError.style.display = "none";
  const url = `${
    urls.langsUrl
  }&productEditionId=${window.location.hash.substring(
    1
  )}&sessionId=${sharedSessionGUID}`;
  fetchContent(url, getDownload);
};

const prepareDownload = (id) => {
  elements.productsList.style.display = "none";
  elements.backToProductsDiv.style.display = "block";
  elements.pleaseWait.style.display = "block";
  const url = `${urls.sessionUrl}${elements.sessionId.value}`;
  fetchContent(
    url,
    () => getLanguages(id),
    () => getLanguages(id)
  );
};

const addTableElement = (table, value, data) => {
  const a = document.createElement("a");
  a.href = `#${value}`;
  a.setAttribute("onClick", `prepareDownload(${value});`);
  a.textContent = data[value];
  const tr = table.insertRow();
  tr.insertCell().appendChild(a);
  tr.insertCell().textContent = value;
};

const createTable = (data, search) => {
  const table = document.getElementById("products-table-body");
  const regex = new RegExp(search, "ig");
  table.innerHTML = "";
  Object.keys(data).forEach((value) => {
    if (data[value].match(regex)) {
      addTableElement(table, value, data);
    }
  });
};

const updateResults = () => {
  const search = document.getElementById("search-products");
  createTable(state.availableProducts, search.value);
};

const setSearch = (query) => {
  document.getElementById("search-products").value = query;
  updateResults();
};

const checkHash = () => {
  const hash = window.location.hash;
  if (hash) {
    prepareDownload(hash.substring(1));
  }
};

const preparePage = (responseText) => {
  state.availableProducts = JSON.parse(responseText) || {};
  if (!Object.keys(state.availableProducts).length) {
    elements.pleaseWait.style.display = "none";
    elements.processingError.style.display = "block";
    return;
  }
  elements.pleaseWait.style.display = "none";
  elements.productsList.style.display = "block";
  updateResults();
  checkHash();
};

elements.sessionId.value = uuidv4();

const initializePage = () => {
  fetchContent("data/products.json", preparePage);
  elements.pleaseWait.style.display = "block";

  fetchContent(
    `${urls.apiUrl}use_shared_session`,
    () => {},
    () => {
      state.shouldUseSharedSession = false;
    }
  );
};

initializePage();
