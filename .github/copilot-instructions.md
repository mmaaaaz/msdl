# Copilot Coding Agent Instructions for msdl

## Project Overview

- **Microsoft Software Download Listing (msdl)** is a static web app for browsing and downloading Microsoft products, using data from official Microsoft Software Download pages.
- The app is client-side only: all logic is in JavaScript (`js/msdl.js`), with no backend code in this repo.
- Product metadata is stored in `data/products.json` and must be updated manually when new products are released.

## Architecture & Data Flow

- **Entry point:** `index.html` loads `js/msdl.js` and `css/style.css`.
- **Product List:**
  - On load, `msdl.js` fetches `data/products.json` via XHR and populates the product table.
  - Users can search/filter products using regex or quick-search buttons.
- **Product Selection:**
  - Selecting a product triggers an API call to `https://api.gravesoft.dev/msdl/skuinfo?product_id=...` to fetch available languages.
  - After language selection, another API call to `https://api.gravesoft.dev/msdl/proxy?...` retrieves download links.
- **UI State:**
  - All UI state is managed via DOM manipulation in `msdl.js`.
  - Visibility of sections (`#products-list`, `#msdl-ms-content`, etc.) is toggled as the user navigates.

## Developer Workflows

- **No build step:** All files are static. Just open `index.html` in a browser to test.
- **Update product list:**
  - Edit `data/products.json` to add/remove products. Use the format `{ "id": "Product Name" }`.
  - No schema validation is enforced; keep keys as strings and values as product names.
- **Styling:**
  - Uses [Water.css](https://watercss.kognise.dev/) via CDN for base styles, with overrides in `css/style.css`.
- **API integration:**
  - All product/language/download data is fetched from the external `api.gravesoft.dev` service. No local API mocking.

## Project Conventions

- **No frameworks:** Pure JavaScript, no build tools, no dependencies.
- **DOM IDs:** All dynamic UI elements are referenced by hardcoded IDs (see `index.html` and `msdl.js`).
- **Error handling:** Errors are shown by toggling the `#msdl-processing-error` div.
- **Session:** A random session ID is generated per page load and stored in a hidden input (`#msdl-session-id`).
- **Accessibility:** No explicit ARIA or accessibility features beyond Water.css defaults.

## Key Files

- `index.html`: Main HTML, includes all UI elements and script/style references.
- `js/msdl.js`: All client logic, including data fetching, UI updates, and event handlers.
- `data/products.json`: List of available products (edit to update offerings).
- `css/style.css`: Custom styles overriding Water.css.

## License

- Licensed under GNU Affero General Public License v3. See `LICENSE` for details.

---

**For AI agents:**

- Do not add build tools, frameworks, or server-side code.
- When adding new products, update only `data/products.json`.
- When changing UI, update both `index.html` and `msdl.js` as needed.
- Keep all logic in plain JavaScript and static assets.
