# System Architecture & Technical Handover Document for AI Agents (HANDOVER_FOR_AI.md)

This document is specifically formatted and structured for **LLM-based AI Coding Assistants** (e.g. Antigravity, Claude, ChatGPT, GitHub Copilot) to immediately understand, maintain, modify, and deploy the **Kawabata Suisan (Yaeyama) E-Commerce & Shipment Operations System**.

---

## 1. System Architecture & Tech Stack

```
[Client (Mobile Browsers / LINE LIFF)]
  ├─ index.html   (Customer Storefront / Order Form)
  ├─ request.html (Customer Specimen Catch Request Form)
  ├─ picking.html (Internal Staff Shipment & Picking Console)
  └─ stock.html   (Internal Staff Stock & Product Catalog Console)
        │
        │ HTTP REST (JSON / Text payloads)
        ▼
[Backend: Google Apps Script (V8 Engine) - 無題.js]
  ├─ doGet(e)  -> Query products, lookup customer master, get picking list, get all stock
  └─ doPost(e) -> Create order, PDF invoice, ship status, add phone item, update stock, create product, upload image
        │
        ▼
[Database / Drive: Google Workspace]
  ├─ Google Sheets: 「石垣島観賞魚販売管理」本番 (ID: 1ZiNel8TGJ-e7JMolkRagVOSnr1UczvQzpTa-F-s4p3o)
  └─ Google Drive Folders:
        ├─ 川畑水産_納品書PDF (Auto-generated delivery invoices)
        └─ 川畑水産_観賞魚画像 (Auto-uploaded fish/coral photos with public thumbnail links)
```

- **Frontend**: Vanilla HTML5, CSS3, modern ES6+ JavaScript. No build step (Babel/Webpack not needed). Single-page standalone files hosted on GitHub Pages.
- **Backend**: Google Apps Script (GAS) running on V8 runtime. Exposed via Google Apps Script Web App Deployment.
- **Authentication & Authorization**:
  - Storefront: LINE Front-end Framework (LIFF SDK v2) automatically extracts LINE User ID (`sub` / `userId`) and queries the customer master.
  - Internal Consoles (`picking.html`, `stock.html`): Completely severed from public storefront navigation (hidden URLs, direct access for staff on-site). Mutual navigation links between picking and stock consoles.

---

## 2. Environment Constants & Configuration Dictionary

When assisting the user with modifications or deployments, AI MUST use these active production constants:

```javascript
// Google Apps Script Web App Endpoint
const GAS_API_URL = "https://script.google.com/macros/s/AKfycbxIYsPdreUjaOFutfshhQIwOLb_Ke1JprcDZj_LAT9oZnc-l-3jZLsPAsFd_QdQJAQ4eA/exec";

// LINE Ecosystem
const OFFICIAL_LINE_ID = "@721pyqfn";
const LINE_CHANNEL_ID  = "2011871129";
// CRITICAL: Notice the 13th character is uppercase 'I' (India), NOT lowercase 'l' (Lima):
const LIFF_ID           = "2011871129-P4IflaGc"; 

// Google Spreadsheet (Production Database)
const SPREADSHEET_ID   = "1ZiNel8TGJ-e7JMolkRagVOSnr1UczvQzpTa-F-s4p3o";
const SPREADSHEET_NAME = "「石垣島観賞魚販売管理」本番";

// Google Apps Script (Clasp Project)
const SCRIPT_ID        = "1tIl_FYj2PWaSfsvajNXkBgZdgqr4-NMtb9yZsbWPQmSvdfCVQR3n6lv0";
const DEPLOYMENT_ID    = "AKfycbxIYsPdreUjaOFutfshhQIwOLb_Ke1JprcDZj_LAT9oZnc-l-3jZLsPAsFd_QdQJAQ4eA";

// GitHub Pages Deployment Target
const GITHUB_PAGES_REPO = "kawabatasuisanyaeyama/kawabatasuisanyaeyama.github.io";
// GitHub Personal Access Token (stored securely in environment / secret manager)
const GITHUB_PAT        = "<GITHUB_PERSONAL_ACCESS_TOKEN>"; // e.g. ghp_xxxxxxxxxxxxxxxxxxxx
```

---

## 3. Repository File Map & Roles

```
c:\Users\81909\Desktop\shop-project\
├── .clasp.json          # Clasp project bindings (scriptId: 1tIl_FYj2...)
├── appsscript.json      # GAS Manifest (timeZone, scopes, webapp anonymous access)
├── index.html           # Public customer order form (responsive mobile-first)
├── request.html         # Public specimen collection request form
├── picking.html         # Staff picking and dispatch management console
├── stock.html           # Staff stock inventory & new product registration console
├── slides.html          # Executive presentation slide deck for stakeholder review
├── 無題.js              # GAS backend implementation (~1600+ LOC)
├── HANDOVER_FOR_HUMAN.md# Operational handover manual for human administrators
└── HANDOVER_FOR_AI.md   # This specification document for AI assistants
```

---

## 4. API Endpoints & Request/Response Contracts

### A. `doGet(e)` Endpoints

#### 1. `action=fetchProducts`
- **Purpose**: Fetch all live inventory from category sheets for catalog rendering.
- **Query Params**:
  - `action`: `"fetchProducts"`
  - `userId` (optional): Customer LINE User ID (resolves customer rate/group permissions).
  - `phone` (optional): Legacy phone parameter.
- **Response**:
```json
{
  "products": [
    {
      "id": "魚_1",
      "name": "カクレクマノミ (ペア)",
      "category": "魚",
      "price": 3200,
      "isAsk": false,
      "vol": 500,
      "stock": 4,
      "size": "4-5cm",
      "img": "https://..."
    }
  ],
  "customer": {
    "isRegistered": true,
    "name": "サンゴ堂 石垣店",
    "phone": "090-XXXX-XXXX",
    "rate": 0.9,
    "allowedCategories": ["魚", "サンゴ", ...]
  }
}
```

#### 2. `action=lookupCustomer` / `action=searchCustomer`
- **Purpose**: Query customer master by LINE User ID or Phone.
- **Query Params**: `userId`, `phone`, `name`.
- **Response**: `{ "found": true, "name": "...", "phone": "...", "pref": "東京都" }`

#### 3. `action=getPickingList`
- **Purpose**: Retrieve unfulfilled (or all) orders grouped for picking console.
- **Query Params**:
  - `showAll`: `"0"` (unshipped only) or `"1"` (includes shipped).
- **Response**:
```json
{
  "status": "success",
  "orders": [
    {
      "orderId": "ORD-20261011_120000",
      "orderDate": "2026/10/11 12:00",
      "customerName": "アクアショップA",
      "phone": "090-1111-2222",
      "pdfUrl": "https://drive.google.com/...",
      "ngDates": "10/14(水), 10/16(金)",
      "notes": "電話追加あり",
      "isShipped": false,
      "items": [
        {
          "rowNumber": 15,
          "category": "魚",
          "name": "ハタタテハゼ",
          "qty": 3,
          "price": 1200,
          "isShipped": false
        }
      ]
    }
  ],
  "totalSummary": [
    {
      "category": "魚",
      "items": [
        { "name": "ハタタテハゼ", "totalQty": 3, "orderCount": 1 }
      ]
    }
  ]
}
```

#### 4. `action=getStockList`
- **Purpose**: Retrieve full inventory across all categories for the `stock.html` management console.
- **Query Params**: `action`: `"getStockList"`
- **Response**:
```json
{
  "status": "success",
  "items": [
    {
      "category": "魚",
      "rowNumber": 2,
      "id": "魚_1",
      "name": "カクレクマノミ (ペア)",
      "size": "4-5cm",
      "price": 3200,
      "isAsk": false,
      "vol": 500,
      "stock": 4,
      "img": "https://drive.google.com/thumbnail?id=...&sz=w1000",
      "isSoldOut": false
    }
  ],
  "categories": ["魚", "サンゴ", "シャコガイ", "貝", "イソギンチャク", "擬岩", "その他"]
}
```

---

### B. `doPost(e)` Endpoints

Payload is sent as `application/json` (or `text/plain` to bypass CORS preflight):

#### 1. `action=order`
- **Payload**:
```json
{
  "action": "order",
  "items": [
    { "id": "魚_1", "name": "カクレクマノミ", "category": "魚", "qty": 2, "appliedPrice": 1500, "vol": 200 }
  ],
  "customer": { "name": "アクアショップA", "phone": "090...", "lineUserId": "U1234..." },
  "ngDates": ["10/13(火)", "10/15(木)"],
  "notes": "午前着希望"
}
```
- **Processing**:
  1. Copies `納品書テンプレート` sheet into a new Google Spreadsheet in Drive.
  2. Generates PDF via Google Drive export endpoint and saves to `川畑水産_納品書PDF` folder.
  3. Appends line items to `注文明細`.
  4. Deducts stock in category sheets.
- **Response**: `{ "status": "success", "orderId": "ORD-...", "invoiceUrl": "https://..." }`

#### 2. `action=updateShipStatus`
- **Payload**:
```json
{
  "action": "updateShipStatus",
  "orderId": "ORD-20261011_120000",
  "isShipped": true
}
```
- **Processing**: Updates Column K (発送済チェック) of matching rows in `注文明細`.
- **Response**: `{ "status": "success", "isShipped": true }`

#### 3. `action=addOrderItem` (Phone Order / Manual Item Addition)
- **Payload**:
```json
{
  "action": "addOrderItem",
  "orderId": "ORD-20261011_120000",
  "customerName": "...", 
  "phone": "...",
  "item": {
    "name": "タテジマキンチャクダイ",
    "category": "魚",
    "qty": 1,
    "price": 5000,
    "subtotal": 5000,
    "isAsk": false
  }
}
```
- **Processing**:
  - Clones order metadata if `orderId` exists. If `orderId === "NEW"`, generates `TEL-yyyyMMdd_HHmmss`.
  - Appends item to `注文明細`.
  - Decrements category sheet stock (allows stock <= 0).
- **Response**: `{ "status": "success", "message": "...", "orderId": "..." }`

#### 4. `action=updateStock` (Stock Increment / Decrement / Detail Update)
- **Payload**:
```json
{
  "action": "updateStock",
  "category": "魚",
  "rowNumber": 2,
  "stock": 5,
  "isSoldOut": false,
  "size": "4-5cm",
  "price": 1500,
  "vol": 100,
  "isAsk": false,
  "img": "https://..."
}
```
- **Processing**:
  - Updates Column F (在庫数), Column C (サイズ), Column D (単価), Column E (vol), Column G (画像URL).
  - When stock reaches 0 or `isSoldOut === true`, stamps Column H (売り切れ日時). Clears Column H when restocked.
- **Response**: `{ "status": "success", "updatedStock": 5, "isSoldOut": false }`

#### 5. `action=createProduct` (New Specimen Registration)
- **Payload**:
```json
{
  "action": "createProduct",
  "category": "魚",
  "tier": "NONE",
  "name": "カクレクマノミ (ペア)",
  "size": "4〜5cm前後",
  "price": 3200,
  "isAsk": false,
  "vol": 100,
  "stock": 3,
  "images": [
    { "base64": "data:image/jpeg;base64,...", "fileName": "photo_1.jpg" },
    { "base64": "data:image/jpeg;base64,...", "fileName": "photo_2.jpg" }
  ]
}
```
*(Note: Single photo legacy format `imageBase64` / `imageFileName` is also backward-compatible)*
- **Processing**:
  - For each image in `images`, uploads file to Google Drive folder `川畑水産_観賞魚画像` via `saveImageToDrive()`, sets public view permission, and generates direct thumbnail URL (`https://lh3.googleusercontent.com/d/{ID}` / `https://drive.google.com/thumbnail?id={ID}&sz=w1000`).
  - Auto-generates Product ID: `{TIER}_{CATEGORY}_{TIMESTAMP}` (e.g. `A_魚_1791730000000`) if tier is A/B/C, or `{CATEGORY}_{TIMESTAMP}` if tier is NONE.
  - Joins all image URLs with comma (`,`) and appends new row to the category sheet: `[id, name, size, price, vol, stock, commaSeparatedImgUrls, ""]`.
- **Response**: `{ "status": "success", "productId": "...", "imgUrl": "..." }`

#### 6. `action=uploadImage` (Update Product Image / Multi-Photo Upload)
- **Payload**:
```json
{
  "action": "uploadImage",
  "category": "魚",
  "rowNumber": 2,
  "images": [
    { "base64": "data:image/jpeg;base64,...", "fileName": "photo_1.jpg" },
    { "base64": "data:image/jpeg;base64,...", "fileName": "photo_2.jpg" }
  ]
}
```
- **Processing**: Uploads all images to Drive folder `川畑水産_観賞魚画像`, joins URLs with comma, and updates Column G (画像URL) of specified row.
- **Response**: `{ "status": "success", "imgUrl": "url1,url2", "imgUrls": ["url1", "url2"] }`

---

## 5. Google Spreadsheet Data Model (Sheet Schema)

### Sheet: `注文明細` (Order Line Items)
Zero-indexed column mapping in GAS:
- `Col 0 (A)`: 注文日時 (`yyyy/MM/dd HH:mm:ss`)
- `Col 1 (B)`: 注文番号 (`orderId`, e.g. `ORD-...` or `TEL-...`)
- `Col 2 (C)`: 店舗名 (`custName`)
- `Col 3 (D)`: 電話番号 (`phone`)
- `Col 4 (E)`: 商品カテゴリ (`category`)
- `Col 5 (F)`: 商品名 (`prodName`)
- `Col 6 (G)`: 数量 (`qty`, Integer)
- `Col 7 (H)`: 単価 (`price`)
- `Col 8 (I)`: 小計 (`subtotal`)
- `Col 9 (J)`: 納品書PDF (`pdfUrl`)
- `Col 10 (K)`: **発送済チェック** (`boolean`: TRUE / FALSE)
- `Col 11 (L)`: 発送LINEリンク (`=HYPERLINK(...)`)
- `Col 12 (M)`: 編集用URL
- `Col 13 (N)`: 受取不可日 (`ngDates`)
- `Col 14 (O)`: 備考 (`notes`)

### Sheet: `顧客マスタ` (Customer Master)
- `Col B`: LINE UID (`dbUserId`)
- `Col C`: 店舗名・氏名 (`dbName`)
- `Col D`: 電話番号 (`dbPhone`)
- `Col A`: 個別掛率 (`colRate`, e.g. 0.9 = 10% off)
- Other columns: Permission checkboxes for categories (`魚`, `サンゴ`, `擬岩` etc.) and product tiers (`A`, `B`, `C`).

### Sheets: `魚`, `サンゴ`, `貝`, `シャコガイ`, `イソギンチャク`, `擬岩`, `その他` (Inventory Catalogs)
- Standard headers: `[商品ID, 商品名, サイズ, 単価, 容量vol, 在庫, 画像URL, 売り切れ日時]`
- **Column G (`画像URL`)**: Stores single or multiple Drive thumbnail URLs separated by commas (e.g. `url1,url2`).
  - Storefront (`index.html`) and Staff console (`stock.html`) parse with `split(/[,\n]/).map(s=>s.trim()).filter(Boolean)`.
  - When `length > 1`, UI displays badge `📷 N枚` and enables multi-photo gallery navigation (arrows, thumbnails, swipe gestures).
- **Column F (`在庫`)**: Updated via `updateStock`. Direct numeric typing (e.g. `550`) and delta steppers (`±1`, `±5`) both target this column. Clearing or 0 sets Column H (`売り切れ日時`). Restocking clears Column H.

---

## 6. Deployment Workflows & Commands

### A. Deploying Frontend Changes (HTML / CSS / JS)
Frontend changes must be pushed to GitHub Pages repo `kawabatasuisanyaeyama/kawabatasuisanyaeyama.github.io`.

**PowerShell REST API Method (Direct commit to GitHub Pages)**:
```powershell
$token = $env:GITHUB_PAT # or "<YOUR_GITHUB_PERSONAL_ACCESS_TOKEN>"
$headers = @{ "Authorization" = "token $token"; "Accept" = "application/vnd.github.v3+json" }
$repo = "kawabatasuisanyaeyama/kawabatasuisanyaeyama.github.io"
$path = "picking.html" # or index.html / request.html

$sha = (Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/contents/$path" -Headers $headers -Method Get).sha
$bytes = [System.IO.File]::ReadAllBytes("c:\Users\81909\Desktop\shop-project\$path")
$b64 = [Convert]::ToBase64String($bytes)
$body = @{ message = "Update $path"; content = $b64; sha = $sha } | ConvertTo-Json

Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/contents/$path" -Headers $headers -Method Put -Body $body
```

**Git Local Repository Commit**:
```powershell
git add index.html picking.html request.html "無題.js"
git commit -m "Your commit message"
git push origin main
```

### B. Deploying Backend GAS Changes (`無題.js`)
Google Apps Script changes are deployed using **clasp**:
```powershell
# 1. Push code to Google Apps Script project
npx @google/clasp push --force

# 2. Redeploy the production Web App deployment
npx @google/clasp deploy -i AKfycbxIYsPdreUjaOFutfshhQIwOLb_Ke1JprcDZj_LAT9oZnc-l-3jZLsPAsFd_QdQJAQ4eA -d "Deployment description"
```

---

## 7. Critical Gotchas, Traps & Historical Failure Modes

Future AI agents MUST be aware of these past pitfalls to avoid repeating known issues:

1. **LIFF ID Character Casing Traps (`P4IflaGc`)**:
   - The LIFF ID is `2011871129-P4IflaGc`. The 13th letter is a capital **`I`**, NOT a lowercase `l`.
   - Never replace `P4IflaGc` with `P4lflaGc`. Doing so will trigger `LIFF app '...' was not found (404)`.

2. **Customer Master LINE UID Matching**:
   - Customer authentication is strictly performed by matching the LINE User ID (`sub` or `ctx.userId`).
   - Do NOT fall back to raw telephone number matching for customer identity resolution on the storefront, because anyone could enter another person's phone number and view their negotiated wholesale discount rates.

3. **Picking Console Isolation**:
   - `picking.html` MUST NOT be linked from `index.html` or `request.html`. Customers must never see picking or unfulfilled orders.
   - The user requested "ツーカーにして" (seamless, direct access) on `picking.html`, meaning NO modal password or PIN prompt on screen. Access security is maintained by URL obscurity (not indexed, no storefront links).

4. **Negative Inventory Allowance for Manual Phone Orders**:
   - In `addOrderItem`, staff can enter any quantity even if the stock in the spreadsheet is `0` or negative.
   - Do NOT add client-side or server-side stock-capping validations to `addOrderItem`. Phone sales frequently happen when fresh unlogged fish arrives on-site.

5. **GAS Web App Deployment Ownership**:
   - In `appsscript.json`, `"webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }` is required.
   - When deploying through clasp, ensure the executing Google account has editor rights on the container-bound spreadsheet.

6. **Google Drive Image Direct Links for Web Rendering**:
   - Standard Google Drive file links (`https://drive.google.com/file/d/{ID}/view`) will NOT render in browser `<img>` tags due to Google Drive HTTP headers.
   - Files MUST be made publicly readable (`DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW`), and the image source URL MUST use the direct thumbnail endpoint: `https://drive.google.com/thumbnail?id={ID}&sz=w1000`.

7. **Product ID Tier Prefix Scheme**:
   - Products with ID starting with `A_`, `B_`, or `C_` are restricted. The customer storefront checks `customer.tiers` permissions from `顧客マスタ` and hides restricted items if the customer lacks permissions.
   - When generating new product IDs in `createProduct`:
     - If Tier is selected (`A`, `B`, or `C`): ID = `{TIER}_{CATEGORY}_{Date.now()}`
     - If Tier is `NONE`: ID = `{CATEGORY}_{Date.now()}`

8. **GitHub Push Protection & Token Security**:
   - NEVER commit raw GitHub personal access tokens (`ghp_...`) into any markdown, json, or js file in this repository.
   - GitHub secret scanning triggers immediately and blocks the `git push origin main` command. Use placeholder `<GITHUB_PERSONAL_ACCESS_TOKEN>` or environment variables in documentation.

9. **Client-side SWR Caching & Cart Persistence (`index.html`)**:
   - **Storage Keys**:
     - `kawabata_catalog_cache_v1`: `{ timestamp, userId, data: { isRegistered, customer, products, userHistory } }`. Dual-stored in `sessionStorage` and `localStorage`.
     - `kawabata_cart_cache_v1`: Current cart dictionary `{ [productId]: qty }`.
     - `kawabata_ng_dates_v1`: Array of selected delivery blackout date strings (`["10/14(月)", ...]`).
     - `kawabata_active_tab`: Last viewed category tab name.
   - **Lifecycle on `DOMContentLoaded`**:
     1. `restoreCartFromStorage()`: Instantly rehydrates `cart` and `selectedNgDates`.
     2. `loadCatalogCache()`: If valid cache exists, immediately invokes `applyCatalogData(cachedData, true)` to present all products, prices, and cart counts in **<0.05 seconds**, completely bypassing the loading spinner. Activates header sync pill (`🔄 最新同期中...`).
     3. Concurrently, `initLiff()` and `fetchProducts(uid, isBackgroundSync = true)` query fresh inventory in background (SWR pattern).
     4. When background sync arrives: updates local cache, smooths inventory numbers into UI, maintains scroll offset (`window.scrollTo(0, prevScrollY)`), and clamps any cart item quantity exceeding updated stock without discarding the rest of the cart.
     5. On order success: clears `kawabata_cart_cache_v1` and `kawabata_ng_dates_v1`.

