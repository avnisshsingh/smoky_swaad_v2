/**
 * ==========================================
 * SMOKY SWAAD ERP
 * Constants
 * ==========================================
 */

// Sheet Names
const SHEETS = {

  POS: "POS",
  ORDERS: "Orders",
  ORDER_ITEMS: "OrderItems",
  CUSTOMERS: "Customers",
  MENU: "Menu",
  SETTINGS: "Settings",
  PURCHASE_REGISTER: "Purchase Register",
  RECIPE: "Recipe Master",
  DASHBOARD: "Dashboard",
  BUSINESS_COSTS: "BusinessCosts",
  CONFIG: "Config"

};

// POS Cell References
const POS = {

  // Customer Details
  CUSTOMER_NAME: "B2",
  MOBILE: "B3",
  ORDER_DATE: "B4",
  DELIVERY_AREA: "B5",
  HOUSE_ADDRESS: "B6",
  DELIVERY_SLOT: "B7",
  ORDER_TYPE: "B8",
  PAYMENT_MODE: "B9",
  PAYMENT_STATUS: "B10",
  SPECIAL_INSTRUCTIONS: "B11",

  // Order Table
  ITEM_START_ROW: 15,
  ITEM_END_ROW: 24,

  ITEM_COL: 1,      // A
  PRICE_COL: 2,     // B
  QTY_COL: 3,       // C
  TOTAL_COL: 4,     // D

  // Totals
  SUBTOTAL: "B27",
  DISCOUNT: "B28",
  DELIVERY_CHARGE: "B29",
  GRAND_TOTAL: "B30"

};