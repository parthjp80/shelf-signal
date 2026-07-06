// Shared between the frontend (src/App.jsx) and the backend (server/server.js).
// Plain ESM, no DOM/browser APIs, no JSX — safe to import from a Node process.

export function normalizeHeader(h) {
  return String(h || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

// Fallback classifier: infer a category from product text (name, or an
// Open Food Facts category tag) when no mapped category column is available.
// Order matters — more specific/branded keywords are listed before generic
// catch-alls (e.g. "snickers" before the bare "bar" in Bars) so a product
// like "Snickers Bar" resolves to Candy rather than Bars.
export const PRODUCT_CATEGORY_KEYWORDS = {
  'Drinks': ['soda', 'cola', 'pepsi', 'coke', 'sprite', 'juice', 'drink', 'beverage', 'tea', 'lemonade',
    'gatorade', 'powerade', 'water', 'sparkling', 'seltzer', 'energy drink', 'red bull', 'redbull',
    'monster', 'iced coffee', 'coffee', 'smoothie', 'kombucha'],
  'Chips': ['chip', 'crisp', 'tortilla', 'doritos', 'lays', "lay's", 'pringles', 'fritos', 'cheetos',
    'sun chips', 'popcorn', 'pretzel'],
  'Candy': ['candy', 'chocolate', 'gummy', 'gummies', 'snickers', 'skittles', 'm&m', "reese's", 'reeses',
    'twix', 'kitkat', 'kit kat', 'sour patch', 'starburst', 'mint', 'gum', 'licorice', 'jelly bean'],
  'Cookies': ['cookie', 'oreo', 'wafer'],
  'Crackers': ['cracker', 'ritz', 'goldfish', "cheez-it", 'cheez it'],
  'Nuts & Trail Mix': ['almond', 'peanut', 'cashew', 'trail mix', 'pistachio', 'nut'],
  'Jerky & Meat Snacks': ['jerky', 'meat stick', 'slim jim'],
  'Ice Cream & Frozen': ['ice cream', 'popsicle', 'frozen'],
  'Bakery': ['muffin', 'donut', 'doughnut', 'pastry', 'danish', 'bun'],
  'Bars': ['granola bar', 'protein bar', 'energy bar', 'snack bar', 'cereal bar', 'bar'],
};

// Matches keyword at a word boundary (start of a token), tolerating simple
// plurals like "donut" -> "donuts". Deliberately NOT a plain substring test:
// naive .includes() matched "cola" inside "chocolate" and "nut" inside "donuts".
export function guessCategoryFromProduct(text) {
  const norm = normalizeHeader(text);
  if (!norm) return '';
  for (const [category, keywords] of Object.entries(PRODUCT_CATEGORY_KEYWORDS)) {
    for (const k of keywords) {
      const kn = normalizeHeader(k);
      if (kn && new RegExp(`(^|\\s)${kn}`).test(norm)) return category;
    }
  }
  return '';
}
