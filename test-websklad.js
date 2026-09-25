// ============================================================
// test-websklad.js
// Набір автоматичних тестів та валідатор для фіду Websklad (KASTA)
// Запуск:
//   node test-websklad.js              -> запуск unit-тестів логіки
//   node test-websklad.js websklad.xml -> перевірка згенерованого XML
// ============================================================

import fs from 'fs';
import path from 'path';

// ─── Модуль перевірки логіки націнки та цін ──────────────────────────────────
function testMarkupCalculation() {
  console.log('🧪 Тест 1: Розрахунок націнки для товарів <= 1000 грн (45% + 40 грн)...');
  const calcPrice = (p) => p <= 1000 ? Math.round(p * 1.45 + 40) : Math.round(p * 1.40 + 50);

  // 200 грн -> 200 * 1.45 + 40 = 290 + 40 = 330
  assert(calcPrice(200) === 330, `Очікувалось 330 для 200 грн, отримано ${calcPrice(200)}`);
  // 500 грн -> 500 * 1.45 + 40 = 725 + 40 = 765
  assert(calcPrice(500) === 765, `Очікувалось 765 для 500 грн, отримано ${calcPrice(500)}`);
  // 1000 грн -> 1000 * 1.45 + 40 = 1450 + 40 = 1490
  assert(calcPrice(1000) === 1490, `Очікувалось 1490 для 1000 грн, отримано ${calcPrice(1000)}`);

  console.log('🧪 Тест 2: Розрахунок націнки для товарів > 1000 грн (40% + 50 грн)...');
  // 1500 грн -> 1500 * 1.40 + 50 = 2100 + 50 = 2150
  assert(calcPrice(1500) === 2150, `Очікувалось 2150 для 1500 грн, отримано ${calcPrice(1500)}`);
  // 3000 грн -> 3000 * 1.40 + 50 = 4200 + 50 = 4250
  assert(calcPrice(3000) === 4250, `Очікувалось 4250 для 3000 грн, отримано ${calcPrice(3000)}`);

  console.log('✅ Тести розрахунку націнки успішно пройдені.');
}

function testPromoAndOldPriceRules() {
  console.log('🧪 Тест 3: Розрахунок old_price (+20%) та price_promo (-5%)...');
  const price = 1000;
  const oldPrice = Math.round(price * 1.20); // 1200
  const promoPrice = Math.round(price * 0.95); // 950

  assert(oldPrice === 1200, `oldPrice для 1000 має бути 1200, отримано ${oldPrice}`);
  assert(promoPrice === 950, `promoPrice для 1000 має бути 950, отримано ${promoPrice}`);
  assert(promoPrice < price, 'promoPrice має бути суворо меншим за price');
  assert(oldPrice > price, 'oldPrice має бути суворо більшим за price');

  console.log('✅ Тест розрахунку промо-ціни та старої ціни успішно пройдено.');
}

function testOfferIdPrefix() {
  console.log('🧪 Тест 4: Префікс 1818 для offer id...');
  const sampleOffer = '<offer id="3325" selling_type="r" available="true" in_stock="true"><name>Тест</name></offer>';
  const processed = sampleOffer.replace(/(<offer\b[^>]*\bid=")([^"]+)(")/i, (m, prefix, id, suffix) => prefix + '1818' + id + suffix);
  assert(processed.includes('id="18183325"'), `Очікувався id="18183325", отримано: ${processed}`);
  console.log('✅ Тест префіксу ID успішно пройдено.');
}

function testStockFilterLogic() {
  console.log('🧪 Тест 5: Фільтрація тільки товарів у наявності...');
  const isAvailable = (offerXml) => {
    if (/\bavailable\s*=\s*["']false["']/i.test(offerXml)) return false;
    const mAvail = offerXml.match(/<available\b[^>]*>([\s\S]*?)<\/available>/i);
    if (mAvail && String(mAvail[1]).trim().toLowerCase() === 'false') return false;

    if (/\bin_stock\s*=\s*["']false["']/i.test(offerXml)) return false;
    const mStock = offerXml.match(/<in_stock\b[^>]*>([\s\S]*?)<\/in_stock>/i);
    if (mStock && String(mStock[1]).trim().toLowerCase() === 'false') return false;

    const mQty = offerXml.match(/<(?:quantity|stock_quantity)\b[^>]*>(\d+)<\/(?:quantity|stock_quantity)>/i);
    if (mQty && Number(mQty[1]) <= 0) return false;

    if (/<param\s+name=["'](?:Наявність|Наличие|Статус|Stock)["'][^>]*>[\s\S]*?(?:немає|нет|закінчи|out of stock|під замовлення)[\s\S]*?<\/param>/i.test(offerXml)) {
      return false;
    }
    return true;
  };

  assert(isAvailable('<offer available="true" in_stock="true"></offer>') === true, 'Товар у наявності має проходити');
  assert(isAvailable('<offer available="false"></offer>') === false, 'available=false має блокуватися');
  assert(isAvailable('<offer in_stock="false"></offer>') === false, 'in_stock=false має блокуватися');
  assert(isAvailable('<offer><quantity>0</quantity></offer>') === false, 'quantity=0 має блокуватися');
  assert(isAvailable('<offer><param name="Наявність">Немає в наявності</param></offer>') === false, 'Немає в наявності має блокуватися');

  console.log('✅ Тест фільтрації наявності успішно пройдено.');
}

function testDescriptionCleaning() {
  console.log('🧪 Тест 6: Очищення описів від шуму та заборонених маркерів Kasta...');
  const dirty = '<description><![CDATA[🔴 УВАГА ВІДЕО РОЗПАКОВКИ ТОВАРУ 🔴<br>Гарний товар. 👇 Детальніше на відео 👇]]></description>';
  let cleaned = dirty.replace(/<!\[CDATA\[/gi, '').replace(/\]\]>/g, '');
  cleaned = cleaned.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, ' ');
  cleaned = cleaned.replace(/🔴[^🔴\n]{0,80}(?:ВІДЕО|ВИДЕО)\s+(?:РОЗПАКОВКИ|РАСПАКОВКИ)[^🔴\n]{0,80}🔴/gi, '');
  cleaned = cleaned.replace(/[👇👉][^👇👉\n]{0,70}(?:Детальніше|Подробнее|Детальнее)[^👇👉\n]{0,30}[👇👉]/gi, '');
  cleaned = cleaned.trim();

  assert(!cleaned.includes('ВІДЕО РОЗПАКОВКИ'), 'Опис містить залишки шуму відео');
  assert(!cleaned.includes('Детальніше на відео'), 'Опис містить залишки посилань на відео');
  assert(cleaned.includes('Гарний товар'), 'Корисний текст опису не повинен бути видалений');
  console.log('✅ Тест очищення описів успішно пройдено.');
}

// ─── Валідатор повноти та коректності згенерованого XML-файлу ─────────────────
function validateXmlFeed(filePath) {
  console.log(`\n🔍 Початок валідації файлу: ${filePath}`);
  assert(fs.existsSync(filePath), `Файл ${filePath} не знайдено!`);

  const stat = fs.statSync(filePath);
  const sizeMb = (stat.size / (1024 * 1024)).toFixed(2);
  console.log(`📊 Розмір файлу: ${sizeMb} МБ`);

  // Мінімальний допустимий розмір для websklad
  assert(stat.size > 5 * 1024 * 1024, `Файл занадто малий (${sizeMb} МБ). Можливо, завантаження обірвалося!`);

  const content = fs.readFileSync(filePath, 'utf8');

  // Перевірка коректності заголовка і закриття каталогу
  assert(content.trim().startsWith('<?xml'), 'Файл повинен починатися з XML декларації: <?xml ... ?>');
  assert(content.trim().endsWith('</yml_catalog>'), 'Файл повинен коректно закриватися тегом </yml_catalog>');
  assert(content.includes('<shop>'), 'Файл повинен містити тег <shop>');
  assert(content.includes('</shop>'), 'Файл повинен закривати тег </shop>');
  assert(content.includes('<categories>'), 'Файл повинен містити блок <categories>');
  assert(content.includes('<offers>'), 'Файл повинен містити блок <offers>');

  // Збіг відкриваючих і закриваючих тегів offer
  const openCount = (content.match(/<offer\b/gi) || []).length;
  const closeCount = (content.match(/<\/offer>/gi) || []).length;
  console.log(`📦 Загальна кількість товарів у наявності: ${openCount}`);

  assert(openCount > 1000, `Занадто мало товарів (${openCount}). Очікується понад 1000!`);
  assert(openCount === closeCount, `Помилка структури XML: <offer> = ${openCount}, а </offer> = ${closeCount}`);

  // Перевірка правил Kasta для товарів
  const offerMatches = content.match(/<offer[\s\S]*?<\/offer>/g);
  let checkedCount = 0;
  let invalidIdCount = 0;
  let cheapPriceCount = 0;
  let missingOldPriceCount = 0;
  let missingPromoPriceCount = 0;
  let oldPriceErrorCount = 0;
  let promoPriceErrorCount = 0;
  let unavailCount = 0;
  let noNameCount = 0;
  let noPictureCount = 0;

  for (const offer of offerMatches) {
    checkedCount++;

    // 1. Перевірка префіксу 1818
    const idMatch = offer.match(/<offer\b[^>]*\bid="([^"]+)"/i);
    if (!idMatch || !idMatch[1].startsWith('1818')) {
      invalidIdCount++;
    }

    // 2. Перевірка наявності
    if (/\bavailable\s*=\s*["']false["']/i.test(offer) || /\bin_stock\s*=\s*["']false["']/i.test(offer)) {
      unavailCount++;
    }

    // 3. Перевірка ціни продажу
    const pMatch = offer.match(/<price\b[^>]*>([\d.,]+)<\/price>/i);
    const price = pMatch ? Number(pMatch[1]) : 0;
    if (!price || price < 150) {
      cheapPriceCount++;
    }

    // 4. Перевірка старої ціни (old_price та oldprice)
    const opMatch = offer.match(/<(?:old_price|oldprice)\b[^>]*>([\d.,]+)<\/(?:old_price|oldprice)>/i);
    if (!opMatch) {
      missingOldPriceCount++;
    } else {
      const oldPrice = Number(opMatch[1]);
      if (oldPrice <= price) {
        oldPriceErrorCount++;
      }
    }

    // 5. Перевірка промо-ціни (price_promo)
    const promoMatch = offer.match(/<price_promo\b[^>]*>([\d.,]+)<\/price_promo>/i);
    if (!promoMatch) {
      missingPromoPriceCount++;
    } else {
      const promoPrice = Number(promoMatch[1]);
      if (promoPrice >= price) {
        promoPriceErrorCount++;
      }
    }

    // 6. Перевірка наявності назви
    const hasName = /<(name|name_ua)\b[^>]*>[^<]+<\/(name|name_ua)>/i.test(offer);
    if (!hasName) noNameCount++;

    // 7. Перевірка зображення
    const hasPic = /<picture\b[^>]*>https?:\/\/[^<]+<\/picture>/i.test(offer);
    if (!hasPic) noPictureCount++;
  }

  console.log(`✅ Перевірено товарів: ${checkedCount}`);
  assert(invalidIdCount === 0, `Знайдено ${invalidIdCount} товарів без обов'язкового префіксу 1818!`);
  assert(unavailCount === 0, `Знайдено ${unavailCount} товарів, які не в наявності (мають бути видалені)!`);
  assert(cheapPriceCount === 0, `Знайдено ${cheapPriceCount} товарів з ціною < 150 грн!`);
  assert(missingOldPriceCount === 0, `Знайдено ${missingOldPriceCount} товарів без старої ціни (old_price)!`);
  assert(missingPromoPriceCount === 0, `Знайдено ${missingPromoPriceCount} товарів без промо-ціни (price_promo)!`);
  assert(oldPriceErrorCount === 0, `Знайдено ${oldPriceErrorCount} товарів де old_price <= price (помилка Kasta PRICE_ERROR)!`);
  assert(promoPriceErrorCount === 0, `Знайдено ${promoPriceErrorCount} товарів де price_promo >= price!`);
  assert(noNameCount === 0, `Знайдено ${noNameCount} товарів без назви!`);
  assert(noPictureCount === 0, `Знайдено ${noPictureCount} товарів без зображень!`);

  console.log('🎉 ВСІ КРИТИЧНІ ПРАВИЛА KASTA ТА СТРУКТУРА ФІДУ ВАЛІДНІ!');
}

function assert(condition, message) {
  if (!condition) {
    console.error(`\n❌ ТЕСТ ПРОВАЛЕНО: ${message}\n`);
    process.exit(1);
  }
}

// ─── Запуск ───────────────────────────────────────────────────────────────────
function main() {
  const args = process.argv.slice(2);
  const targetFile = args[0];

  console.log('====================================================');
  console.log('   WEBSKLAD FEED TEST SUITE (KASTA MARKETPLACE)     ');
  console.log('====================================================');

  // Спочатку завжди запускаємо unit-тести бізнес-логіки
  testMarkupCalculation();
  testPromoAndOldPriceRules();
  testOfferIdPrefix();
  testStockFilterLogic();
  testDescriptionCleaning();

  // Якщо передано шлях до XML файлу — проводимо повну валідацію файлу
  if (targetFile) {
    validateXmlFeed(targetFile);
  } else {
    console.log('\n💡 Підказка: передайте шлях до файлу (node test-websklad.js websklad.xml) для повної перевірки XML.');
  }

  console.log('\n🌟 УСІ ТЕСТИ УСПІШНО ПРОЙДЕНО!');
}

main();
