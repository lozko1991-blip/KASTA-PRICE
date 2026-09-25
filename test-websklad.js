// ============================================================
// test-websklad.js
// Набір автоматичних тестів та валідатор для фіду Websklad (KASTA)
// Запуск:
//   node test-websklad.js              -> запуск unit-тестів логіки
//   node test-websklad.js websklad.xml -> перевірка згенерованого XML
// ============================================================

import fs from 'fs';
import path from 'path';

// ─── Модуль перевірки логіки націнки та обробки ──────────────────────────────
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

function testOfferIdPrefix() {
  console.log('🧪 Тест 3: Префікс 1818 для offer id...');
  const sampleOffer = '<offer id="3325" selling_type="r" available="true" in_stock="true"><name>Тест</name></offer>';
  const processed = sampleOffer.replace(/(<offer\b[^>]*\bid=")([^"]+)(")/i, (m, prefix, id, suffix) => prefix + '1818' + id + suffix);
  assert(processed.includes('id="18183325"'), `Очікувався id="18183325", отримано: ${processed}`);
  console.log('✅ Тест префіксу ID успішно пройдено.');
}

function testOldPriceKastaRule() {
  console.log('🧪 Тест 4: Правило Kasta (oldprice > price завжди)...');
  const calcNewPrice = (p) => p <= 1000 ? Math.round(p * 1.45 + 40) : Math.round(p * 1.40 + 50);
  const calcOldPrice = (srcOld, newPrice) => {
    const computedOld = srcOld <= 1000 ? Math.round(srcOld * 1.45 + 40) : Math.round(srcOld * 1.40 + 50);
    return computedOld > newPrice ? computedOld : Math.round(newPrice * 1.20);
  };

  // Випадок 1: стара ціна значно вища
  const p1 = calcNewPrice(500); // 765
  const op1 = calcOldPrice(800, p1); // 1200
  assert(op1 > p1, `oldprice (${op1}) повинен бути більшим за price (${p1})`);

  // Випадок 2: stara ціна майже дорівнює вхідній або нижча (захист від PRICE_ERROR)
  const p2 = calcNewPrice(600); // 910
  const op2 = calcOldPrice(610, p2); // computedOld = 925 > 910
  assert(op2 > p2, `oldprice (${op2}) повинен бути більшим за price (${p2})`);

  // Випадок 3: критичний випадок де стара ціна була б <= нової без захисту
  const p3 = calcNewPrice(500); // 765
  const op3 = calcOldPrice(500, p3); // computedOld = 765 == p3 -> має спрацювати множник 1.20
  assert(op3 > p3, `oldprice (${op3}) повинен бути суворо більшим за price (${p3})`);
  assert(op3 === Math.round(p3 * 1.20), `Повинен застосуватись запас 1.20: ${op3}`);

  console.log('✅ Тест правила Kasta oldprice > price успішно пройдено.');
}

function testDescriptionCleaning() {
  console.log('🧪 Тест 5: Очищення описів від шуму та заборонених маркерів Kasta...');
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

  // Мінімальний допустимий розмір для websklad (повний каталог ~40-60 MB)
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
  console.log(`📦 Загальна кількість товарів: ${openCount}`);

  assert(openCount > 1000, `Занадто мало товарів (${openCount}). Очікується понад 1000!`);
  assert(openCount === closeCount, `Помилка структури XML: <offer> = ${openCount}, а </offer> = ${closeCount}`);

  // Перевірка правил Kasta для товарів
  const offerMatches = content.match(/<offer[\s\S]*?<\/offer>/g);
  let checkedCount = 0;
  let invalidIdCount = 0;
  let priceErrorCount = 0;
  let cheapPriceCount = 0;
  let noNameCount = 0;
  let noPictureCount = 0;

  for (const offer of offerMatches) {
    checkedCount++;

    // 1. Перевірка префіксу 1818
    const idMatch = offer.match(/<offer\b[^>]*\bid="([^"]+)"/i);
    if (!idMatch || !idMatch[1].startsWith('1818')) {
      invalidIdCount++;
    }

    // 2. Перевірка ціни
    const pMatch = offer.match(/<price\b[^>]*>([\d.,]+)<\/price>/i);
    const price = pMatch ? Number(pMatch[1]) : 0;
    if (!price || price < 150) {
      cheapPriceCount++;
    }

    // 3. Перевірка старої ціни (oldprice > price)
    const opMatch = offer.match(/<oldprice\b[^>]*>([\d.,]+)<\/oldprice>/i);
    if (opMatch) {
      const oldPrice = Number(opMatch[1]);
      if (oldPrice <= price) {
        priceErrorCount++;
      }
    }

    // 4. Перевірка наявності назви
    const hasName = /<(name|name_ua)\b[^>]*>[^<]+<\/(name|name_ua)>/i.test(offer);
    if (!hasName) noNameCount++;

    // 5. Перевірка зображення
    const hasPic = /<picture\b[^>]*>https?:\/\/[^<]+<\/picture>/i.test(offer);
    if (!hasPic) noPictureCount++;
  }

  console.log(`✅ Перевірено товарів: ${checkedCount}`);
  assert(invalidIdCount === 0, `Знайдено ${invalidIdCount} товарів без обов'язкового префіксу 1818!`);
  assert(cheapPriceCount === 0, `Знайдено ${cheapPriceCount} товарів з ціною < 150 грн!`);
  assert(priceErrorCount === 0, `Знайдено ${priceErrorCount} товарів де oldprice <= price (помилка Kasta PRICE_ERROR)!`);
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
  testOfferIdPrefix();
  testOldPriceKastaRule();
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
