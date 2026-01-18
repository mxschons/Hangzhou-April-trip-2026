/**
 * MxSchons Tours - Multi-language Site Builder
 * Generates language-specific HTML files from translations.js
 *
 * Usage: node build-site.js
 * Output: index.html (EN), zh/index.html (ZH), de/index.html (DE)
 */

const fs = require('fs');
const path = require('path');

const translations = require('./translations.js');

// Languages to generate
const LANGUAGES = ['en', 'zh', 'de'];
const PRIMARY_LANG = 'en';

/**
 * Generate the language switcher HTML
 */
function generateLanguageSwitcher(currentLang) {
  const t = translations.langSwitcher[currentLang];
  const links = LANGUAGES.map(lang => {
    const href = lang === PRIMARY_LANG ? '../index.html' : (lang === currentLang ? '#' : `../${lang}/index.html`);
    const adjustedHref = currentLang === PRIMARY_LANG
      ? (lang === PRIMARY_LANG ? '#' : `${lang}/index.html`)
      : href;
    const activeClass = lang === currentLang ? ' class="active"' : '';
    return `<a href="${adjustedHref}"${activeClass}>${t[lang]}</a>`;
  }).join('');

  return `<div class="lang-switcher">${links}</div>`;
}

/**
 * Generate browser language detection script
 */
function generateLangDetectionScript() {
  return `
  <script>
    // Browser language detection - only redirect if no language preference stored
    (function() {
      if (sessionStorage.getItem('langChosen')) return;

      var lang = navigator.language || navigator.userLanguage;
      var currentPath = window.location.pathname;

      // Check if already on a language-specific page
      if (currentPath.includes('/zh/') || currentPath.includes('/de/')) return;

      // Redirect based on browser language
      if (lang.startsWith('zh')) {
        window.location.href = 'zh/index.html';
      } else if (lang.startsWith('de')) {
        window.location.href = 'de/index.html';
      }
    })();
  </script>`;
}

/**
 * Generate script to mark language as chosen when clicking switcher
 */
function generateLangChoiceScript() {
  return `
  <script>
    document.querySelectorAll('.lang-switcher a').forEach(function(link) {
      link.addEventListener('click', function() {
        sessionStorage.setItem('langChosen', 'true');
      });
    });
  </script>`;
}

/**
 * Generate meta tags for a language
 */
function generateMetaTags(lang) {
  const t = translations.meta[lang];
  return `<title>${t.title}</title>
  <meta name="description" content="${t.description}">`;
}

/**
 * Generate hero section HTML
 */
function generateHeroSection(lang) {
  const t = translations.hero[lang];
  return `
  <!-- Hero Section -->
  <section class="hero">
    <div class="hero-content">
      <div class="logo">
        <img src="${lang === PRIMARY_LANG ? '' : '../'}images/logo-mxschons.png" alt="MxSchons Tours" class="logo-img">
      </div>
      <p class="tagline">${t.tagline}</p>

      <h1 class="hero-title">
        <span class="chinese-title">${t.chineseTitle}</span>
        <span class="english-title">${t.englishTitle}</span>
      </h1>

      <div class="hero-meta">
        <span class="diamond">◆</span>
        <span class="dates">${t.dates}</span>
        <span class="diamond">◆</span>
      </div>

      <p class="hero-subtitle">${t.subtitle}</p>

      <div class="travelers-badge">
        ${t.travelers}
      </div>
    </div>

    <div class="hero-scroll">
      <span>${t.scroll}</span>
      <div class="scroll-arrow">↓</div>
    </div>
  </section>`;
}

/**
 * Generate welcome section HTML
 */
function generateWelcomeSection(lang) {
  const t = translations.welcome[lang];
  const imgPath = lang === PRIMARY_LANG ? '' : '../';

  const paragraphs = t.paragraphs.map(p => `          <p>${p}</p>`).join('\n\n');

  return `
  <!-- Welcome Section -->
  <section class="section section-mist" id="welcome">
    <div class="container">
      <div class="welcome-grid">
        <div class="welcome-photo">
          <img src="${imgPath}images/max-professional.JPG" alt="Max, your host">
        </div>
        <div class="welcome-content">
          <p class="section-label">${t.label}</p>
          <h2>${t.title}</h2>

          <p class="salutation">${t.salutation}</p>

${paragraphs}

          <p class="highlight-text">${t.highlight}</p>

          <div class="signature">
            <p class="closing">${t.closing}</p>
            <p class="signature-name">${t.signatureName}</p>
            <p class="signature-title">${t.signatureTitle}</p>
          </div>

          <p class="ps">${t.ps}</p>
        </div>
      </div>
    </div>
  </section>`;
}

/**
 * Generate travelers section HTML
 */
function generateTravelersSection(lang) {
  const t = translations.travelers[lang];
  const imgPath = lang === PRIMARY_LANG ? '' : '../';

  const travelers = [
    { name: 'Max', img: 'max.png', city: t.frankfurt, bioKey: 'max' },
    { name: 'Dion', img: 'dion.jpg', city: t.singapore, bioKey: 'dion' },
    { name: 'Margot', img: 'margot.png', city: t.frankfurt, bioKey: 'margot' },
    { name: 'Alex', img: 'alex.png', city: t.singapore, bioKey: 'alex' },
    { name: 'Lynn', img: 'lynn.png', city: t.singapore, bioKey: 'lynn' }
  ];

  const cards = travelers.map(tr => `
        <div class="traveler-card">
          <div class="traveler-photo">
            <img src="${imgPath}images/${tr.img}" alt="${tr.name}">
          </div>
          <h3>${tr.name}</h3>
          <p class="traveler-city">${tr.city}</p>
          <p class="traveler-bio">${t.bios[tr.bioKey]}</p>
        </div>`).join('\n');

  return `
  <!-- Travelers Section -->
  <section class="section" id="travelers">
    <div class="container">
      <p class="section-label">${t.label}</p>
      <h2>${t.title}</h2>

      <div class="travelers-grid">
${cards}
      </div>
    </div>
  </section>`;
}

/**
 * Generate overview section HTML
 */
function generateOverviewSection(lang) {
  const t = translations.overview[lang];

  const tableRows = t.tableRows.map(row => `
            <tr>
              <td>${row[0]}</td>
              <td>${row[1]}</td>
              <td>${row[2]}</td>
            </tr>`).join('');

  return `
  <!-- Journey Overview -->
  <section class="section section-mist" id="overview">
    <div class="container">
      <p class="section-label">${t.label}</p>
      <h2>${t.title}</h2>

      <div class="journey-routes">
        <p class="route">${t.route1}</p>
        <p class="route route-secondary">${t.route2}</p>
      </div>

      <p class="journey-description">${t.description}</p>

      <div class="overview-table-wrapper">
        <table class="overview-table">
          <thead>
            <tr>
              <th>${t.tableHeaders[0]}</th>
              <th>${t.tableHeaders[1]}</th>
              <th>${t.tableHeaders[2]}</th>
            </tr>
          </thead>
          <tbody>${tableRows}
          </tbody>
        </table>
      </div>
    </div>
  </section>`;
}

/**
 * Generate glance section HTML
 */
function generateGlanceSection(lang) {
  const t = translations.glance[lang];

  const details = t.details.map(d => `
        <div class="detail-card">
          <h4>${d.title}</h4>
          <p>${d.content}</p>
        </div>`).join('');

  return `
  <!-- At a Glance -->
  <section class="section" id="glance">
    <div class="container">
      <p class="section-label">${t.label}</p>
      <h2>${t.title}</h2>

      <p class="section-intro">${t.intro}</p>

      <div class="details-grid">
${details}
      </div>
    </div>
  </section>`;
}

/**
 * Generate highlights section HTML
 */
function generateHighlightsSection(lang) {
  const t = translations.highlights[lang];

  const highlights = t.items.map(h => `
        <div class="highlight-card">
          <div class="highlight-icon">${h.icon}</div>
          <h3>${h.title}</h3>
          <p>${h.desc}</p>
        </div>`).join('');

  return `
  <!-- Highlights -->
  <section class="section section-dark" id="highlights">
    <div class="container">
      <p class="section-label section-label-light">${t.label}</p>
      <h2>${t.title}</h2>

      <div class="highlights-grid">
${highlights}
      </div>
    </div>
  </section>`;
}

/**
 * Generate single day card HTML
 */
function generateDayCard(day, lang) {
  const imgPath = lang === PRIMARY_LANG ? '' : '../';
  const triviaData = translations.trivia[lang].dayCallouts;

  const imageHtml = day.image
    ? `        <div class="day-image">
          <img src="${imgPath}images/${day.image}" alt="">
        </div>\n`
    : '';

  const scheduleItems = day.schedule.map(item => `
          <div class="schedule-item">
            <span class="time">${item.time}</span>
            <span class="activity">${item.activity}</span>
          </div>`).join('');

  // Check if this day has a trivia callout
  const dayKey = `day${day.num}`;
  const triviaCallout = triviaData[dayKey]
    ? `
        <div class="trivia-callout">
          <p class="trivia-callout-label">${triviaData[dayKey].label}</p>
          <p>${triviaData[dayKey].text}</p>
        </div>`
    : '';

  return `
      <!-- Day ${day.num} -->
      <article class="day-card" id="day-${day.num}">
${imageHtml}        <div class="day-header">
          <div class="day-number">Day ${day.num}</div>
          <div class="day-info">
            <h3>${day.date}</h3>
            <p class="day-title">${day.title}</p>
          </div>
        </div>

        <p class="day-description">${day.desc}</p>

        <div class="schedule">${scheduleItems}
        </div>

        <p class="day-note">${day.note}</p>${triviaCallout}
      </article>`;
}

/**
 * Generate itinerary section HTML
 */
function generateItinerarySection(lang) {
  const t = translations.itinerary[lang];
  const hotels = translations.hotels[lang];

  // Hangzhou hotel images
  const hangzhouImages = [
    { src: 'hangzhou-daytime.jpg', alt: 'Hangzhou hotel view during daytime' },
    { src: 'hangzhou-pavilion-night.jpg', alt: 'Hangzhou pavilion at night' },
    { src: 'hangzhou-skyline-night.jpg', alt: 'Hangzhou skyline at night' }
  ];

  // Build days with hotel callouts inserted at appropriate positions
  let daysHtml = '';
  t.days.forEach(day => {
    daysHtml += generateDayCard(day, lang);

    // Insert Hangzhou hotel callout after Day 1
    if (day.num === 1) {
      daysHtml += generateHotelCallout(hotels.hangzhou, lang, hangzhouImages);
    }

    // Insert Shanghai hotel callout after Day 7
    if (day.num === 7) {
      daysHtml += generateHotelCallout(hotels.shanghai, lang, null); // null = use placeholders
    }
  });

  return `
  <!-- Day by Day Itinerary -->
  <section class="section" id="itinerary">
    <div class="container">
      <p class="section-label">${t.label}</p>
      <h2>${t.title}</h2>
${daysHtml}
    </div>
  </section>`;
}

/**
 * Generate hotel callout HTML
 */
function generateHotelCallout(hotelData, lang, images) {
  const imgPath = lang === PRIMARY_LANG ? '' : '../';

  let galleryContent;
  if (images && images.length > 0) {
    galleryContent = images.map(img =>
      `          <img src="${imgPath}images/${img.src}" alt="${img.alt}">`
    ).join('\n');
  } else {
    galleryContent = `          <div class="hotel-placeholder">${hotelData.placeholder}</div>
          <div class="hotel-placeholder">${hotelData.placeholder}</div>
          <div class="hotel-placeholder">${hotelData.placeholder}</div>`;
  }

  return `
      <!-- Hotel Callout -->
      <div class="hotel-callout">
        <div class="hotel-callout-header">
          <span class="hotel-callout-icon">${hotelData.icon}</span>
          <div>
            <h4 class="hotel-callout-title">${hotelData.title}</h4>
            <p class="hotel-callout-subtitle">${hotelData.subtitle}</p>
          </div>
        </div>
        <p class="hotel-callout-description">${hotelData.description}</p>
        <div class="hotel-gallery">
${galleryContent}
        </div>
      </div>`;
}

/**
 * Generate restaurants section HTML
 */
function generateRestaurantsSection(lang) {
  const t = translations.restaurants[lang];

  const hangzhouCards = t.hangzhou.map(r => {
    const featuredClass = r.featured ? ' restaurant-card-featured' : '';
    return `
        <div class="restaurant-card${featuredClass}">
          <h4>${r.name}</h4>
          <p>${r.desc}</p>
        </div>`;
  }).join('');

  const shanghaiCards = t.shanghai.map(r => {
    const featuredClass = r.featured ? ' restaurant-card-featured' : '';
    return `
        <div class="restaurant-card${featuredClass}">
          <h4>${r.name}</h4>
          <p>${r.desc}</p>
        </div>`;
  }).join('');

  return `
  <!-- Restaurant Guide -->
  <section class="section section-mist" id="restaurants">
    <div class="container">
      <p class="section-label">${t.label}</p>
      <h2>${t.title}</h2>

      <h3 class="subsection-title">${t.hangzhouTitle}</h3>
      <p class="section-intro">${t.hangzhouIntro}</p>

      <div class="restaurant-grid">
${hangzhouCards}
      </div>

      <h3 class="subsection-title">${t.shanghaiTitle}</h3>

      <div class="restaurant-grid">
${shanghaiCards}
      </div>

      <p class="restaurant-note">${t.note}</p>
    </div>
  </section>`;
}

/**
 * Generate practical section HTML
 */
function generatePracticalSection(lang) {
  const t = translations.practical[lang];

  const contacts = t.contacts.items.map(c => `
            <dt>${c.label}</dt>
            <dd>${c.value}</dd>`).join('');

  const timeline = t.booking.items.map(item => `            <li>${item}</li>`).join('\n');

  const phrases = t.phrases.items.map(p => `
            <dt>${p.en}</dt>
            <dd>${p.pinyin} · <span class="chinese">${p.zh}</span></dd>`).join('');

  const packing = t.packing.items.map(item => `            <li>${item}</li>`).join('\n');

  return `
  <!-- Practical Info -->
  <section class="section" id="practical">
    <div class="container">
      <p class="section-label">${t.label}</p>
      <h2>${t.title}</h2>

      <div class="practical-grid">
        <div class="practical-card">
          <h3>${t.contacts.title}</h3>
          <dl class="contact-list">${contacts}
          </dl>
        </div>

        <div class="practical-card">
          <h3>${t.booking.title}</h3>
          <ul class="timeline-list">
${timeline}
          </ul>
        </div>

        <div class="practical-card">
          <h3>${t.phrases.title}</h3>
          <dl class="phrase-list">${phrases}
          </dl>
        </div>

        <div class="practical-card">
          <h3>${t.packing.title}</h3>
          <ul class="packing-list">
${packing}
          </ul>
        </div>
      </div>
    </div>
  </section>`;
}

/**
 * Generate trivia section HTML
 */
function generateTriviaSection(lang) {
  const t = translations.trivia[lang];

  const cards = t.cards.map(card => `
        <div class="trivia-card">
          <span class="trivia-location">${card.location}</span>
          <h4>${card.title}</h4>
          <p>${card.text}</p>
        </div>`).join('');

  const cultureItems = t.culture.map(item => `
        <div class="culture-item">
          <h5>${item.title}</h5>
          <p>${item.text}</p>
        </div>`).join('');

  return `
  <!-- Trivia Section -->
  <section class="section section-mist" id="trivia">
    <div class="container">
      <p class="section-label">${t.label}</p>
      <h2>${t.title}</h2>

      <p class="section-intro">${t.intro}</p>

      <div class="trivia-grid">
${cards}
      </div>

      <h3 class="subsection-title" style="margin-top: 48px;">${t.cultureTitle}</h3>

      <div class="culture-grid">
${cultureItems}
      </div>
    </div>
  </section>`;
}

/**
 * Generate footer HTML
 */
function generateFooter(lang) {
  const t = translations.footer[lang];

  return `
  <!-- Footer -->
  <footer class="footer">
    <div class="container">
      <div class="footer-content">
        <div class="footer-blessing">
          <span class="chinese-blessing">${t.blessing}</span>
          <span class="blessing-translation">${t.translation}</span>
        </div>

        <div class="footer-diamond">◆</div>

        <div class="footer-brand">
          <p>${t.brand}</p>
          <p class="footer-tagline">${t.tagline}</p>
        </div>
      </div>
    </div>
  </footer>`;
}

/**
 * Generate navigation HTML
 */
function generateNav(lang) {
  const t = translations.nav[lang];

  return `
  <!-- Navigation -->
  <nav class="nav">
    <div class="nav-inner">
      ${generateLanguageSwitcher(lang)}
      <a href="#welcome">${t.welcome}</a>
      <a href="#travelers">${t.travelers}</a>
      <a href="#overview">${t.overview}</a>
      <a href="#highlights">${t.highlights}</a>
      <a href="#itinerary">${t.itinerary}</a>
      <a href="#restaurants">${t.dining}</a>
      <a href="#practical">${t.practical}</a>
      <a href="#trivia">${t.trivia}</a>
    </div>
  </nav>`;
}

/**
 * Generate complete HTML for a language
 */
function generateFullHTML(lang) {
  const isRoot = lang === PRIMARY_LANG;
  const cssPath = isRoot ? 'styles.css' : '../styles.css';
  const langDetection = isRoot ? generateLangDetectionScript() : '';

  return `<!DOCTYPE html>
<html lang="${lang}">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  ${generateMetaTags(lang)}

  <!-- Prevent search engine indexing -->
  <meta name="robots" content="noindex, nofollow, noarchive, nosnippet, noimageindex">
  <meta name="googlebot" content="noindex, nofollow">
  <meta name="bingbot" content="noindex, nofollow">

  <!-- Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Instrument+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link href="https://cdn.jsdelivr.net/npm/lxgw-wenkai-webfont@1.7.0/style.css" rel="stylesheet">

  <link rel="stylesheet" href="${cssPath}">
  ${langDetection}
</head>
<body>
${generateHeroSection(lang)}
${generateWelcomeSection(lang)}
${generateTravelersSection(lang)}
${generateOverviewSection(lang)}
${generateGlanceSection(lang)}
${generateHighlightsSection(lang)}
${generateItinerarySection(lang)}
${generateRestaurantsSection(lang)}
${generatePracticalSection(lang)}
${generateTriviaSection(lang)}
${generateFooter(lang)}
${generateNav(lang)}
${generateLangChoiceScript()}
</body>
</html>
`;
}

/**
 * Main build function
 */
function build() {
  const websiteDir = __dirname;

  console.log('🌐 MxSchons Tours - Building multi-language site...\n');

  LANGUAGES.forEach(lang => {
    const html = generateFullHTML(lang);

    let outputPath;
    if (lang === PRIMARY_LANG) {
      outputPath = path.join(websiteDir, 'index.html');
    } else {
      const langDir = path.join(websiteDir, lang);
      if (!fs.existsSync(langDir)) {
        fs.mkdirSync(langDir, { recursive: true });
      }
      outputPath = path.join(langDir, 'index.html');
    }

    fs.writeFileSync(outputPath, html, 'utf8');
    console.log(`  ✓ Generated ${lang === PRIMARY_LANG ? 'index.html' : lang + '/index.html'} (${lang.toUpperCase()})`);
  });

  console.log('\n✨ Build complete! Generated files:');
  console.log('   - index.html (English - primary)');
  console.log('   - zh/index.html (Chinese)');
  console.log('   - de/index.html (German)');
  console.log('\n📝 Note: Browser language detection active on index.html');
  console.log('   Users can override via language switcher in navigation.\n');
}

// Run build
build();
