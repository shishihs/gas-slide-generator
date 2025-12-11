
function renderSlides(slides) {
    const container = document.getElementById('app');

    slides.forEach((slide, index) => {
        const slideEl = document.createElement('div');
        slideEl.className = `slide-container slide-${slide.type}`;
        slideEl.id = `slide-${index + 1}`;

        let contentHtml = '';

        // Header (Title & Subhead) for standard slides
        if (slide.type !== 'title' && slide.type !== 'section') {
            contentHtml += `
                <div style="padding: 40px 60px 0 60px;">
                    <h2>${slide.title}</h2>
                    ${slide.subhead ? `<div class="subhead">${slide.subhead}</div>` : ''}
                </div>
                <div class="content-area">
            `;
        }

        // Specific Render Logic
        switch (slide.type) {
            case 'title':
                slideEl.innerHTML = `
                    <div style="flex:1; display:flex; flex-direction:column; justify-content:center; align-items:center; z-index:2;">
                        <h1>${slide.title}</h1>
                        <div class="subhead">${slide.subhead}</div>
                        <div class="meta">${slide.date} | ${slide.author}</div>
                    </div>
                `;
                break;

            case 'section':
                slideEl.innerHTML = `
                   <div style="flex:1; display:flex; flex-direction:column; justify-content:center; z-index:2;">
                        <h1>${slide.title}</h1>
                   </div>
                   ${slide.sectionNo ? `<div class="section-no">${slide.sectionNo.toString().padStart(2, '0')}</div>` : ''}
                `;
                break;

            case 'agenda':
                contentHtml += `<div class="agenda-list">`;
                slide.items.forEach((item, i) => {
                    contentHtml += `<div class="agenda-item">${i + 1}. ${item}</div>`;
                });
                contentHtml += `</div>`;
                break;

            case 'content':
                if (slide.twoColumn) {
                    contentHtml += `<div class="two-column">`;
                    slide.columns.forEach(col => {
                        contentHtml += `<div>`;
                        col.forEach(text => {
                            contentHtml += renderTextItem(text);
                        });
                        contentHtml += `</div>`;
                    });
                    contentHtml += `</div>`;
                } else {
                    contentHtml += `<ul>`;
                    slide.points.forEach(point => {
                        contentHtml += `<li>${parseMarkdown(point)}</li>`;
                    });
                    contentHtml += `</ul>`;
                }
                break;

            case 'table':
                contentHtml += `<table class="custom-table"><thead><tr>`;
                slide.headers.forEach(h => contentHtml += `<th>${h}</th>`);
                contentHtml += `</tr></thead><tbody>`;
                slide.rows.forEach(row => {
                    contentHtml += `<tr>`;
                    row.forEach(cell => contentHtml += `<td>${cell}</td>`);
                    contentHtml += `</tr>`;
                });
                contentHtml += `</tbody></table>`;
                break;

            case 'cards':
            case 'headerCards':
                const cols = slide.columns || 3;
                contentHtml += `<div class="cards-grid" style="grid-template-columns: repeat(${cols}, 1fr);">`;
                slide.items.forEach(item => {
                    contentHtml += `
                        <div class="card">
                            <div class="card-title">${item.title}</div>
                            ${item.desc ? `<div class="card-desc">${item.desc}</div>` : ''}
                        </div>
                    `;
                });
                contentHtml += `</div>`;
                break;

            case 'compare':
                contentHtml += `
                    <div class="compare-container">
                        <div class="compare-col">
                            <h3>${slide.leftTitle}</h3>
                            <ul>${slide.leftItems.map(i => `<li>${parseMarkdown(i)}</li>`).join('')}</ul>
                        </div>
                        <div class="compare-col">
                            <h3>${slide.rightTitle}</h3>
                            <ul>${slide.rightItems.map(i => `<li>${parseMarkdown(i)}</li>`).join('')}</ul>
                        </div>
                    </div>
                `;
                break;

            case 'pyramid':
                contentHtml += `<div class="pyramid-container">`;
                slide.levels.forEach(level => {
                    contentHtml += `
                        <div class="pyramid-level">
                            ${level.label}
                            ${level.subLabel ? `<span>${level.subLabel}</span>` : ''}
                        </div>
                    `;
                });
                contentHtml += `</div>`;
                if (slide.notes) contentHtml += `<div style="margin-top:20px; text-align:center;">${slide.notes}</div>`;
                break;

            case 'stepUp':
                contentHtml += `<div class="step-up-container">`;
                slide.items.forEach(item => {
                    contentHtml += `
                        <div class="step-item">
                            <div style="font-weight:bold; font-size:20px; margin-bottom:10px;">${item.title}</div>
                            <div>${item.desc}</div>
                        </div>
                     `;
                });
                contentHtml += `</div>`;
                break;

            case 'process':
                contentHtml += `<div class="process-steps">`;
                slide.steps.forEach((step, i) => {
                    contentHtml += `<div class="process-step">${step.replace('\n', '<br>')}</div>`;
                    if (i < slide.steps.length - 1) contentHtml += `<div class="process-arrow"></div>`;
                });
                contentHtml += `</div>`;
                break;

            case 'barCompare':
                contentHtml += `<div class="bar-chart">`;
                slide.stats.forEach(stat => {
                    const maxVal = 150; // hardcoded max for simplicity
                    const width = (stat.leftValue / maxVal) * 100;
                    contentHtml += `
                        <div class="bar-row">
                            <div class="bar-label">${stat.label}</div>
                            <div class="bar-track">
                                <div class="bar-fill ${stat.trend === 'up' ? 'highlight' : ''}" style="width: ${width}%;">${stat.leftValue}</div>
                            </div>
                        </div>
                    `;
                });
                contentHtml += `</div>`;
                if (slide.notes) contentHtml += `<p style="margin-top:20px;">${slide.notes}</p>`;
                break;

            default:
                contentHtml += `<p>Unknown Slide Type: ${slide.type}</p>`;
        }

        if (slide.type !== 'title' && slide.type !== 'section') {
            contentHtml += `</div>`; // Close content-area
            slideEl.innerHTML = contentHtml;
        }

        container.appendChild(slideEl);
    });
}

function renderTextItem(text) {
    if (text.startsWith('**') || text.startsWith('[[')) return `<p style="margin-bottom:10px;">${parseMarkdown(text)}</p>`;
    return `<p style="margin-bottom:10px;">${parseMarkdown(text)}</p>`;
}

function parseMarkdown(text) {
    if (!text) return '';
    return text
        .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
        .replace(/\[\[(.*?)\]\]/g, '<span class="highlight">$1</span>');
}

// Init
window.onload = () => {
    if (typeof mockData !== 'undefined') {
        renderSlides(mockData);
    } else {
        console.error('No mockData found');
    }
};
