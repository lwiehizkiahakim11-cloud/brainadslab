// Form handling & prompt generation
const form = document.getElementById('strategyForm');
const hasPreviousDataRadios = document.querySelectorAll('input[name="hasPreviousData"]');
const previousDataInputs = document.getElementById('previousDataInputs');

// Show/hide previous data inputs
hasPreviousDataRadios.forEach(radio => {
    radio.addEventListener('change', (e) => {
        previousDataInputs.style.display = e.target.value === 'yes' ? 'block' : 'none';
    });
});

// TAB NAVIGATION
document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        const tabName = btn.dataset.tab;
        
        // Remove active from all
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        
        // Add active to current
        btn.classList.add('active');
        document.getElementById(tabName).classList.add('active');
    });
});

// FORM SUBMISSION
form.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    // Collect form data
    const formData = new FormData(form);
    const data = Object.fromEntries(formData);
    
    // Get selected formats & angles
    data.selectedFormats = Array.from(document.querySelectorAll('input[name="formats"]:checked'))
        .map(cb => cb.value);
    data.selectedAngles = Array.from(document.querySelectorAll('input[name="angles"]:checked'))
        .map(cb => cb.value);

    // Check if has previous data
    data.hasPreviousData = document.querySelector('input[name="hasPreviousData"]:checked').value;

    // Call function to generate prompt
    generatePrompt(data);
    
    // Go to preview tab
    document.querySelector('[data-tab="preview-prompt"]').click();
});

// GENERATE PROMPT FUNCTION
function generatePrompt(data) {
    let prompt = `=== AD STRATEGY OPTIMIZER PROMPT ===
[Paste ke Claude yang sudah connected dengan ads platform]

---

CONTEXT:

## PREVIOUS CAMPAIGN PERFORMANCE DATA
`;

    if (data.hasPreviousData === 'yes' && data.campaignHistory) {
        prompt += `
${data.campaignHistory}

Lessons learned:
${data.lessonsLearned || 'N/A'}
`;
    } else {
        prompt += `
[First campaign - no historical data]
`;
    }

    prompt += `

## CURRENT PRODUCT/BUSINESS CONTEXT

Product: ${data.productName}
Description: ${data.productDesc}
Price point: ${data.pricePoint}
Niche: ${data.niche}
Target audience: ${data.targetAudience}
Pain points: ${data.painPoints}

## PLATFORM & GOALS

Platform: ${data.platform.toUpperCase()}
Campaign objective: ${data.campaignObjective}
Testing budget: ${data.testingBudget}
Minimum ROAS target: ${data.minROAS}:1
Maximum CPC budget: ${data.maxCPC}
Timeline to identify winners: ${data.timelineWinners} days
Primary conversion metric: ${data.conversionMetric}

## PREFERENCES

Preferred creative formats: ${data.selectedFormats.join(', ') || 'All formats'}
Preferred copy angles: ${data.selectedAngles.join(', ') || 'All angles'}
${data.additionalNotes ? `Additional notes: ${data.additionalNotes}` : ''}

---

## TASK

Based on above context, generate WINNING ADS RECOMMENDATION with:

### 1. ADS SET TESTING (Audience Strategy)
Create 3 test audience sets with:
- Audience specification (interests, lookalikes, custom segments)
- Budget allocation percentage
- Rationale and expected ROAS range
- Confidence level (High/Medium/Low)
${data.hasPreviousData === 'yes' ? '- Reference to historical winners where applicable' : ''}

### 2. CREATIVE TESTING (Format + Copy Variations)
Generate 3+ creative variations with:
- Format recommendation
- Copy angle & hook (exact wording)
- Visual style guidance
- CTA variation
- Why this should work
${data.hasPreviousData === 'yes' ? '- Confidence based on historical patterns' : ''}

### 3. INTEREST TESTING (Detailed Targeting)
Provide:
- Primary interests (core buyer intent)
- Secondary modifiers
- Exclusions
- Interest stack rankings by confidence
- Rationale for each

### 4. TESTING ROADMAP (14 days)
Day-by-day action plan with:
- Phase 1 (Days 1-3): Setup & learning
- Phase 2 (Days 4-7): Optimization & pivots
- Phase 3 (Days 8-14): Scaling & winning lock-in
- Decision checkpoints every 3 days

### 5. WINNING ADS PREDICTION
Based on data analysis:
- Predicted winning format & angle
- Predicted winning audience
- Expected ROAS range
- Win probability score
- Key risks & mitigation

---

## OUTPUT FORMAT

Use this exact structure:

**[ADS SET TESTING]**
Test Set 1: [Name]
├─ Audience: [Spec]
├─ Budget: [%]
├─ Expected ROAS: [Range]
├─ Confidence: [High/Med/Low]

[Repeat for Set 2, 3]

**[CREATIVE TESTING]**
Variation A:
├─ Format: [Type]
├─ Angle: [Description]
├─ Hook: "[Exact copy]"
├─ CTA: "[Exact copy]"
├─ Why: [Rationale]

[Repeat for variations]

**[INTEREST TESTING]**
🔴 HIGH CONFIDENCE (3+ successes):
├─ Stack: [Interests]
├─ Expected ROAS: [Range]

🟡 MEDIUM CONFIDENCE (1-2 successes):
├─ Stack: [Interests]
├─ Expected ROAS: [Range]

🟢 LOW CONFIDENCE (exploratory):
├─ Stack: [Interests]
├─ Expected ROAS: [Range]

**[TESTING ROADMAP - 14 DAYS]**
[Detailed day-by-day breakdown]

**[WINNING ADS PREDICTION]**
Best predicted winner:
├─ Creative: [Format + angle]
├─ Audience: [Segment]
├─ Expected ROAS: [Range]
├─ Win probability: [%]

---

Now:
1. Connect to ${data.platform} ads account
2. Set up audience sets per recommendations
3. Create ad creatives per variations
4. Launch testing with budget allocation
5. Monitor daily metrics
6. Report back daily for optimization

Format results as JSON for Excel import.
`;

    // Display prompt
    const promptPreview = document.getElementById('promptPreview');
    promptPreview.innerHTML = `<pre>${escapeHtml(prompt)}</pre>`;
    
    // Show copy button
    document.getElementById('copyPromptBtn').style.display = 'inline-block';
    document.getElementById('copyPromptBtn').onclick = () => {
        navigator.clipboard.writeText(prompt).then(() => {
            alert('✅ Prompt copied to clipboard!');
        });
    };

    // Generate Excel preview
    generateExcelPreview(data);

    // Store data globally for download
    window.generatedData = data;
    window.generatedPrompt = prompt;
    
    // Show download buttons in output tab
    setTimeout(() => {
        showDownloadButtons(data);
    }, 100);
}

// EXCEL PREVIEW
function generateExcelPreview(data) {
    const preview = document.getElementById('excelPreview');
    
    const previewHTML = `
        <table style="font-size: 12px; width: 100%; border-collapse: collapse;">
            <thead>
                <tr style="background: #f0f0f0;">
                    <th style="border: 1px solid #ddd; padding: 10px; text-align: left;">Audience Set</th>
                    <th style="border: 1px solid #ddd; padding: 10px; text-align: center;">Budget %</th>
                    <th style="border: 1px solid #ddd; padding: 10px; text-align: center;">Expected ROAS</th>
                    <th style="border: 1px solid #ddd; padding: 10px; text-align: center;">Confidence</th>
                </tr>
            </thead>
            <tbody>
                <tr style="background: #fff9e6;">
                    <td style="border: 1px solid #ddd; padding: 10px;">Lookalike 1% (High confidence)</td>
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">40%</td>
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">${data.minROAS || 2.5}:1</td>
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">🔴 High</td>
                </tr>
                <tr style="background: #fff;">
                    <td style="border: 1px solid #ddd; padding: 10px;">Interest Stack (Medium confidence)</td>
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">35%</td>
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">${data.minROAS - 0.5 || 2.0}:1</td>
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">🟡 Medium</td>
                </tr>
                <tr style="background: #f0f0f0;">
                    <td style="border: 1px solid #ddd; padding: 10px;">Broad Exploration (Low confidence)</td>
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">25%</td>
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">${data.minROAS - 1 || 1.5}:1</td>
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">🟢 Low</td>
                </tr>
            </tbody>
        </table>
        
        <p style="margin-top: 15px; color: #666; font-size: 13px;">
            💡 Download Excel untuk tracking lengkap: Campaign Overview, Daily Metrics, Creative Performance, Testing Roadmap, dan Insights
        </p>
    `;
    
    preview.innerHTML = previewHTML;
}

// SHOW DOWNLOAD BUTTONS
function showDownloadButtons(data) {
    document.getElementById('downloadPrompt').style.display = 'inline-block';
    document.getElementById('downloadExcel').style.display = 'inline-block';
    document.getElementById('downloadJSON').style.display = 'inline-block';
    document.getElementById('summarySection').style.display = 'block';
    
    // Download Prompt
    document.getElementById('downloadPrompt').onclick = () => {
        const element = document.createElement('a');
        element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(window.generatedPrompt));
        element.setAttribute('download', `ad-strategy-prompt-${new Date().toISOString().split('T')[0]}.txt`);
        element.style.display = 'none';
        document.body.appendChild(element);
        element.click();
        document.body.removeChild(element);
    };
    
    // Download JSON
    document.getElementById('downloadJSON').onclick = () => {
        const jsonData = JSON.stringify(window.generatedData, null, 2);
        const element = document.createElement('a');
        element.setAttribute('href', 'data:application/json;charset=utf-8,' + encodeURIComponent(jsonData));
        element.setAttribute('download', `ad-strategy-${new Date().toISOString().split('T')[0]}.json`);
        element.style.display = 'none';
        document.body.appendChild(element);
        element.click();
        document.body.removeChild(element);
    };
}

// HELPER
function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}