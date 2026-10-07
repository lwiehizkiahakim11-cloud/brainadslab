// Excel generation untuk monitoring sheet
document.getElementById('downloadExcel').addEventListener('click', () => {
    const data = window.generatedData;
    
    // Create workbook
    const wb = XLSX.utils.book_new();
    
    // Sheet 1: Campaign Overview
    const overviewData = [
        ['CAMPAIGN OVERVIEW'],
        ['Campaign name', data.productName],
        ['Platform', data.platform],
        ['Campaign objective', data.campaignObjective],
        ['Start date', new Date().toISOString().split('T')[0]],
        ['Total budget', data.testingBudget],
        ['KPI - Min ROAS', `${data.minROAS}:1`],
        ['KPI - Max CPC', data.maxCPC],
        ['Testing timeline', `${data.timelineWinners} days`],
        ['Status', 'In Progress'],
        ['Current ROAS', '=IF(COUNTA(D:D)>2, INDEX(D:D,COUNTA(D:D)), "Pending metrics"'] ]
    ];
    const overviewWS = XLSX.utils.aoa_to_sheet(overviewData);
    XLSX.utils.book_append_sheet(wb, overviewWS, 'Campaign Overview');
    
    // Sheet 2: Audience Sets Tracking
    const audienceHeaders = ['Set Name', 'Audience Spec', 'Budget %', 'Days Running', 'Impressions', 'Clicks', 'CTR %', 'Conversions', 'Cost/Conv', 'ROAS', 'Status', 'Notes'];
    const audienceData = [
        audienceHeaders,
        ['Test Set 1: Lookalike 1%', '1% Lookalike dari converter terbaik', 40, 0, '', '', '=IF(E2="","",E2/D2)', '', '', '=(E2*100)/D2', 'Pending', ''],
        ['Test Set 2: Interest Stack', 'Interest A + B + C, exclude competitor', 35, 0, '', '', '=IF(E3="","",E3/D3)', '', '', '=(E3*100)/D3', 'Pending', ''],
        ['Test Set 3: Broad Explore', 'Broad + behavior targeting, new segment', 25, 0, '', '', '=IF(E4="","",E4/D4)', '', '', '=(E4*100)/D4', 'Pending', '']
    ];
    const audienceWS = XLSX.utils.aoa_to_sheet(audienceData);
    audienceWS['!cols'] = [
        { wch: 20 }, { wch: 30 }, { wch: 10 }, { wch: 12 }, 
        { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, 
        { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 20 }
    ];
    XLSX.utils.book_append_sheet(wb, audienceWS, 'Audience Sets');
    
    // Sheet 3: Creative Performance
    const creativeHeaders = ['Creative ID', 'Format', 'Copy Angle', 'Audience Pairing', 'Impressions', 'Clicks', 'CTR %', 'Conversions', 'Cost/Conv', 'ROAS', 'Ad Fatigue', 'Status', 'Notes'];
    const creativeData = [
        creativeHeaders,
        ['CREAT_001', 'Video 15s', 'Urgency + Social Proof', 'Test Set 1', '', '', '', '', '', '', '', 'Pending', ''],
        ['CREAT_002', 'Video 15s', 'Benefit-driven', 'Test Set 1', '', '', '', '', '', '', '', 'Pending', ''],
        ['CREAT_003', 'Carousel', 'Curiosity gap', 'Test Set 2', '', '', '', '', '', '', '', 'Pending', ''],
        ['CREAT_004', 'Static', 'Educational', 'Test Set 3', '', '', '', '', '', '', '', 'Pending', '']
    ];
    const creativeWS = XLSX.utils.aoa_to_sheet(creativeData);
    creativeWS['!cols'] = [
        { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 15 }, 
        { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, 
        { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 20 }
    ];
    XLSX.utils.book_append_sheet(wb, creativeWS, 'Creative Performance');
    
    // Sheet 4: Daily Metrics
    const dailyHeaders = ['Date', 'Phase', 'Total Impressions', 'Total Clicks', 'Daily CTR %', 'Total Conversions', 'Daily Spend', 'Cumulative ROAS', 'Top Audience Set', 'Top Creative', 'Action Taken', 'Notes'];
    const dailyData = [dailyHeaders];
    for (let i = 1; i <= 14; i++) {
        dailyData.push([
            `=TODAY()+${i}`, 
            i <= 3 ? 'Phase 1' : i <= 7 ? 'Phase 2' : 'Phase 3',
            '', '', '', '', '', '', '', '', '', ''
        ]);
    }
    const dailyWS = XLSX.utils.aoa_to_sheet(dailyData);
    dailyWS['!cols'] = [
        { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, 
        { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, 
        { wch: 16 }, { wch: 14 }, { wch: 16 }, { wch: 20 }
    ];
    XLSX.utils.book_append_sheet(wb, dailyWS, 'Daily Metrics');
    
    // Sheet 5: Testing Roadmap
    const roadmapData = [
        ['TESTING ROADMAP - 14 DAYS'],
        [],
        ['Phase', 'Days', 'Planned Action', 'Metrics to Watch', 'Decision Criteria', 'Actual Results', 'Decision Made', 'Next Action'],
        ['Phase 1: Learning', 'Days 1-3', 'Launch all 3 audience sets with Tier A creative', 'CTR, CPC, Quality score', 'If CTR < 3% by Day 2 → pause set', '', '', ''],
        ['Phase 2: Optimization', 'Days 4-7', 'Scale winning set +30%, test Tier B creative', 'Conversion rate, ROAS trajectory', 'If ROAS < 1.5 by Day 6 → pivot angle', '', '', ''],
        ['Phase 3: Scaling', 'Days 8-14', 'Scale top performer +50%, explore cold audience', 'Sustained ROAS, ad fatigue signals', 'Lock winning combo by Day 12', '', '', '']
    ];
    const roadmapWS = XLSX.utils.aoa_to_sheet(roadmapData);
    roadmapWS['!cols'] = [
        { wch: 18 }, { wch: 10 }, { wch: 35 }, { wch: 20 }, 
        { wch: 25 }, { wch: 20 }, { wch: 18 }, { wch: 20 }
    ];
    XLSX.utils.book_append_sheet(wb, roadmapWS, 'Testing Roadmap');
    
    // Sheet 6: Winning Ads Summary
    const winnerData = [
        ['WINNING ADS SUMMARY'],
        [],
        ['Metric', 'Value'],
        ['Predicted Winner Format', '[From AI recommendation]'],
        ['Predicted Winner Angle', '[From AI recommendation]'],
        ['Predicted Winner Audience', '[From AI recommendation]'],
        ['Expected ROAS', `${data.minROAS}:1 - ${data.minROAS + 1}:1`],
        ['Avg Conversion Rate', '[Track when available]'],
        ['Cost per Conversion', `${data.maxCPC}`],
        ['Recommendation', '[Update after Day 14]'],
        ['Next Steps', '[Scaling strategy]']
    ];
    const winnerWS = XLSX.utils.aoa_to_sheet(winnerData);
    XLSX.utils.book_append_sheet(wb, winnerWS, 'Winners Summary');
    
    // Sheet 7: Analysis & Insights
    const insightData = [
        ['ANALYSIS & INSIGHTS'],
        [],
        ['Metric', 'Observation', 'Impact'],
        ['Format Performance', 'Video vs Carousel vs Static - which performed best?', 'Informs next campaign creative priority'],
        ['Angle Performance', 'Urgency vs Benefit vs Curiosity - which converted best?', 'Informs messaging strategy'],
        ['Audience Ranking', 'Which segments delivered best ROAS?', 'Budget allocation for scaling'],
        ['Interest Stack Effectiveness', 'Which interest combinations worked best?', 'Targeting refinement'],
        ['Ad Fatigue Timeline', 'When did creative fatigue appear?', 'Ad rotation & refresh frequency'],
        ['Copy Hook Performance', 'Which hooks/CTAs had highest CTR?', 'Copy testing winner for next campaign'],
        ['Seasonal Patterns', 'Did performance vary by day of week?', 'Optimal scheduling for future'],
        [],
        ['RECOMMENDATIONS FOR NEXT CAMPAIGN'],
        ['1. Focus on [winning format]'],
        ['2. Lead with [winning angle]'],
        ['3. Target primarily [winning audience]'],
        ['4. Explore variations of [winning combo]'],
        ['5. Test new angles on stable winners only']
    ];
    const insightWS = XLSX.utils.aoa_to_sheet(insightData);
    XLSX.utils.book_append_sheet(wb, insightWS, 'Insights');
    
    // Write file
    XLSX.writeFile(wb, `ad-monitoring-${data.productName}-${new Date().toISOString().split('T')[0]}.xlsx`);
});