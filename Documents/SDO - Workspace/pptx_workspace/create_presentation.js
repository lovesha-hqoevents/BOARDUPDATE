const pptxgen = require('pptxgenjs');
const html2pptx = require('/Users/loveg/.claude/plugins/cache/anthropic-agent-skills/document-skills/69c0b1a06741/skills/pptx/scripts/html2pptx');
const path = require('path');

async function createPresentation() {
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'Blue Drop, LLC';
    pptx.title = 'FY26 Q2 Board Update';
    pptx.subject = 'Quarterly Board Presentation Template';

    const workDir = '/Users/loveg/Documents/SDO - Workspace/pptx_workspace';
    const logoPath = path.join(workDir, 'bluedrop_logo.png');

    // Slide 1: Title with Logo
    console.log('Creating slide 1: Title');
    const slide1 = pptx.addSlide();
    slide1.addText([
        { text: 'Blue Drop Board', options: { fontSize: 40, bold: true, color: '1C4587', breakLine: true } },
        { text: 'Update', options: { fontSize: 40, bold: true, color: '1C4587', breakLine: true } },
        { text: '', options: { fontSize: 12, breakLine: true } },
        { text: 'FY26 - Q2 Financial Review', options: { fontSize: 24, color: '1C4587', breakLine: true } },
        { text: '', options: { fontSize: 12, breakLine: true } },
        { text: 'YTD Performance: October 1, 2025 - January 19, 2026', options: { fontSize: 18, color: '1C4587' } }
    ], { x: 0.5, y: 1.5, w: 6, h: 3, valign: 'middle' });
    slide1.addImage({ path: logoPath, x: 7.5, y: 1.5, w: 2, h: 1.5 });

    // Slide 2: Executive Summary
    console.log('Creating slide 2: Executive Summary');
    await html2pptx(path.join(workDir, 'slide2_exec_summary.html'), pptx);

    // Slide 3: YTD Revenue vs Q1 Budget
    console.log('Creating slide 3: YTD Revenue vs Q1 Budget');
    const { slide: slide3, placeholders: ph3 } = await html2pptx(path.join(workDir, 'slide3_ytd_performance.html'), pptx);
    if (ph3.length > 0) {
        slide3.addChart(pptx.charts.BAR, [
            { name: "Q1 Budget", labels: ["RECs", "Bloom Mktg", "Events", "IP", "Interest"], values: [1215163, 826235, 174000, 85232, 72500] },
            { name: "YTD Actual", labels: ["RECs", "Bloom Mktg", "Events", "IP", "Interest"], values: [2518742, 261185, 43299, 61500, 72048] }
        ], { ...ph3[0], barDir: 'bar', barGrouping: 'clustered', showLegend: true, legendPos: 'b', chartColors: ["1C4587", "50B432"], showValue: false, valAxisDisplayUnit: 'thousands', valAxisMaxVal: 3000000 });
    }

    // Slide 4: Segment Profitability
    console.log('Creating slide 4: YTD Segment Profitability');
    await html2pptx(path.join(workDir, 'slide4_profitability.html'), pptx);

    // Slide 5: Key Issues
    console.log('Creating slide 5: Key Issues');
    await html2pptx(path.join(workDir, 'slide5_key_issues.html'), pptx);

    // Slide 6: Board Actions
    console.log('Creating slide 6: Board Actions');
    await html2pptx(path.join(workDir, 'slide6_board_actions.html'), pptx);

    // Slide 7: Forecast
    console.log('Creating slide 7: Full Year Forecast');
    const { slide: slide7, placeholders: ph7 } = await html2pptx(path.join(workDir, 'slide7_forecast.html'), pptx);
    if (ph7.length > 0) {
        slide7.addChart(pptx.charts.BAR, [
            { name: "Revenue", labels: ["Q1 (YTD)", "Q2 Forecast", "Q3 Forecast", "Q4 Forecast"], values: [3020957, 2400000, 2500000, 2500000] },
            { name: "Net Income", labels: ["Q1 (YTD)", "Q2 Forecast", "Q3 Forecast", "Q4 Forecast"], values: [1857668, 1100000, 1200000, 1200000] }
        ], { ...ph7[0], barDir: 'col', barGrouping: 'clustered', showTitle: true, title: 'FY26 Quarterly Forecast', showLegend: true, legendPos: 'b', chartColors: ["058DC7", "50B432"], valAxisDisplayUnit: 'millions', valAxisMinVal: 0, valAxisMaxVal: 4000000, showValue: false });
    }

    // Slide 8: Bloom Sales Update
    console.log('Creating slide 8: Bloom Sales Update');
    const { slide: slide8, placeholders: ph8 } = await html2pptx(path.join(workDir, 'slide8_bloom_sales.html'), pptx);
    if (ph8.length > 0) {
        slide8.addChart(pptx.charts.BAR, [
            { name: "FY26 YTD", labels: ["Non-Ag", "Ag Direct", "Landscapers"], values: [45000, 35000, 10769] },
            { name: "Target", labels: ["Non-Ag", "Ag Direct", "Landscapers"], values: [75000, 50000, 12500] }
        ], { ...ph8[0], barDir: 'col', barGrouping: 'clustered', showLegend: true, legendPos: 'b', chartColors: ["50B432", "1C4587"], showValue: false });
    }

    // Slide 9: Bloom Marketing Highlights
    console.log('Creating slide 9: Bloom Marketing Highlights');
    await html2pptx(path.join(workDir, 'slide9_bloom_marketing.html'), pptx);

    // Slide 10: RECs Performance
    console.log('Creating slide 10: RECs Performance');
    const { slide: slide10, placeholders: ph10 } = await html2pptx(path.join(workDir, 'slide10_recs_performance.html'), pptx);
    if (ph10.length > 0) {
        slide10.addChart(pptx.charts.LINE, [
            { name: "REC Price ($/MWh)", labels: ["Oct", "Nov", "Dec", "Jan"], values: [45, 47, 48, 50] }
        ], { ...ph10[0], showLegend: false, chartColors: ["058DC7"], lineDataSymbol: 'circle', lineDataSymbolSize: 8 });
    }

    // Slide 11: HQO Events Overview
    console.log('Creating slide 11: HQO Events Overview');
    await html2pptx(path.join(workDir, 'slide11_hqo_overview.html'), pptx);

    // Slide 12: Other Initiatives
    console.log('Creating slide 12: Other Initiatives');
    await html2pptx(path.join(workDir, 'slide12_other_initiatives.html'), pptx);

    // Slide 13: Q&A
    console.log('Creating slide 13: Questions & Comments');
    const slide13 = pptx.addSlide();
    slide13.addText('Questions & Comments', { x: 0.5, y: 2, w: 9, h: 1.5, fontSize: 36, bold: true, color: '1C4587', align: 'center' });
    slide13.addText('Thank you for your time and engagement', { x: 0.5, y: 3.2, w: 9, h: 0.5, fontSize: 16, color: '666666', align: 'center' });
    slide13.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 3, y: 4, w: 4, h: 1, fill: { color: 'F8F9FA' }, line: { color: 'E0E0E0', width: 1 } });
    slide13.addText([
        { text: 'Blue Drop, LLC', options: { fontSize: 12, color: '1C4587', bold: true, breakLine: true } },
        { text: 'FY26 Q2 Board Update', options: { fontSize: 10, color: '666666' } }
    ], { x: 3, y: 4.15, w: 4, h: 0.7, align: 'center', valign: 'middle' });
    slide13.addImage({ path: logoPath, x: 4.25, y: 0.3, w: 1.5, h: 1.1 });

    // Save presentation
    const outputPath = '/Users/loveg/Documents/SDO - Workspace/FY26_Q2_Board_Presentation.pptx';
    await pptx.writeFile({ fileName: outputPath });
    console.log('Presentation saved to: ' + outputPath);
}

createPresentation().catch(err => {
    console.error('Error creating presentation:', err);
    process.exit(1);
});
