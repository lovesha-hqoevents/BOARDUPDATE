const pptxgen = require('pptxgenjs');
const html2pptx = require('/Users/loveg/.claude/plugins/cache/anthropic-agent-skills/document-skills/69c0b1a06741/skills/pptx/scripts/html2pptx');
const path = require('path');

async function createPresentation() {
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'Blue Drop, LLC';
    pptx.title = 'FY26 Q2 Board Financial Update - YTD vs Budget';
    pptx.subject = 'Financial Performance Through January 19, 2026';

    const workDir = '/Users/loveg/Documents/SDO - Workspace/pptx_workspace';

    // Slide 1: Title
    console.log('Creating slide 1: Title');
    await html2pptx(path.join(workDir, 'slide1_title.html'), pptx);

    // Slide 2: Executive Summary with YTD vs Budget
    console.log('Creating slide 2: Executive Summary - YTD vs Budget');
    await html2pptx(path.join(workDir, 'slide2_exec_summary.html'), pptx);

    // Slide 3: YTD Revenue vs Q1 Budget with comparison chart
    console.log('Creating slide 3: YTD Revenue vs Q1 Budget');
    const { slide: slide3, placeholders: ph3 } = await html2pptx(path.join(workDir, 'slide3_ytd_performance.html'), pptx);

    if (ph3.length > 0) {
        // Clustered bar chart showing Q1 Budget vs YTD Actual
        slide3.addChart(pptx.charts.BAR, [
            {
                name: "Q1 Budget",
                labels: ["RECs", "Bloom Mktg", "Events", "IP", "Interest"],
                values: [1215163, 826235, 174000, 85232, 72500]
            },
            {
                name: "YTD Actual",
                labels: ["RECs", "Bloom Mktg", "Events", "IP", "Interest"],
                values: [2518742, 261185, 43299, 61500, 72048]
            }
        ], {
            ...ph3[0],
            barDir: 'bar',
            barGrouping: 'clustered',
            showLegend: true,
            legendPos: 'b',
            chartColors: ["1C2833", "27AE60"],
            showValue: false,
            valAxisDisplayUnit: 'thousands',
            valAxisMaxVal: 3000000
        });
    }

    // Slide 4: Segment Profitability Analysis
    console.log('Creating slide 4: YTD Segment Profitability');
    await html2pptx(path.join(workDir, 'slide4_profitability.html'), pptx);

    // Slide 5: Key Issues
    console.log('Creating slide 5: Key Issues');
    await html2pptx(path.join(workDir, 'slide5_key_issues.html'), pptx);

    // Slide 6: Board Actions
    console.log('Creating slide 6: Board Actions');
    await html2pptx(path.join(workDir, 'slide6_board_actions.html'), pptx);

    // Slide 7: Forecast with updated quarterly projections
    console.log('Creating slide 7: Full Year Forecast');
    const { slide: slide7, placeholders: ph7 } = await html2pptx(path.join(workDir, 'slide7_forecast.html'), pptx);

    if (ph7.length > 0) {
        slide7.addChart(pptx.charts.BAR, [
            {
                name: "Revenue",
                labels: ["Q1 (YTD)", "Q2 Forecast", "Q3 Forecast", "Q4 Forecast"],
                values: [3020957, 2400000, 2500000, 2500000]
            },
            {
                name: "Net Income",
                labels: ["Q1 (YTD)", "Q2 Forecast", "Q3 Forecast", "Q4 Forecast"],
                values: [1857668, 1100000, 1200000, 1200000]
            }
        ], {
            ...ph7[0],
            barDir: 'col',
            barGrouping: 'clustered',
            showTitle: true,
            title: 'FY26 Quarterly Forecast',
            showLegend: true,
            legendPos: 'b',
            chartColors: ["0077B6", "27AE60"],
            valAxisDisplayUnit: 'millions',
            valAxisMinVal: 0,
            valAxisMaxVal: 4000000,
            showValue: false
        });
    }

    // Save presentation
    const outputPath = '/Users/loveg/Documents/SDO - Workspace/FY26_Q2_Board_Presentation.pptx';
    await pptx.writeFile({ fileName: outputPath });
    console.log('Presentation saved to: ' + outputPath);
}

createPresentation().catch(err => {
    console.error('Error creating presentation:', err);
    process.exit(1);
});
