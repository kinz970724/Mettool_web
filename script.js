// --- Globals ---
const worker = new Worker('worker.js');
const loader = document.getElementById('loader-overlay');
let periods = [];

// Buffers to store the last generated Excel file data
let lastCorrBuffer = null;
let lastCusumBuffer = null;
let lastCtrlBuffer = null;

// --- DOM Element Cache ---
const fileInput = document.getElementById('file-input');
const sheetSelect = document.getElementById('sheet-name');
const btnLoad = document.getElementById('btn-load');
const btnClear = document.getElementById('btn-clear');

const btnPlotCorr = document.getElementById('btn-plot-corr');
const btnExportCorr = document.getElementById('btn-export-corr');
const btnPlotCusum = document.getElementById('btn-plot-cusum');
const btnExportCusum = document.getElementById('btn-export-cusum');
const btnPlotCtrl = document.getElementById('btn-plot-ctrl');
const btnExportCtrl = document.getElementById('btn-export-ctrl');
const btnAddPeriod = document.getElementById('btn-period-add');

// --- NEW: Professional UI Functions ---

/**
 * Shows a toast notification.
 * @param {string} message The text to display.
 * @param {'success' | 'error'} type The type of toast.
 */
function showToast(message, type = 'error') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;

    container.appendChild(toast);

    // Remove the toast after it fades out
    setTimeout(() => {
        toast.remove();
    }, 4000); // 3.5s fade + 0.5s buffer
}

/**
 * Toggles the loading state of a button.
 * @param {HTMLButtonElement} button The button element.
 * @param {boolean} isLoading Whether to show the loading spinner.
 */
function setButtonLoading(button, isLoading) {
    if (isLoading) {
        button.classList.add('btn-loading');
        button.disabled = true;
    } else {
        button.classList.remove('btn-loading');
        button.disabled = false;
    }
}

/**
 * Sets the disabled state for all analysis controls.
 * @param {boolean} isDisabled Whether to disable the controls.
 */
function setControlsDisabled(isDisabled) {
    btnPlotCorr.disabled = isDisabled;
    btnPlotCusum.disabled = isDisabled;
    btnPlotCtrl.disabled = isDisabled;
    btnAddPeriod.disabled = isDisabled;
    // Export buttons remain disabled until a plot is made
}

// --- Main UI Logic ---

function openTab(evt, tabName) {
    const tabContents = document.getElementsByClassName('tab-content');
    for (let i = 0; i < tabContents.length; i++) {
        tabContents[i].classList.remove('active-tab');
    }
    const tabLinks = document.getElementsByClassName('tab-link');
    for (let i = 0; i < tabLinks.length; i++) {
        tabLinks[i].className = tabLinks[i].className.replace(' active', '');
    }
    document.getElementById(tabName).classList.add('active-tab');
    evt.currentTarget.className += ' active';
}

function populateColumnLists(columns) {
    const allSelects = [
        { el: document.getElementById('corr-date-col'), default: 'Date' },
        { el: document.getElementById('cusum-date-col'), default: 'Date' },
        { el: document.getElementById('ctrl-date-col'), default: 'Date' },
        { el: document.getElementById('corr-cols'), isList: true },
        { el: document.getElementById('cusum-cols'), isList: true },
        { el: document.getElementById('ctrl-col'), isList: false }
    ];

    allSelects.forEach(sel => {
        sel.el.innerHTML = ''; // Clear existing

        const numericCols = columns.filter(c => c.toLowerCase() !== 'date');
        const targetCols = sel.isList ? numericCols : (sel.default ? columns : numericCols);

        targetCols.forEach(col => {
            const option = document.createElement('option');
            option.value = col;
            option.textContent = col;
            sel.el.appendChild(option);
        });

        if (sel.default && columns.includes(sel.default)) {
            sel.el.value = sel.default;
        } else if (targetCols.length > 0) {
            sel.el.value = targetCols[0];
        }
    });
}

function triggerDownload(buffer, filename) {
    if (!buffer || buffer.byteLength === 0) {
        showToast('No data to export.', 'error');
        return;
    }
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    console.log(`Exported ${filename}`);
}

// --- Worker Communication ---

worker.onmessage = (e) => {
    const { type, payload, error } = e.data;

    if (type === 'error') {
        // Handle errors from the worker
        showToast(error, 'error');
        // Re-enable any buttons that were loading
        if (btnPlotCorr.classList.contains('btn-loading')) setButtonLoading(btnPlotCorr, false);
        if (btnPlotCusum.classList.contains('btn-loading')) setButtonLoading(btnPlotCusum, false);
        if (btnPlotCtrl.classList.contains('btn-loading')) setButtonLoading(btnPlotCtrl, false);
        return;
    }

    switch (type) {
        case 'ready':
            loader.style.display = 'none';
            console.log('Pyodide is ready.');
            // Do not enable controls yet, wait for file load
            break;

        case 'sheetNames':
            sheetSelect.innerHTML = '';
            payload.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                sheetSelect.appendChild(option);
            });
            sheetSelect.disabled = false;
            console.log('Sheet names loaded:', payload);
            break;

        case 'dataLoaded':
            populateColumnLists(payload.columns);
            setControlsDisabled(false); // Enable plot buttons!
            showToast('Data loaded successfully!', 'success');
            break;

        case 'correlationResult':
            setButtonLoading(btnPlotCorr, false); // Stop spinner
            drawCorrelationPlot(payload.plotData);
            lastCorrBuffer = payload.fileBuffer; // Save buffer
            btnExportCorr.disabled = false; // Enable export
            break;

        case 'cusumResult':
            setButtonLoading(btnPlotCusum, false); // Stop spinner
            drawCusumPlot(payload.plotData);
            lastCusumBuffer = payload.fileBuffer; // Save buffer
            btnExportCusum.disabled = false; // Enable export
            break;

        case 'controlGraphResult':
            setButtonLoading(btnPlotCtrl, false); // Stop spinner
            drawControlGraphPlot(payload.plotData, payload.col, payload.conf);
            lastCtrlBuffer = payload.fileBuffer; // Save buffer
            btnExportCtrl.disabled = false; // Enable export
            break;
    }
};

// --- Event Listeners ---

fileInput.onchange = async () => {
    const file = fileInput.files[0];
    if (file) {
        const buffer = await file.arrayBuffer();
        worker.postMessage({
            type: 'getSheetNames',
            buffer: buffer
        });
        sheetSelect.disabled = true;
    }
};

btnLoad.onclick = async () => {
    const file = fileInput.files[0];
    const sheetName = sheetSelect.value;
    if (!file || !sheetName) {
        showToast('Please select a file and a sheet.', 'error');
        return;
    }
    const buffer = await file.arrayBuffer();
    worker.postMessage({
        type: 'loadData',
        buffer: buffer,
        sheetName: sheetName
    });
};

btnClear.onclick = () => {
    fileInput.value = '';
    sheetSelect.innerHTML = '';
    sheetSelect.disabled = true;
    populateColumnLists([]);
    setControlsDisabled(true); // Disable buttons
    // Disable export buttons and clear buffers
    btnExportCorr.disabled = true;
    btnExportCusum.disabled = true;
    btnExportCtrl.disabled = true;
    lastCorrBuffer = null;
    lastCusumBuffer = null;
    lastCtrlBuffer = null;
    // Clear plot areas
    document.getElementById('plot-corr').innerHTML = '';
    document.getElementById('plot-cusum').innerHTML = '';
    document.getElementById('plot-ctrl').innerHTML = '';
    worker.postMessage({ type: 'clearData' });
    console.log('Data cleared.');
};

btnPlotCorr.onclick = () => {
    const selectedCols = Array.from(document.getElementById('corr-cols').selectedOptions).map(opt => opt.value);
    const payload = {
        start: document.getElementById('corr-start').value,
        end: document.getElementById('corr-end').value,
        dateCol: document.getElementById('corr-date-col').value,
        cols: selectedCols
    };
    if (!payload.start || !payload.end || payload.cols.length === 0) {
        showToast('Please select dates and at least one column.', 'error');
        return;
    }
    setButtonLoading(btnPlotCorr, true); // Start spinner
    worker.postMessage({ type: 'runCorrelation', payload });
};
btnExportCorr.onclick = () => {
    triggerDownload(lastCorrBuffer, 'Output_Correlation.xlsx');
};


btnPlotCusum.onclick = () => {
    const selectedCols = Array.from(document.getElementById('cusum-cols').selectedOptions).map(opt => opt.value);
    if (selectedCols.length === 0 || selectedCols.length > 2) {
        showToast('Please select 1 or 2 columns for CUSUM.', 'error');
        return;
    }
    const payload = {
        start: document.getElementById('cusum-start').value,
        end: document.getElementById('cusum-end').value,
        dateCol: document.getElementById('cusum-date-col').value,
        cols: selectedCols
    };
    if (!payload.start || !payload.end) {
        showToast('Please select dates.', 'error');
        return;
    }
    setButtonLoading(btnPlotCusum, true); // Start spinner
    worker.postMessage({ type: 'runCusum', payload });
};
btnExportCusum.onclick = () => {
    triggerDownload(lastCusumBuffer, 'Output_CUSUM.xlsx');
};

btnAddPeriod.onclick = () => {
    const start = document.getElementById('ctrl-period-start').value;
    const end = document.getElementById('ctrl-period-end').value;
    if (!start || !end) {
        showToast('Please select a start and end date for the period.', 'error');
        return;
    }
    if (new Date(start) > new Date(end)) {
        showToast('Start date cannot be after end date.', 'error');
        return;
    }

    const startDMY = start.split('-').reverse().join('/');
    const endDMY = end.split('-').reverse().join('/');

    periods.push({ start: startDMY, end: endDMY });
    refreshPeriodList();
};

document.getElementById('btn-period-del').onclick = () => {
    const listbox = document.getElementById('ctrl-periods-list');
    const selectedIndex = listbox.selectedIndex;
    if (selectedIndex > -1) {
        periods.splice(selectedIndex, 1);
        refreshPeriodList();
    } else {
        showToast('Please select a period to delete.', 'error');
    }
};

function refreshPeriodList() {
    const listbox = document.getElementById('ctrl-periods-list');
    listbox.innerHTML = '';
    periods.forEach((p, i) => {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = `${i + 1}. ${p.start} → ${p.end}`;
        listbox.appendChild(option);
    });
}

btnPlotCtrl.onclick = () => {
    const conf = parseFloat(document.getElementById('ctrl-conf').value);
    if (isNaN(conf)) {
        showToast('Confidence must be a number.', 'error');
        return;
    }
    if (periods.length === 0) {
        showToast('Please add at least one date period.', 'error');
        return;
    }

    const payload = {
        col: document.getElementById('ctrl-col').value,
        dateCol: document.getElementById('ctrl-date-col').value,
        conf: conf,
        periods: periods,
        showLim: document.getElementById('ctrl-show-lim').checked,
        showAvg: document.getElementById('ctrl-show-avg').checked,
    };
    setButtonLoading(btnPlotCtrl, true); // Start spinner
    worker.postMessage({ type: 'runControlGraph', payload });
};
btnExportCtrl.onclick = () => {
    triggerDownload(lastCtrlBuffer, 'Output_Control.xlsx');
};

// --- Plotting Functions (using Plotly.js) ---

// === 💎 UPDATED Professional Global layout ===
const globalPlotLayout = {
    font: {
        family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
        size: 12,
        color: '#333333'
    },
    // Set a generous top margin for legend+title
    // The other margins will be set by automargin
    margin: { t: 100, l: 60, b: 60, r: 40 },
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#f4f7f6',
    title: {
        x: 0.5,
        xanchor: 'center',
        font: { size: 18, color: '#005a9c' }
    },
    // NEW Legend: Horizontal, centered, and ABOVE the title
    legend: {
        orientation: 'h',
        yanchor: 'bottom',
        y: 1.02, // 1 is top of plot, 1.02 is just above
        xanchor: 'center',
        x: 0.5
    },
    // NEW: Apply automargin to all axes to prevent cut-off labels
    xaxis: { automargin: true },
    yaxis: { automargin: true }
};

function drawCorrelationPlot(data) {
    const plotArea = document.getElementById('plot-corr');

    // NEW: Robustness check for empty data
    if (!data || !data.z || data.z.length === 0) {
        plotArea.innerHTML = '<div class="empty-plot-message">No correlation data to display.</div>';
        return;
    }
    // Clear empty message if data is valid
    plotArea.innerHTML = '';

    const plotData = [{
        z: data.z,
        x: data.x,
        y: data.y,
        type: 'heatmap',
        colorscale: 'RdBu',
        zmin: -1,
        zmax: 1,
        reversescale: true,
        hovertemplate: 'X: %{x}<br>Y: %{y}<br>Corr: %{z:.3f}<extra></extra>'
    }];

    const layout = {
        ...globalPlotLayout,
        title: { ...globalPlotLayout.title, text: 'Correlation Matrix' },
        // Override global automargin for heatmap's top-side X-axis
        xaxis: { automargin: true, side: 'top' },
        yaxis: { automargin: true },
        hovermode: 'closest'
    };
    delete layout.legend; // No legend for heatmaps

    Plotly.newPlot(plotArea, plotData, layout, {responsive: true});
}

function drawCusumPlot(data) {
    const plotArea = document.getElementById('plot-cusum');

    // NEW: Robustness check for empty data
    if (!data || !data.traces || data.traces.length === 0 || data.traces[0].data.length === 0) {
        plotArea.innerHTML = '<div class="empty-plot-message">No CUSUM data to display for this selection.</div>';
        return;
    }
    plotArea.innerHTML = ''; // Clear empty message

    const plotTraces = [];
    data.traces.forEach((trace, i) => {
        plotTraces.push({
            x: data.dateCol,
            y: trace.data,
            name: trace.name,
            type: 'scatter',
            mode: 'lines+markers',
            yaxis: `y${i+1}`,
            hovertemplate: `%{x}<br><b>${trace.name}</b>: %{y:.2f}<extra></extra>`
        });
    });

    const layout = {
        ...globalPlotLayout,
        title: { ...globalPlotLayout.title, text: 'CUSUM Chart' },
        xaxis: { ...globalPlotLayout.xaxis, title: data.dateColName },
        yaxis: { ...globalPlotLayout.yaxis, title: data.traces[0].name },
        hovermode: 'x unified',
        shapes: [{
            type: 'line',
            xref: 'paper', x0: 0, x1: 1,
            yref: 'y', y0: 0, y1: 0,
            line: { color: 'black', dash: 'dash' }
        }]
    };

    if (data.traces.length > 1) {
        layout.yaxis2 = {
            ...globalPlotLayout.yaxis, // Inherit automargin
            title: data.traces[1].name,
            overlaying: 'y',
            side: 'right'
        };
        // Add a bit more right margin for the second axis title
        layout.margin.r = 80;
    }

    Plotly.newPlot(plotArea, plotTraces, layout, {responsive: true});
}

function drawControlGraphPlot(data, colName, conf) {
    const plotArea = document.getElementById('plot-ctrl');

    // NEW: Robustness check for empty data
    if (!data || data.length === 0) {
        plotArea.innerHTML = '<div class="empty-plot-message">No Control Graph data to display for these periods.</div>';
        return;
    }
    plotArea.innerHTML = ''; // Clear empty message

    const plotTraces = [];

    data.forEach(trace => {
        let hovertemplate = '';
        if (trace.mode === 'markers') {
            hovertemplate = `%{x}<br><b>${colName}</b>: %{y:.2f}<br><i>${trace.name}</i><extra></extra>`;
        } else {
            hovertemplate = `<b>${trace.name}</b>: %{y:.2f}<extra></extra>`;
        }

        plotTraces.push({
            x: trace.x,
            y: trace.y,
            name: trace.name,
            type: 'scatter',
            mode: trace.mode,
            line: {
                color: trace.color,
                dash: trace.dash,
                width: trace.width
            },
            marker: {
                color: trace.color
            },
            hovertemplate: hovertemplate
        });
    });

    const layout = {
        ...globalPlotLayout,
        title: { ...globalPlotLayout.title, text: `${colName} – Control Graph (${conf}% CI)` },
        xaxis: { ...globalPlotLayout.xaxis, title: data[0].dateColName },
        yaxis: { ...globalPlotLayout.yaxis, title: colName },
        hovermode: 'x unified',
        grid: {rows: 1, columns: 1, pattern: 'independent'},
    };

    Plotly.newPlot(plotArea, plotTraces, layout, {responsive: true});
}