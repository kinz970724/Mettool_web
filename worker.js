// worker.js

// Load the Pyodide main script
importScripts('https://cdn.jsdelivr.net/pyodide/v0.25.1/full/pyodide.js');

// This will hold the fully initialized Pyodide instance.
let pyodide;

async function loadPyodideAndPackages() {
    // 1. Assign the loaded instance to the *global* pyodide variable
    pyodide = await loadPyodide();

    console.log('Loading micropip...');
    // 2. Load the package
    await pyodide.loadPackage('micropip');
    // 3. Import it into the Python scope
    const micropip = pyodide.pyimport('micropip');

    console.log('Installing Python packages...');
    await micropip.install([
        'pandas',
        'numpy',
        'scipy',
        'openpyxl',
        'xlrd'
    ]);
    console.log('All packages installed.');

    // 4. Load our external Python script (keeps code clean)
    console.log('Fetching Python script (mettool_web.py)...');
    const response = await fetch('./mettool_web.py');
    if (!response.ok) {
        throw new Error(`Failed to fetch mettool_web.py: ${response.status} ${response.statusText}`);
    }
    const mettoolScript = await response.text();

    console.log('Running Python script...');
    pyodide.runPython(mettoolScript);

    // 5. Signal that Pyodide is fully ready
    self.postMessage({ type: 'ready' });
}

// 6. Run the setup and catch any initialization errors
const pyodideReadyPromise = loadPyodideAndPackages().catch(error => {
    console.error('Pyodide initialization failed:', error);
    // Post an error message back to the main thread
    self.postMessage({ type: 'error', error: `Pyodide init failed: ${error.message}` });
});

// 7. Listen for messages from the main thread
self.onmessage = async (e) => {
    try {
        // Wait for the initialization to complete
        await pyodideReadyPromise;

        if (!pyodide) {
            throw new Error("Pyodide is not initialized. Check for startup errors.");
        }

        const { type, payload, buffer, sheetName } = e.data;
        let result;

        // Get handles to our Python functions
        const getSheetNamesPy = pyodide.globals.get('get_sheet_names_from_buffer');
        const loadDataPy = pyodide.globals.get('load_data_from_buffer');
        const getColumnsPy = pyodide.globals.get('get_columns');
        const clearDataPy = pyodide.globals.get('clear_data');
        const runCorrelationPy = pyodide.globals.get('run_correlation');
        const runCusumPy = pyodide.globals.get('run_cusum');
        const runControlGraphPy = pyodide.globals.get('run_control_graph');

        switch (type) {
            case 'getSheetNames': {
                result = getSheetNamesPy(buffer);
                postMessage({ type: 'sheetNames', payload: result.toJs() });
                result.destroy();
                break;
            }

            case 'loadData': {
                loadDataPy(buffer, sheetName);
                const columns = getColumnsPy();
                postMessage({ type: 'dataLoaded', payload: { columns: columns.toJs() } });
                columns.destroy();
                break;
            }

            case 'clearData': {
                clearDataPy();
                break;
            }

            case 'runCorrelation': {
                result = runCorrelationPy(payload);

                // --- VERIFIED FIX ---
                // 1. Get the JSON string and the file buffer proxy
                const plotDataJson = result.get('plot_data_json');
                const fileBufferPy = result.get('file_buffer');

                // 2. Parse the JSON to a pure JS object. Convert the buffer.
                const plotData = JSON.parse(plotDataJson);
                const fileBuffer = fileBufferPy.toJs();

                // 3. Post the message and transfer the buffer
                postMessage({
                    type: 'correlationResult',
                    payload: { plotData, fileBuffer }
                }, [fileBuffer.buffer]);

                // 4. Clean up proxies
                fileBufferPy.destroy();
                result.destroy();
                break;
            }

            case 'runCusum': {
                result = runCusumPy(payload);

                const plotDataJson = result.get('plot_data_json');
                const fileBufferPy = result.get('file_buffer');

                const plotData = JSON.parse(plotDataJson);
                const fileBuffer = fileBufferPy.toJs();

                postMessage({
                    type: 'cusumResult',
                    payload: { plotData, fileBuffer }
                }, [fileBuffer.buffer]);

                fileBufferPy.destroy();
                result.destroy();
                break;
            }

            case 'runControlGraph': {
                result = runControlGraphPy(payload);

                const plotDataJson = result.get('plot_data_json');
                const fileBufferPy = result.get('file_buffer');

                // Parse the JSON string into a pure JS array/object
                const plotData = JSON.parse(plotDataJson);
                const fileBuffer = fileBufferPy.toJs();

                postMessage({
                    type: 'controlGraphResult',
                    payload: {
                        plotData, // This is now a pure JS object
                        fileBuffer,
                        col: payload.col,
                        conf: payload.conf
                    }
                }, [fileBuffer.buffer]);

                fileBufferPy.destroy();
                result.destroy();
                break;
            }

            default:
                console.warn(`Unknown message type: ${type}`);
        }
    } catch (error) {
        console.error('Error in onmessage handler:', error);
        postMessage({ type: 'error', error: error.message });
    }
};

// 9. Global unhandled rejection listener
self.addEventListener('unhandledrejection', event => {
    console.error('Unhandled promise rejection in worker:', event.reason);
    postMessage({ type: 'error', error: `Unhandled Rejection: ${event.reason.message || event.reason}` });
});