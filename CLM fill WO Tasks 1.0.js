// ==UserScript==
// @name         CLM Fill Debrief Rows
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Automatically fill data in CLM
// @author       Denis Kiselev
// @match        *://elekta--svmxc.vf.force.com/apex/*ELK_WO_Task_Debrief*
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function () {
    'use strict';

    const DEBUG = true;
    // CSS classes for columns
    const CLASS_CD = "svmx-grid-cell-gridcolumn-1064"; // Check Description
    const CLASS_MR = "svmx-grid-cell-gridcolumn-1070"; // Manual Reference
    const CLASS_CID = "svmx-grid-cell-gridcolumn-1071"; // Check ID
    const CLASS_TS = "svmx-grid-cell-gridcolumn-1065"; // Task Status
    const CLASS_TC = "svmx-grid-cell-gridcolumn-1066"; // Technician Comments

    let comparisonWindow = null;
    let doc = null;
    let cloudData = [];
    let taskStatusOptions = [];
    let TARGET_ROWS = [];

    // Constants for messages
    const messages = {
        fileRequired: "Please select a CSV file to upload.",
        tableNotFound: "Table with ID 'gridview-1080' not found.",
        noRows: "Table contains no rows.",
        dataExtractionFailed: "Failed to extract data from table.",
        noDifferences: "No differences found.",
        noMissingInCloud: "No missing rows in Cloud found.",
        noMissingInFile: "No missing rows in File found.",
        noMatching: "No matching rows found.",
        selectStatus: "Please select a Task Status.",
        copyNotImplemented: "Copy to CLM functionality not yet implemented.",
        doExactly: "Do exactly what was requested, no more, no less."
    };

    // State Management System
    const StateManager = {
        // Possible states
        States: {
            IDLE: 'IDLE',
            PROCESSING: 'PROCESSING',
            WINDOW_OPEN: 'WINDOW_OPEN',
            COPYING: 'COPYING',
            FILLING: 'FILLING',
            ERROR: 'ERROR'
        },
        currentState: 'IDLE',
        operationQueue: [],
        operationLock: false,
        progress: {total: 0, current: 0, status: ''},
        handlers: {onStateChange: null, onError: null, onProgress: null},
        hasError: false, // Flag to track if we're already handling an error

        init(callbacks = {}) {
            // Initialize the state manager
            this.handlers = { ...this.handlers, ...callbacks };
            this.currentState = this.States.IDLE;
            this.hasError = false;
            this.updateUI();
        },

        // Validate state transition
        setState(newState, details = {}) {
            // Don't process new state changes if we're already handling an error
            if (this.hasError && newState !== this.States.IDLE) {
                return false;
            }

            const oldState = this.currentState;
            if (!this.isValidTransition(oldState, newState)) {
                // Only handle the error if we're not already in an error state
                if (!this.hasError) {
                    this.handleError(new Error(`Invalid state transition from ${oldState} to ${newState}`));
                }
                return false;
            }

            this.currentState = newState;
            this.updateUI();

            if (this.handlers.onStateChange) {
                this.handlers.onStateChange(newState, oldState, details);
            }

            // Reset error flag when returning to IDLE
            if (newState === this.States.IDLE) {
                this.hasError = false;
            }

            return true;
        },

        // Validate state transitions
        isValidTransition(fromState, toState) {
            // Any state can transition to ERROR
            if (toState === this.States.ERROR) return true;

            // From ERROR you can only go to IDLE
            if (fromState === this.States.ERROR) {
                return toState === this.States.IDLE;
            }

            const validTransitions = {
                [this.States.IDLE]: [this.States.PROCESSING],
                [this.States.PROCESSING]: [this.States.WINDOW_OPEN, this.States.IDLE],
                [this.States.WINDOW_OPEN]: [this.States.COPYING, this.States.IDLE],
                [this.States.COPYING]: [this.States.FILLING],
                [this.States.FILLING]: [this.States.IDLE]
            };

            return validTransitions[fromState]?.includes(toState) ?? false;
        },

        // Update UI based on current state
        updateUI() {
            const button = document.getElementById('fillRowsButton');
            if (!button) return;

            const stateConfig = {
                [this.States.IDLE]: { text: 'Select Rows', disabled: false, style: 'default' },
                [this.States.PROCESSING]: { text: 'Processing...', disabled: true, style: 'processing' },
                [this.States.WINDOW_OPEN]: { text: 'Comparison Window Open', disabled: true, style: 'window-open' },
                [this.States.COPYING]: { text: 'Copying Data...', disabled: true, style: 'processing' },
                [this.States.FILLING]: { text: 'Filling Rows...', disabled: true, style: 'filling' },
                [this.States.ERROR]: { text: 'Error Occurred', disabled: false, style: 'error' }
            };

            const config = stateConfig[this.currentState];

            button.textContent = config.text;
            button.disabled = config.disabled;
            button.className = `fill-rows-button ${config.style}`;
            button.style.opacity = config.disabled ? '0.6' : '1';
            button.style.cursor = config.disabled ? 'not-allowed' : 'pointer';
        },

        // Error handling
        handleError(error) {
            // Only handle the first error
            if (this.hasError) {
                return;
            }

            this.hasError = true;
            console.error('Error:', error.message);

            if (this.handlers.onError) {
                this.handlers.onError(error);
            }

            // Set state to ERROR and show alert only for the first error
            this.setState(this.States.ERROR, { error });
            alert(error.message);

            // Reset state after a short delay
            setTimeout(() => {
                this.setState(this.States.IDLE);
            }, 2000);
        },

        // Progress tracking
        updateProgress(current, total, status = '') {
            if (this.hasError) return;

            this.progress = { current, total, status };

            if (this.handlers.onProgress) {
                this.handlers.onProgress(this.progress);
            }
        }
    };

    // ====== Utility / Wait ======
    function wait(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * Counts rows in a given table, optionally excluding those with colspan.
     * @param {string} tableId - ID of the table.
     * @param {boolean} excludeColspan - Whether to exclude rows with colspan.
     * @returns {number} - Number of rows.
     */
    function countTableRows(tableId, excludeColspan = false) {
        const tbody = doc.querySelector(`#${tableId} tbody`);
        if (!tbody) return 0;

        const rows = Array.from(tbody.querySelectorAll('tr'));
        if (excludeColspan) {
            return rows.filter(tr => !tr.querySelector('td[colspan]')).length;
        }
        return rows.length;
    }

    /**
     * Updates the counters for tables that display differences, missing rows, and matching data.
     */
    function updateCounters() {
        const diffCount = countTableRows('differences-table', true) / 2;
        updateTableCount('differences-title', diffCount);

        const missingFileCount = countTableRows('missing-rows-file-table', true);
        updateTableCount('missing-file-title', missingFileCount);

        const matchingCount = countTableRows('matching-table', true);
        updateTableCount('matching-title', matchingCount);

        const missingCloudCount = countTableRows('missing-rows-cloud-table', true);
        updateTableCount('missing-cloud-title', missingCloudCount);

        const originalCount = countTableRows('data-table');
        updateOriginalDataCount(originalCount);
    }

    /**
     * Sets the Task Status of a specific cell by clicking and selecting the correct option.
     * @param {HTMLElement} cell - The cell element to interact with.
     * @param {string} status - The desired task status to select.
     */
    async function setTaskStatus(cell, status) {
        cell.click();
        await wait(30);

        const picklistWrapper = document.querySelector("#sfm-picklistcelleditor-1049-triggerWrap");
        if (!picklistWrapper) {
            console.error("Task Status dropdown container not found.");
            return;
        }

        const triggerArrow = picklistWrapper.querySelector(".svmx-form-trigger.svmx-form-arrow-trigger");
        if (triggerArrow) {
            triggerArrow.click();
            await wait(30);

            const boundList = document.querySelector(".svmx-boundlist-list-ct");
            if (!boundList) {
                console.error("Task Status options list not found.");
                return;
            }

            const options = boundList.querySelectorAll('.svmx-boundlist-item');
            let optionFound = false;
            for (let option of options) {
                if (option.textContent.trim() === status) {
                    option.click();
                    optionFound = true;
                    break;
                }
            }

            if (!optionFound) {
                console.warn(`Option "${status}" not found in Task Status dropdown.`);
            }
        } else {
            console.error("Task Status dropdown arrow not found.");
        }

        await wait(30);
    }

    /**
     * Adds a "Fill Rows" button to the page and initializes the State Manager.
     */
    function addButton() {
        const button = document.createElement('button');
        button.id = 'fillRowsButton';
        button.textContent = "Fill Rows";
        button.style.position = 'absolute';
        button.style.left = '50%';
        button.style.top = '2px';
        button.style.padding = '10px';
        button.style.transform = 'translateX(-50%)';
        button.style.padding = '8px 20px';
        button.style.setProperty('background', '#a8e4f4');
        button.style.border = 'none';
        button.style.borderRadius = '5px';
        button.style.cursor = 'pointer';
        button.style.zIndex = '1000';
        button.style.color = '#000';
        button.style.boxShadow = '0 2px 4px rgba(0,0,0,0.2)';
        button.style.fontSize = '14px';

        button.addEventListener('click', processRows);
        document.body.appendChild(button);
        StateManager.init();
    }

    /**
     * Sets technician comments in the Technician Comments cell.
     * @param {HTMLElement} cell - The cell element to interact with.
     * @param {string} comments - The comments to input.
     */
    async function setTechnicianComments(cell, comments) {
        cell.click();
        await wait(30);

        const textarea = document.querySelector("#sfm-textarea-1050-inputEl");
        if (textarea) {
            textarea.value = comments;
            textarea.dispatchEvent(new Event('input', { bubbles: true }));
            await wait(30);
            // Click outside the textarea to save
            document.body.click();
        } else {
            console.error("Technician Comments textarea not found.");
        }

        await wait(30);
    }

    /**
     * Object that validates the file being uploaded, checks headers, size, etc.
     */
    const FileProcessor = {
        ALLOWED_EXTENSIONS: ['.csv'],
        MAX_FILE_SIZE: 5 * 1024 * 1024, // 5MB
        REQUIRED_HEADERS: ['Check Description', 'Task Status', 'Technician Comments', 'Manual Reference', 'Check ID'],

        validateFileExtension(filename) {
            const ext = '.' + filename.split('.').pop().toLowerCase();
            if (!this.ALLOWED_EXTENSIONS.includes(ext)) {
                throw new Error(`Invalid file type. Allowed types: ${this.ALLOWED_EXTENSIONS.join(', ')}`);
            }
        },

        validateFileSize(size) {
            if (size > this.MAX_FILE_SIZE) {
                const maxSizeMB = this.MAX_FILE_SIZE / (1024 * 1024);
                throw new Error(`File size exceeds ${maxSizeMB}MB limit`);
            }
        },

        validateHeaders(headers) {
            const missingHeaders = this.REQUIRED_HEADERS.filter(
                required => !headers.includes(required)
            );
            if (missingHeaders.length > 0) {
                throw new Error(`Missing required headers: ${missingHeaders.join(', ')}`);
            }
        },

        readFileAsText(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => resolve(e.target.result);
                reader.onerror = (e) => reject(new Error('Failed to read file'));
                reader.readAsText(file);
            });
        }
    };

    // Data handling functions

    /**
     * Parses CSV text into an array of objects.
     * @param {string} text - The CSV content as a string.
     * @returns {Array<Object>} - Array of row objects.
     */
    function parseCSV(text) {
        const lines = text.trim().split(/\r?\n/);
        const headers = lines[0].split(',').map(header => header.replace(/(^"|"$)/g, '').trim());
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = lines[i]
                .split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/)
                .map(value => value.replace(/(^"|"$)/g, '').trim());
            const row = {};
            headers.forEach((header, index) => {
                row[header] = values[index] || "";
            });
            data.push(row);
        }

        return data;
    }

    /**
     * Compares file data with cloud data to find differences, matching rows, and missing rows.
     * @param {Array<Object>} fileData - Array of row objects from the file.
     * @param {Array<Object>} cloudData - Array of row objects from the cloud.
     */
    function compareData(fileData, cloudData) {
        const differences = [];
        const matching = [];
        const missingInCloud = [];
        const missingInFile = [];
        const cloudDataCopy = [...cloudData];

        // Iterate through each row from the file
        fileData.forEach(fileRow => {
            let matched = false;

            // Search for matches in cloud data
            for (let i = 0; i < cloudDataCopy.length; i++) {
                const cloudRow = cloudDataCopy[i];

                const cdMatch = fileRow["Check Description"] === cloudRow["Check Description"];
                const mdMatch = fileRow["Manual Reference"] === cloudRow["Manual Reference"];
                const cidMatch = fileRow["Check ID"] === cloudRow["Check ID"];

                // Check for full match
                if (cdMatch && mdMatch && cidMatch) {
                    matching.push(fileRow);
                    cloudDataCopy.splice(i, 1);
                    matched = true;
                    break;
                }

                // Check for partial match
                if (
                    (cdMatch && mdMatch && !cidMatch) ||
                    (cdMatch && !mdMatch && cidMatch) ||
                    (cdMatch && !mdMatch && !cidMatch) ||
                    (!cdMatch && mdMatch && cidMatch)
                ) {
                    differences.push({
                        source: "File",
                        ...fileRow
                    });
                    differences.push({
                        source: "Cloud",
                        ...cloudRow
                    });
                    cloudDataCopy.splice(i, 1);
                    matched = true;
                    break;
                }
            }

            // If no match found
            if (!matched) {
                missingInCloud.push(fileRow);
            }
        });

        // Remaining cloud data not found in file
        cloudDataCopy.forEach(cloudRow => {
            missingInFile.push(cloudRow);
        });

        // Display the results
        displayDifferences(differences);
        displayMissingRowsInCloud(missingInCloud);
        displayMissingRowsInFile(missingInFile);
        displayMatching(matching);

        // Show the comparison results section
        doc.getElementById('comparison-results').classList.remove('hidden');

        // Enable or disable the "Copy Data to CLM" button based on matching rows
        doc.getElementById('copy-data-button').disabled = matching.length === 0;
    }

    /**
     * Handles file upload, parses and compares data.
     * @param {Event} e - The change event from the file input.
     */
    async function handleFileUpload(e) {
        try {
            const file = e.target.files[0];
            if (!file) {
                throw new Error(messages.fileRequired);
            }

            FileProcessor.validateFileExtension(file.name);
            FileProcessor.validateFileSize(file.size);

            const fileContent = await FileProcessor.readFileAsText(file);

            const firstLine = fileContent.split('\n')[0];
            const headers = firstLine.split(',').map(h => h.trim().replace(/"/g, ''));
            FileProcessor.validateHeaders(headers);

            const fileData = parseCSV(fileContent);

            if (!fileData || fileData.length === 0) {
                throw new Error('File contains no data');
            }

            compareData(fileData, cloudData);

        } catch (error) {
            StateManager.handleError(error);
            alert(error.message);
            e.target.value = '';
        }
    }

    // Utility functions

    /**
     * Escapes HTML characters in a string.
     * @param {string} str - The string to escape.
     * @returns {string} - The escaped string.
     */
    function escapeHTML(str) {
        if (!str) return "";
        return str
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    /**
     * Updates the numeric counter in the given table's title.
     * @param {string} titleElementId - The ID of the title element.
     * @param {number} count - The new count value to display.
     */
    function updateTableCount(titleElementId, count) {
        const titleElement = doc.getElementById(titleElementId);
        if (titleElement) {
            const textContent = titleElement.textContent;
            const dashIndex = textContent.indexOf(': ');
            const baseTitle = dashIndex !== -1 ? textContent.substring(0, dashIndex) : textContent;
            titleElement.textContent = `${baseTitle}: ${count}`;
        }
    }

    /**
     * Highlights differences between two values.
     * @param {string} val1 - Value from the first source (File or Cloud).
     * @param {string} val2 - Value from the second source (Cloud or File).
     * @param {boolean} isSource1 - Indicates which source the value belongs to.
     * @returns {string} - HTML string with highlighted difference if any.
     */
    function highlightDifference(val1, val2, isSource1) {
        // Show the correct value based on the source
        const valueToShow = isSource1 ? val1 : val2;
        // Add a check for differences
        if (val1 !== val2) {
            // Show the correct value based on the source
            return `<span class="mismatch">${escapeHTML(valueToShow)}</span>`;
        }
        return escapeHTML(valueToShow);
    }

    /**
     * Returns the Task Status options.
     * @returns {Array<string>}
     */
    function getTaskStatusOptions() {
        return taskStatusOptions;
    }

    // Display functions

    /**
     * Displays rows with differences in the Differences table.
     * @param {Array<Object>} differences - Array of differing row objects.
     */
    function displayDifferences(differences) {
        const tbody = doc.querySelector('#differences-table tbody');
        const fragment = doc.createDocumentFragment();

        if (differences.length === 0) {
            const tr = doc.createElement('tr');
            tr.innerHTML = `<td colspan='6'>${messages.noDifferences}</td>`;
            fragment.appendChild(tr);
            // Update the counter to 0 for an empty table
            updateTableCount('differences-title', 0);
        } else {
            // Assume differences contains pairs: first row File, second Cloud
            for (let i = 0; i < differences.length; i += 2) {
                const fileRow = differences[i];
                const cloudRow = differences[i + 1];

                if (!fileRow || !cloudRow) {
                    console.warn(`Mismatched row pairs: file ${fileRow}, cloud ${cloudRow}`);
                    continue; // Skip incomplete pairs
                }

                // Function to highlight differences
                function highlightDifference(fileValue, cloudValue) {
                    return fileValue !== cloudValue
                        ? `<span class="mismatch">${escapeHTML(fileValue)}</span>`
                        : escapeHTML(fileValue);
                }

                // Create row for File
                const trFile = doc.createElement('tr');
                trFile.innerHTML = `
                    <td>${highlightDifference(fileRow["Check Description"], cloudRow["Check Description"])}</td>
                    <td>${escapeHTML(fileRow["Task Status"])}</td>
                    <td>${escapeHTML(fileRow["Technician Comments"])}</td>
                    <td>${highlightDifference(fileRow["Manual Reference"], cloudRow["Manual Reference"])}</td>
                    <td>${highlightDifference(fileRow["Check ID"], cloudRow["Check ID"])}</td>
                    <td style="display: flex; align-items: center; justify-content: space-between;">
                        <!-- Icon -->
                        <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="#2196F3" stroke-width="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                        </svg>
                        <button class="validate-button" style="display: flex; padding: 2px 4px; align-items: center; justify-content: center;">
                            <svg viewBox="0 0 48 24" width="24" height="12" fill="none" stroke="#FFFFFF" stroke-width="2">
                                <title>Copy rows to the Matching rows</title>
                                <path d="M3.8 3c0 5.3 4.6 7.3 8.1 7.5v-4.4l9.3 7.3-9.3 7.2v-3.9q-.9 0-1.7-.2-0.8-.2-1.5-.5-0.8-.4-1.5-.9t-1.2-1.1q-3.5-3.8-3-11z"/>
                                <path d="M41.8 7.2q-.1 0-.3.1h-.5q-.1.1-.3.1-.3-.6-.9-1.1t-1.1-.8q-.6-.4-1.3-.6t-1.4-.2q-1.1.1-2.1.5-.9.4-1.7 1.1-0.8.7-1.3 1.7-0.5.9-0.6 2-.2-.1-.3-.1t-.3-.1h-.3q-.2-.1-.4-.1-.7.1-1.3.4t-1.2.8q-.5.5-.7 1.2-.3.7-.2 1.4-.1.7.2 1.4t.7 1.2q.5.5 1.2.8t1.3.4h12.8q1-.1 1.8-.5.9-.4 1.6-1.1t.9-1.6q.4-.9.3-1.8.1-1-.3-1.9t-.9-1.6q-.7-.7-1.6-1.1t-1.8-.5z"/>
                            </svg>
                        </button>
                    </td>
                `;

                // Create row for Cloud
                const trCloud = doc.createElement('tr');
                trCloud.innerHTML = `
                    <td>${highlightDifference(cloudRow["Check Description"], fileRow["Check Description"])}</td>
                    <td>${escapeHTML(cloudRow["Task Status"])}</td>
                    <td>${escapeHTML(cloudRow["Technician Comments"])}</td>
                    <td>${highlightDifference(cloudRow["Manual Reference"], fileRow["Manual Reference"])}</td>
                    <td>${highlightDifference(cloudRow["Check ID"], fileRow["Check ID"])}</td>
                    <td>
                        <!-- Source Icon Cloud -->
                        <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="#00A1E0" stroke-width="2">
                            <path d="M17.8 8.2q-0.1 0-0.3 0.1-0.1 0-0.2 0-0.2 0-0.3 0-0.1 0.1-0.3 0.1-0.3-0.6-0.9-1.1-0.5-0.5-1.1-0.8-0.6-0.4-1.3-0.6-0.7-0.2-1.4-0.2-1.1 0.1-2.1 0.5-0.9 0.4-1.7 1.1-0.8 0.7-1.3 1.7-0.5 0.9-0.6 2-0.2-0.1-0.3-0.1-0.2-0.1-0.3-0.1-0.2 0-0.3 0-0.2-0.1-0.4-0.1-0.7 0.1-1.3 0.4-0.7 0.3-1.2 0.8-0.5 0.5-0.7 1.2-0.3 0.7-0.2 1.4-0.1 0.7 0.2 1.4 0.2 0.7 0.7 1.2 0.5 0.5 1.2 0.8 0.6 0.3 1.3 0.4h12.8q1-0.1 1.8-0.5 0.9-0.4 1.6-1.1 0.6-0.7 0.9-1.6 0.4-0.9 0.3-1.8 0.1-1-0.3-1.9-0.3-0.9-0.9-1.6-0.7-0.7-1.6-1.1-0.8-0.4-1.8-0.5z"/>
                        </svg>
                    </td>
                `;

                // Append rows to the fragment
                fragment.appendChild(trFile);
                fragment.appendChild(trCloud);

                // Separator between pairs
                const separator = doc.createElement('tr');
                separator.innerHTML = '<td colspan="6" style="border-bottom: 2px solid #ccc"></td>';
                fragment.appendChild(separator);

                // Add event listeners for the Validate button
                const validateButton = trFile.querySelector('.validate-button');
                validateButton.addEventListener('click', () => {
                    const selectedStatus = trFile.cells[1].textContent.trim();
                    const comments = trFile.cells[2].textContent.trim();

                    // Check for necessary data
                    if (!selectedStatus) {
                        alert(messages.selectStatus);
                        return;
                    }

                    // Create a validated row with data from Cloud and File
                    const validatedRow = {
                        "Check Description": cloudRow["Check Description"], // from Cloud
                        "Task Status": selectedStatus,                       // from File
                        "Technician Comments": comments,                    // from File
                        "Manual Reference": cloudRow["Manual Reference"],   // from Cloud
                        "Check ID": cloudRow["Check ID"]                   // from Cloud
                    };

                    // Move to Matching Rows
                    addToMatchingRows(validatedRow);

                    // Remove current rows from the Differences table
                    trFile.remove();
                    trCloud.remove();
                    separator.remove();

                    // Update the differences counter after removal
                    const diffCount = countDifferences();
                    updateTableCount('differences-title', diffCount);

                    // If the table is empty after removal
                    if (tbody.querySelectorAll('tr').length === 0) {
                        const newTr = doc.createElement('tr');
                        newTr.innerHTML = `<td colspan='6'>${messages.noDifferences}</td>`;
                        tbody.appendChild(newTr);
                        updateTableCount('differences-title', 0);
                    }

                    // Update counters for both differences and matching tables
                    const matchingRows = doc.querySelectorAll('#matching-table tbody tr').length;
                    updateTableCount('matching-title', matchingRows);
                });
            }
        }

        tbody.innerHTML = '';
        tbody.appendChild(fragment);

        // We only count after all rows are added
        const diffCount = countDifferences();
        updateTableCount('differences-title', diffCount);
    }

    /**
     * Displays rows that are missing in Cloud.
     * @param {Array<Object>} missingRows - Array of row objects missing in the Cloud.
     */
    function displayMissingRowsInCloud(missingRows) {
        const tbody = doc.querySelector('#missing-rows-cloud-table tbody');
        const fragment = doc.createDocumentFragment();

        if (missingRows.length === 0) {
            const tr = doc.createElement('tr');
            tr.innerHTML = `<td colspan='6'>${messages.noMissingInCloud}</td>`;
            fragment.appendChild(tr);
        } else {
            missingRows.forEach(row => {
                const tr = doc.createElement('tr');
                tr.innerHTML = `
                    <td>${escapeHTML(row["Check Description"])}</td>
                    <td>${escapeHTML(row["Task Status"])}</td>
                    <td>${escapeHTML(row["Technician Comments"])}</td>
                    <td>${escapeHTML(row["Manual Reference"])}</td>
                    <td>${escapeHTML(row["Check ID"])}</td>
                    <td>
                        <!-- Source Icon File -->
                        <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="#2196F3" stroke-width="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                        </svg>
                    </td>
                `;
                fragment.appendChild(tr);
            });
        }

        tbody.innerHTML = '';
        tbody.appendChild(fragment);

        const missingCloudTitle = doc.getElementById('missing-cloud-title');
        if (missingCloudTitle) {
            missingCloudTitle.textContent = `Rows Not Found in Cloud: ${missingRows.length}`;
        }
    }

    /**
     * Displays rows that are missing in File.
     * @param {Array<Object>} missingRows - Array of row objects missing in File.
     */
    function displayMissingRowsInFile(missingRows) {
        const tbody = doc.querySelector('#missing-rows-file-table tbody');
        const fragment = doc.createDocumentFragment();

        if (missingRows.length === 0) {
            const tr = doc.createElement('tr');
            tr.innerHTML = `<td colspan='6'>${messages.noMissingInFile}</td>`;
            fragment.appendChild(tr);
        } else {
            const taskStatusOptions = getTaskStatusOptions();

            missingRows.forEach(row => {
                const tr = doc.createElement('tr');
                tr.innerHTML = `
                    <td>${escapeHTML(row["Check Description"])}</td>
                    <td class="editable-cell">
                        <select class="task-status-select">
                            <option value="">Select Status</option>
                            ${taskStatusOptions
                                .map(
                                    option =>
                                        `<option value="${escapeHTML(option)}">${escapeHTML(option)}</option>`
                                )
                                .join('')}
                        </select>
                    </td>
                    <td class="editable-cell">
                        <input type="text" class="tech-comments-input" placeholder="Enter comments">
                    </td>
                    <td>${escapeHTML(row["Manual Reference"])}</td>
                    <td>${escapeHTML(row["Check ID"])}</td>
                    <td>
                        <button class="validate-button" disabled>Validate</button>
                    </td>
                `;

                const validateButton = tr.querySelector('.validate-button');
                const statusSelect = tr.querySelector('.task-status-select');

                // Enable the Validate button when a status is selected
                statusSelect.addEventListener('change', function () {
                    if (this.value) {
                        validateButton.disabled = false;
                    } else {
                        validateButton.disabled = true;
                    }
                });

                validateButton.addEventListener('click', function () {
                    // Retrieve the selected Task Status and Technician Comments
                    const selectedStatus = tr.querySelector('.task-status-select').value.trim();
                    const comments = tr.querySelector('.tech-comments-input').value.trim();

                    // Validate that a Task Status has been selected
                    if (!selectedStatus) {
                        alert(messages.selectStatus);
                        return;
                    }

                    // Update row data
                    row["Task Status"] = selectedStatus;
                    row["Technician Comments"] = comments;

                    // Move row to Matching
                    addToMatchingRows(row);
                    tr.remove();

                    // If the table is empty after removal
                    if (tbody.children.length === 0) {
                        const newTr = doc.createElement('tr');
                        newTr.innerHTML = `<td colspan='6'>${messages.noMissingInFile}</td>`;
                        tbody.appendChild(newTr);
                    }

                    // Count how many rows remain (excluding any single message row that has colspan)
                    const remainingRows = Array.from(
                        doc.querySelectorAll('#missing-rows-file-table tbody tr')
                    ).filter(tr => !tr.querySelector('td[colspan="6"]')).length;

                    updateTableCount('missing-file-title', remainingRows);
                });

                fragment.appendChild(tr);
            });
        }

        tbody.innerHTML = '';
        tbody.appendChild(fragment);

        const missingFileTitle = doc.getElementById('missing-file-title');
        if (missingFileTitle) {
            missingFileTitle.textContent = `Rows Not Found in File: ${missingRows.length}`;
        }
    }

    /**
     * Counts how many difference pairs (i.e., those containing a validate-button).
     * @returns {number}
     */
    function countDifferences() {
        const tbody = doc.querySelector('#differences-table tbody');
        if (!tbody) return 0;
        const validateButtons = tbody.querySelectorAll('button.validate-button');
        return validateButtons.length;
    }

    /**
     * Displays matching rows in the Matching table.
     * @param {Array<Object>} matching - Array of row objects that match in both File and Cloud.
     */
    function displayMatching(matching) {
        const tbody = doc.querySelector('#matching-table tbody');
        const fragment = doc.createDocumentFragment();

        if (matching.length === 0) {
            const tr = doc.createElement('tr');
            tr.innerHTML = `<td colspan='5'>${messages.noMatching}</td>`;
            fragment.appendChild(tr);
        } else {
            matching.forEach(row => {
                const tr = doc.createElement('tr');
                tr.innerHTML = `
                    <td>${escapeHTML(row["Check Description"])}</td>
                    <td>${escapeHTML(row["Task Status"])}</td>
                    <td>${escapeHTML(row["Technician Comments"])}</td>
                    <td>${escapeHTML(row["Manual Reference"])}</td>
                    <td>${escapeHTML(row["Check ID"])}</td>
                `;
                fragment.appendChild(tr);
            });
        }

        tbody.innerHTML = '';
        tbody.appendChild(fragment);

        const matchingTitle = doc.getElementById('matching-title');
        if (matchingTitle) {
            matchingTitle.textContent = `Matching Rows: ${matching.length}`;
        }
        updateTableCount('matching-title', matching.length);
    }

    /**
     * Adds a single row object to the Matching rows table.
     * @param {Object} row - The row object to add.
     */
    function addToMatchingRows(row) {
        const tbody = doc.querySelector('#matching-table tbody');

        // If table has a "No matching rows found" message
        if (tbody.children.length === 1 && tbody.children[0].children[0].colSpan) {
            tbody.innerHTML = '';
        }

        const tr = doc.createElement('tr');
        tr.innerHTML = `
            <td>${escapeHTML(row["Check Description"])}</td>
            <td>${escapeHTML(row["Task Status"])}</td>
            <td>${escapeHTML(row["Technician Comments"])}</td>
            <td>${escapeHTML(row["Manual Reference"])}</td>
            <td>${escapeHTML(row["Check ID"])}</td>
        `;
        tbody.appendChild(tr);

        // Update count in matching table
        const matchingRows = doc.querySelectorAll('#matching-table tbody tr').length;
        updateTableCount('matching-title', matchingRows);
    }

    comparisonWindow = null;

    /**
     * Main logic to process rows from the parent window and open a new comparison window.
     */
    async function processRows() {
        if (!StateManager.setState(StateManager.States.PROCESSING)) {
            return;
        }

        try {
            const newWindow = window.open('', '_blank', 'width=1000,height=800,scrollbars=yes,resizable=yes');
            const parentWindow = window;

            if (!newWindow) {
                console.error("Failed to open new window. Check popup blocker settings.");
                alert("Failed to open a new window. Please disable the pop-up blocker and try again.");
                StateManager.setState(StateManager.States.IDLE);
                return;
            }

            StateManager.setState(StateManager.States.WINDOW_OPEN);

            // Extract table from the current page
            const table = document.getElementById("gridview-1080");
            if (!table) {
                console.error(messages.tableNotFound);
                return;
            }

        // Extract table rows
            const rows = table.querySelectorAll("tr");
            if (!rows.length) {
                console.error(messages.noRows);
                return;
            }

            // Collect data from the current page
            cloudData = [];
            rows.forEach(row => {
                const checkDescription = row.querySelector(`.${CLASS_CD}`);
                const taskStatus = row.querySelector(`.${CLASS_TS}`);
                const technicianComments = row.querySelector(`.${CLASS_TC}`);
                const manualReference = row.querySelector(`.${CLASS_MR}`);
                const checkID = row.querySelector(`.${CLASS_CID}`);

                // Ensure at least one cell exists to avoid empty rows
                if (
                    checkDescription ||
                    taskStatus ||
                    technicianComments ||
                    manualReference ||
                    checkID
                ) {
                    cloudData.push({
                        "Check Description": checkDescription ? checkDescription.textContent.trim() : "",
                        "Task Status": taskStatus ? taskStatus.textContent.trim() : "",
                        "Technician Comments": technicianComments ? technicianComments.textContent.trim() : "",
                        "Manual Reference": manualReference ? manualReference.textContent.trim() : "",
                        "Check ID": checkID ? checkID.textContent.trim() : ""
                    });
                }
            });

            if (!cloudData.length) {
                console.error(messages.dataExtractionFailed);
                return;
            }

            // Attempt to gather Task Status options
            try {
                const boundList = document.querySelector('.svmx-boundlist-list-ct');
                if (boundList) {
                    // If dropdown is present in the DOM
                    taskStatusOptions = Array.from(boundList.querySelectorAll('li')).map(li =>
                        li.textContent.trim()
                    );
                } else {
                    // Fallback defaults
                    console.log('Selector not found (select at least one TaskStatus value), using fallback options');
                    taskStatusOptions = [
                        'Pass',
                        'Done by Customer',
                        'Not Done – Customer Request',
                        'Not Done – Not Applicable',
                        'Fail',
                        'Not Started'
                    ];
                }
            } catch (error) {
                console.log('Error getting task options:', error);
            }

            // Write content to new window
            newWindow.document.open();
            newWindow.document.write(getHTMLContent(taskStatusOptions));
            newWindow.document.close();

            // Initialize after child document is loaded
            newWindow.onload = function() {
                initializeApp(newWindow, cloudData, parentWindow);
            };
        } catch (error) {
            StateManager.handleError(error);
        }
    } // end of async function processRows()

    /**
     * Function for initializing the application in the new comparison window.
     * @param {Window} win - The newly opened window object.
     * @param {Array<Object>} data - The extracted cloud data.
     * @param {Window} parentWindow - The parent window object.
     */
    function initializeApp(win, data, parentWindow) {
        doc = win.document;
        comparisonWindow = win;

        // Initialize event listeners
        initializeEventListeners(parentWindow);
        if (DEBUG) {
            compareData(getFileData(), data); // Debug
        } else {
            // compareData(fileData(), data); // Non-debug usage
        }
    }

    /**
     * Returns the full HTML content for the new comparison window.
     * @param {Array<string>} taskStatusOptions - The list of Task Status options to populate.
     * @returns {string} - The complete HTML markup for the new window.
     */
    function getHTMLContent(taskStatusOptions) {
        const taskStatusJSON = JSON.stringify(taskStatusOptions);
        let htmlContent = `
            <html>
            <head>
                <title>Data Comparison Tool</title>
                <style>
                    body {
                        font-family: 'Arial', sans-serif;
                        font-size: 14px;
                        margin: 20px;
                    }
                    h1, h2, h3 {
                        color: #333;
                    }
                    .section {
                        margin-bottom: 40px;
                    }
                    .hidden {
                        display: none;
                    }
                    table {
                        border-collapse: collapse;
                        width: 100%;
                        font-family: 'Arial', sans-serif;
                        table-layout: fixed;
                        border: 1px solid #E0E0E0;
                        font-size: 14px;
                        margin-bottom: 20px;
                    }
                    th, td {
                        border: 1px solid #E0E0E0;
                        border-right: 1px solid #5cc2fc;
                        border-left: 1px solid #5cc2fc;
                        padding: 4px 6px;
                        overflow: hidden;
                        position: relative;
                        box-sizing: border-box;
                    }
                    td {
                        font-size: 12px;
                    }
                    th {
                        background: linear-gradient(to bottom, #D7F2F9, #B7E1EC);
                        color: #000;
                        font-weight: bold;
                        text-align: left;
                        position: relative;
                        min-width: 30px;
                    }
                    th .resize-handle {
                        position: absolute;
                        right: 0;
                        top: 0;
                        width: 5px;
                        height: 100%;
                        cursor: col-resize;
                        user-select: none;
                        background: rgba(0, 0, 0, 0.1);
                        z-index: 1;
                    }

                    th .resize-handle:hover {
                        background: rgba(0, 0, 0, 0.3);
                    }

                    th:nth-child(1) { width: 320px; }
                    th:nth-child(2) { width: 128px; }
                    th:nth-child(3) { width: 144px; }
                    th:nth-child(4) { width: 80px; }
                    th:nth-child(5) { width: 64px; }
                    th:nth-child(6) { width: 64px; }

                    tr:nth-child(even) {
                        background-color: #D7F2F9;
                    }
                    tr:nth-child(odd) {
                        background-color: #F8F8F8;
                    }

                    th.sort-asc::after { content: " ▲"; }
                    th.sort-desc::after { content: " ▼"; }

                    tr:hover {
                        background-color: #B7E1EC;
                        cursor: pointer;
                    }

                    .button {
                        margin: 10px 0;
                        padding: 5px 15px;
                        font-size: 14px;
                        background-color: #2196F3;
                        color: white;
                        border: none;
                        border-radius: 5px;
                        cursor: pointer;
                    }

                    .button:hover {
                        background-color: #1976D2;
                    }

                    .button:disabled {
                        background-color: #cccccc;
                        cursor: not-allowed;
                    }

                    th.sortable:hover {
                        background-color: #005a9e;
                        cursor: pointer;
                    }

                    th .sort-icon {
                        font-size: 0.8em;
                        margin-left: 5px;
                        color: #ffffff;
                    }

                    .icon {
                        width: 16px;
                        height: 16px;
                        vertical-align: middle;
                    }

                    .mismatch {
                        color: #FF0000;
                        font-weight: bold;
                    }

                    .editable-cell {
                        padding: 0;
                    }

                    .editable-cell select,
                    .editable-cell input {
                        width: 100%;
                        padding: 8px;
                        border: 1px solid #ddd;
                        border-radius: 4px;
                    }

                    .validate-button {
                        background-color: #2196F3;
                        color: white;
                        border: none;
                        padding: 5px 10px;
                        border-radius: 4px;
                        cursor: pointer;
                    }

                    .validate-button:hover {
                        background-color: #1976D2;
                    }

                    .validate-button:disabled {
                        background-color: #cccccc;
                        cursor: not-allowed;
                    }

                    #missing-rows-file-table .editable-cell input {
                        height: 20px;
                        font-size: 12px;
                        padding: 2px;
                    }

                    #missing-rows-file-table .editable-cell select {
                        height: 24px;
                        font-size: 12px;
                        padding: 2px;
                    }

                    #missing-rows-file-table .validate-button {
                        height: 18px;
                        font-size: 12px;
                        padding: 2px 6px;
                        line-height: normal;
                    }
                </style>
            </head>
            <body>
                <h1>Data Comparison Tool</h1>

                <div class="section">
                    <h2>Upload CSV File</h2>
                    <input type="file" id="file-input" accept=".csv" />
                    <button
                        class="button"
                        id="copy-data-button" disabled>Copy Data to CLM</button>
                </div>

                <div class="section hidden" id="comparison-results">
                    <h2>Comparison Results</h2>

                    <h3 id="differences-title">Rows with Differences</h3>
                    <table id="differences-table">
                        <thead>
                            <tr>
                                <th data-key="Check Description" class="sortable">Check Description<div class="resize-handle"></div></th>
                                <th data-key="Task Status" class="sortable">Task Status<div class="resize-handle"></div></th>
                                <th data-key="Technician Comments" class="sortable">Technician Comments<div class="resize-handle"></div></th>
                                <th data-key="Manual Reference" class="sortable">Manual Reference<div class="resize-handle"></div></th>
                                <th data-key="Check ID" class="sortable">Check ID<div class="resize-handle"></div></th>
                                <th>Source</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>

                    <h3 id="missing-cloud-title">Rows Not Found in Cloud</h3>
                    <table id="missing-rows-cloud-table">
                        <thead>
                            <tr>
                                <th data-key="Check Description" class="sortable">Check Description<div class="resize-handle"></div></th>
                                <th data-key="Task Status" class="sortable">Task Status<div class="resize-handle"></div></th>
                                <th data-key="Technician Comments" class="sortable">Technician Comments<div class="resize-handle"></div></th>
                                <th data-key="Manual Reference" class="sortable">Manual Reference<div class="resize-handle"></div></th>
                                <th data-key="Check ID" class="sortable">Check ID<div class="resize-handle"></div></th>
                                <th>Source</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>

                    <h3 id="missing-file-title">Rows Not Found in File</h3>
                    <table id="missing-rows-file-table">
                        <thead>
                            <tr>
                                <th data-key="Check Description" class="sortable">Check Description<div class="resize-handle"></div></th>
                                <th>Task Status</th>
                                <th>Technician Comments</th>
                                <th data-key="Manual Reference" class="sortable">Manual Reference<div class="resize-handle"></div></th>
                                <th data-key="Check ID" class="sortable">Check ID<div class="resize-handle"></div></th>
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>

                    <h3 id="matching-title">Matching Rows</h3>
                    <table id="matching-table">
                        <thead>
                            <tr>
                                <th data-key="Check Description" class="sortable">Check Description<div class="resize-handle"></div></th>
                                <th data-key="Task Status" class="sortable">Task Status<div class="resize-handle"></div></th>
                                <th data-key="Technician Comments" class="sortable">Technician Comments<div class="resize-handle"></div></th>
                                <th data-key="Manual Reference" class="sortable">Manual Reference<div class="resize-handle"></div></th>
                                <th data-key="Check ID" class="sortable">Check ID<div class="resize-handle"></div></th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>
                </div>

                <div class="section">
                    <h2>Original Table Data</h2>
                    <button class="button" id="save-button">Save as CSV</button>
                    <table id="data-table">
                        <thead>
                            <tr>
                                <th data-key="Check Description" class="sortable">Check Description<div class="resize-handle"></div></th>
                                <th data-key="Task Status" class="sortable">Task Status<div class="resize-handle"></div></th>
                                <th data-key="Technician Comments" class="sortable">Technician Comments<div class="resize-handle"></div></th>
                                <th data-key="Manual Reference" class="sortable">Manual Reference<div class="resize-handle"></div></th>
                                <th data-key="Check ID" class="sortable">Check ID<div class="resize-handle"></div></th>
                            </tr>
                        </thead>
                        <tbody>
        `;

        // Append the extracted cloudData rows directly here
        cloudData.forEach(row => {
            htmlContent += `
                <tr>
                    <td>${escapeHTML(row["Check Description"])}</td>
                    <td>${escapeHTML(row["Task Status"])}</td>
                    <td>${escapeHTML(row["Technician Comments"])}</td>
                    <td>${escapeHTML(row["Manual Reference"])}</td>
                    <td>${escapeHTML(row["Check ID"])}</td>
                </tr>
            `;
        });

        htmlContent += `
                        </tbody>
                    </table>
                </div>

                <script>
                    updateOriginalDataCount(${cloudData.length});
                    const taskStatusOptions = ${taskStatusJSON};

                    /**
                     * Updates the count of original table data rows in the header (if present).
                     * @param {number} count - New count value.
                     */
                    function updateOriginalDataCount(count) {
                        const titleElement = document.getElementById('original-data-title');
                        if (titleElement) {
                            const textContent = titleElement.textContent;
                            const dashIndex = textContent.indexOf(': ');
                            const baseTitle = dashIndex !== -1 ? textContent.substring(0, dashIndex) : textContent;
                            titleElement.textContent = \`\${baseTitle}: \${count}\`;
                        } else {
                            console.warn(\`Element "\${titleElement}" not found.\`);
                        }
                    }

                    // Remaining scripts
                    window.addEventListener('beforeunload', function() {
                        if (window.opener) {
                            console.log (">>>close<<<");
                            window.opener.postMessage({ type: 'windowClosed' }, '*');
                        }
                    });
                </script>
            </body>
            </html>
        `;

        return htmlContent;
    }

    /**
     * Re-displays the original table data inside the new window if needed.
     * @param {Array<Object>} data - Array of row objects to display.
     */
    function displayOriginalTableData(data) {
        const tbody = doc.querySelector('#data-table tbody');

        if (data && data.length > 0) {
            const count = data.length;
            updateOriginalDataCount(count);
        }
    }

    /**
     * Saves the #data-table content as a CSV file.
     */
    function saveTableAsCSV() {
        const table = doc.getElementById('data-table');
        if (!table) {
            console.error("Table 'data-table' not found.");
            return;
        }
        const rows = table.querySelectorAll('tr');
        const csvContent = Array.from(rows)
            .map(row =>
                Array.from(row.querySelectorAll('th, td'))
                    .map(cell => '"' + cell.textContent.trim().replace(/"/g, '""') + '"')
                    .join(',')
            )
            .join('\n');

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = doc.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'table_data.csv';
        link.style.display = 'none';
        doc.body.appendChild(link);
        link.click();
        doc.body.removeChild(link);
    }

    /**
     * Enables column resizing for a given table.
     * @param {HTMLTableElement} table - The table element to enable resizing on.
     */
    function enableColumnResizing(table) {
        if (!table) return;

        const headers = table.querySelectorAll('th');
        let isResizing = false;
        let currentHeader;
        let nextHeader;
        let startX;
        let startWidth;
        let startNextWidth;

        headers.forEach((header, index) => {
            const handle = header.querySelector('.resize-handle');
            if (!handle) return;

            // Set initial width from CSS
            const width = parseInt(window.getComputedStyle(header).width);
            header.style.width = width + 'px';

            handle.addEventListener('mousedown', function (e) {
                isResizing = true;
                currentHeader = header;
                nextHeader = headers[index + 1];

                // Get current widths
                startWidth = parseInt(currentHeader.style.width);
                if (nextHeader) {
                    startNextWidth = parseInt(nextHeader.style.width);
                }

                startX = e.clientX;
                e.preventDefault();
                doc.body.style.cursor = 'col-resize';
            });
        });

        doc.addEventListener('mousemove', function (e) {
            if (!isResizing || !currentHeader || !nextHeader) return;

            const dx = e.clientX - startX;
            const minWidth = 30;

            let newWidth = startWidth + dx;
            let newNextWidth = startNextWidth - dx;

            if (newWidth < minWidth) {
                newWidth = minWidth;
                newNextWidth = startWidth + startNextWidth - minWidth;
            }

            if (newNextWidth < minWidth) {
                newNextWidth = minWidth;
                newWidth = startWidth + startNextWidth - minWidth;
            }

            currentHeader.style.width = newWidth + 'px';
            nextHeader.style.width = newNextWidth + 'px';
        });

        doc.addEventListener('mouseup', function () {
            if (isResizing) {
                isResizing = false;
                doc.body.style.cursor = 'default';
            }
        });
    }

    /**
     * Sorts a table by the specified column key in ascending or descending order.
     * @param {HTMLTableElement} table - The table to sort.
     * @param {string} key - The data-key attribute identifying the column.
     * @param {number} direction - 1 for ascending, -1 for descending.
     */
    function sortTableByKey(table, key, direction) {
        const tbody = table.querySelector('tbody');

        if (table.id === 'differences-table') {
            // Special logic for sorting paired rows
            const rows = Array.from(tbody.querySelectorAll('tr'));
            const pairedRows = [];

            // Group rows into pairs (File and Cloud) with a separator afterwards
            for (let i = 0; i < rows.length; i += 3) {
                const fileRow = rows[i];
                const cloudRow = rows[i + 1];
                if (fileRow && cloudRow) {
                    pairedRows.push([fileRow, cloudRow]);
                }
            }

            const headers = table.querySelectorAll('th');
            const keyIndex = Array.from(headers).findIndex(th => th.dataset.key === key);
            if (keyIndex === -1) return;

            // Sort pairs based on the value in the File row
            pairedRows.sort((pairA, pairB) => {
                const aText = pairA[0].children[keyIndex].textContent.trim();
                const bText = pairB[0].children[keyIndex].textContent.trim();

                const aNum = parseFloat(aText);
                const bNum = parseFloat(bText);

                if (!isNaN(aNum) && !isNaN(bNum)) {
                    return (aNum - bNum) * direction;
                }
                return aText.localeCompare(bText) * direction;
            });

            tbody.innerHTML = '';
            pairedRows.forEach(pair => {
                pair.forEach(row => tbody.appendChild(row));
                const separator = document.createElement('tr');
                separator.innerHTML = '<td colspan="6" style="border-bottom: 2px solid #ccc"></td>';
                tbody.appendChild(separator);
            });
        } else {
            // Regular sorting for other tables
            const rows = Array.from(tbody.querySelectorAll('tr'));
    
            // Determine the column index for sorting
            const headers = table.querySelectorAll('th');
            const keyIndex = Array.from(headers).findIndex(th => th.dataset.key === key);
            if (keyIndex === -1) return; // If column not found

            // Sort rows
            rows.sort((a, b) => {
                const aText = a.children[keyIndex].textContent.trim();
                const bText = b.children[keyIndex].textContent.trim();

                // Compare as numbers or strings
                const aNum = parseFloat(aText);
                const bNum = parseFloat(bText);

                if (!isNaN(aNum) && !isNaN(bNum)) {
                    return (aNum - bNum) * direction;
                }
                return aText.localeCompare(bText) * direction;
            });

            const fragment = document.createDocumentFragment();
            rows.forEach(row => fragment.appendChild(row));
            tbody.innerHTML = '';
            tbody.appendChild(fragment);
        }
    }

    /**
     * Initializes the sorting functionality for all sortable tables.
     */
    function initializeTableSorting() {
        const tables = doc.querySelectorAll('table');

        tables.forEach(table => {
            const headers = table.querySelectorAll('th.sortable');
            let sortDirection = {};

            headers.forEach(header => {
                const key = header.dataset.key;
                if (!key) return;

                sortDirection[key] = 1;

                header.addEventListener('click', function (e) {
                    if (e.target.classList.contains('resize-handle')) return;

                    const direction = sortDirection[key]; // Current sort direction
                    sortTableByKey(table, key, direction); // Perform the sort
                    sortDirection[key] *= -1; // Toggle the sort direction for next click

                    headers.forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
                    this.classList.add(direction === 1 ? 'sort-asc' : 'sort-desc');
                });
            });
        });
    }

    /**
     * Initializes event listeners in the comparison window.
     * @param {Window} parentWindow - Reference to the parent window.
     */
    async function initializeEventListeners(parentWindow) {
        if (!doc) {
            console.error('Document reference is null!');
            return;
        }
        const fileInput = doc.getElementById('file-input');
        const saveButton = doc.getElementById('save-button');
        const copyDataButton = doc.getElementById('copy-data-button');

        // File input handler
        if (fileInput) {
            fileInput.addEventListener('change', handleFileUpload);
        }
        // Save button handler
        if (saveButton) {
            saveButton.addEventListener('click', saveTableAsCSV);
        }
        // Copy data button handler
        if (copyDataButton) {
            copyDataButton.addEventListener('click', function () {
                if (StateManager.currentState !== StateManager.States.WINDOW_OPEN) {
                    return; // Copying is only possible in WINDOW_OPEN state
                }

                StateManager.setState(StateManager.States.COPYING);

                const matchingTable = doc.getElementById('matching-table');
                if (!matchingTable) {
                    console.error('Matching table not found');
                    return;
                }

                const rows = matchingTable.querySelectorAll('tbody tr');
                const matchingData = Array.from(rows).map(row => {
                    const cells = row.querySelectorAll('td');
                    return {
                        "Check Description": cells[0].textContent.trim(),
                        "Task Status": cells[1].textContent.trim(),
                        "Technician Comments": cells[2].textContent.trim(),
                        "Manual Reference": cells[3].textContent.trim(),
                        "Check ID": cells[4].textContent.trim()
                    };
                });

                // Send data back to the parent window
                try {
                    parentWindow.postMessage(
                        {
                            type: 'matchingData',
                            data: matchingData
                        },
                        '*'
                    );
                } catch (error) {
                    StateManager.handleError(error);
                    StateManager.setState(StateManager.States.WINDOW_OPEN);
                }
            });
        }

        // Enable column resizing on relevant tables
        [
            'data-table',
            'differences-table',
            'missing-rows-cloud-table',
            'missing-rows-file-table',
            'matching-table'
        ].forEach(tableId => {
            const table = doc.getElementById(tableId);
            if (table) {
                enableColumnResizing(table);
            }
        });

        // Initialize sort functionality
        initializeTableSorting();
    }

    // Listen for messages from the child window
    if (!window._messageListenerAdded) {
        window.addEventListener('message', function (event) {
            if (!event.data) return;
            switch (event.data.type) {
                case 'matchingData':
                    if (StateManager.currentState !== StateManager.States.COPYING) {
                        return;
                    }
                    TARGET_ROWS = event.data.data.map(row => ({
                        checkDescription: row["Check Description"],
                        manualReference: row["Manual Reference"],
                        checkID: row["Check ID"],
                        taskStatus: row["Task Status"],
                        technicianComments: row["Technician Comments"]
                    }));

                    processRowsWithData(TARGET_ROWS);
                    break;

                case 'windowClosed':
                    console.log(">>>close");
                    if (
                        (StateManager.currentState !== StateManager.States.ERROR) &&
                        (StateManager.currentState !== StateManager.States.IDLE)
                    ) {
                        StateManager.setState(StateManager.States.IDLE);
                    }
                    break;
            }
        });
        window._messageListenerAdded = true;
    }

    // Initialize the button after the page loads
    window.addEventListener('load', () => {
        // Small delay to ensure all elements are loaded
        setTimeout(() => {
            addButton();
        }, 500); // Reduced delay to 0.5 seconds
    });


    /**
     * Processes the rows that come from the child's matchingData message.
     * @param {Array<Object>} targetRows - The array of rows to process and fill back in the parent window.
     */
    async function processRowsWithData(targetRows) {
        if (!StateManager.setState(StateManager.States.FILLING)) {
            return;
        }

        try {
            const table = document.getElementById("gridview-1080");
            if (!table) {
                throw new Error("Table with id 'gridview-1080' not found.");
            }

            const allRows = table.querySelectorAll("tr.svmx-grid-row, tr.svmx-grid-row-alt");
            if (!allRows.length) {
                console.warn("No rows found in the table.");
                return;
            }

            for (let target of targetRows) {
                let foundRow = null;

                for (let row of allRows) {
                    const cd = row.querySelector(`.${CLASS_CD}`)?.textContent.trim();
                    const mr = row.querySelector(`.${CLASS_MR}`)?.textContent.trim();
                    const cid = row.querySelector(`.${CLASS_CID}`)?.textContent.trim();

                    if (cd === target.checkDescription &&
                        mr === target.manualReference &&
                        cid === target.checkID) {
                        foundRow = row;
                        break;
                    }
                }

                if (!foundRow) {
                    console.warn(`Row not found: ${JSON.stringify(target)}`);
                    continue;
                }

                const cellTS = foundRow.querySelector(`.${CLASS_TS}`);
                const cellTC = foundRow.querySelector(`.${CLASS_TC}`);

                if (cellTS) {
                    await setTaskStatus(cellTS, target.taskStatus);
                }

                if (cellTC) {
                    await setTechnicianComments(cellTC, target.technicianComments);
                }
            }
            console.log("Row processing completed.");
        } catch (error) {
            StateManager.handleError(error);
        } finally {
            if (!StateManager.hasError) {
                if (comparisonWindow && !comparisonWindow.closed) {
                    comparisonWindow.close();
                }
                StateManager.setState(StateManager.States.IDLE);
            }
        }
    }

    /**
     * Returns sample data for debugging purposes.
     * @returns {Array<Object>} - Array of predefined row objects.
     */
    function getFileData() {
        const headers = ["Check Description", "Task Status", "Technician Comments", "Manual Reference", "Check ID"];
        const data = [
            ["Testing the geometry of the coordinate frame", "Pass", "", "1531151_10", "4.10.3.4"],
            ["Examining the LGP printers", "Pass", "", "1531151_10", "4.12.1.2"],
            ["Examining the collimator cap emergency stop and condition", "Pass", "", "1531151_10", "4.8.3.1"],
            ["Examining the office cabinet", "Pass", "", "1531151_10", "04.05.2003"],
            ["Examining the angled fixation posts", "Pass", "", "1531151_10", "4.10.4.3"],
            ["Examining the focus precision with the QA tool for G-frame adapter", "Pass", "", "1531151_10", "4.7.11.2"],
            ["Examining the CT adapters", "Not Done – Not Applicable", "N/A", "1531151_10", "4.10.5.4"],
            ["Examining log files for CS software version 10.0 or higher", "Pass", "", "1531151_10", "04.02.2002"],
            ["Examining the mattress and patient couch", "Pass", "", "1531151_10", "04.03.2002"],
            ["Examining the X drive mechanics", "Pass", "", "1531151_10", "04.03.2005"],
            ["Finalizing the results", "Pass", "", "1531151_10", "4.13.7"],
            ["Examining the Z and X emergency release", "Pass", "", "1531151_10", "04.03.2007"],
            ["Examining the G-frame adapter", "Pass", "", "1531151_10", "04.03.2008"],
            ["Examining the sector locking solenoids", "Pass", "", "1531151_10", "04.04.2004"],
            ["Examining the door mechanics", "Pass", "", "1531151_10", "04.04.2005"],
            ["Examining the clearance check tool and the QA tool", "Pass", "", "1531151_10", "04.07.2005"],
            ["Validating the mechanical stop", "Pass", "", "1531151_10", "04.07.2009"],
            ["Handling issues from customer and log files", "Pass", "", "1531151_10", "4.13.1"],
            ["Examining the clearance check tool accuracy with G-frame adapter and QA tool", "Pass", "", "1531151_10", "04.07.2007"],
            ["Examining radioactive contamination", "Pass", "", "1531151_10", "4.8.3.2"],
            ["Examining the safety system interlocks", "Pass", "", "1531151_10", "04.08.2004"],
            ["Examining the radiation leakage rate", "Pass", "", "1531151_10", "04.09.2002"],
            ["Measuring the absorbed dose rate", "Not Done – Not Applicable", "N/A", "1531151_10", "04.09.2003"],
            ["Examining the general conditions of the coordinate frame", "Pass", "", "1531151_10", "4.10.3.2"],
            ["Examining the MR adapters", "Pass", "", "1531151_10", "4.10.5.2"],
            ["Examining the X-ray indicators (text changed)", "Fail", "Fail", "1531151_10", "4.10.5.5"],
            ["Examining the X-ray adapters (text changed)", "Not Done – Not Applicable", "Fail", "1531151_10", "4.10.5.6"],
            ["Examining the radiation phantoms", "Pass", "", "changed", "4.10.7.1"],
            ["Examining the Film Holder Tool", "Pass", "", "changed", "4.10.7.2"],
            ["Recording the serial number and age of each Vantage Head Frame", "Not Done – Not Applicable", "N/A", "1531151_10", "changed"],
            ["Examining the Leksell® Vantage™ MRI Fiducial Box", "Pass", "", "1531151_10", "4.11.5.1"],
            ["Examining LGP application software", "Pass", "", "changed", "changed"],
            ["Checking the Emergency Procedures notices", "Pass", "", "1531151_10", "4.13.5"],
            ["Restoring hardware and software", "Pass", "", "1531151_10", "4.13.4"],
            ["Examining the docking device with the G-frame adaper", "Pass", "", "1531151_10", "04.03.2010"],
            ["Examining the Leksell® Vantage™ Frame Holder, the Leksell® Vantage™ Frame Holder extension, the Leksell® Vantage™ CT interface, and the Leksell® Vantage™ CT Table fixation", "Not Done – Not Applicable", "N/A", "1531151_10", "4.11.5.4"],
            ["Examining Z drive mechanics", "Pass", "", "1531151_10", "04.03.2004"],
            ["Exporting statistics from Leksell Gamma Knife treatments", "Pass", "", "1531151_10", "04.12.2003"],
            ["Measuring the accuracy in dose delivery center", "Pass", "", "1531151_10", "04.09.2004"],
            ["Examining the Leksell® Vantage™ CT Fiducial Box", "Pass", "", "1531151_10", "4.11.5.3"],
            ["Examining the quick fixation screws", "Pass", "", "1531151_10", "4.10.4.1"],
            ["Examining the sensors", "Pass", "", "1531151_10", "04.04.2008"],
            ["Cleaning the system and the site", "Pass", "", "1531151_10", "4.13.6"],
            ["Examining the Vantage frame adapter", "Not Done – Not Applicable", "N/A", "1531151_10", "04.03.2009"],
            ["Examining the patient surveillance system", "Pass", "", "1531151_10", "04.07.2003"],
            ["Storing configuration files at the site for CS SW versions 10.0 or higher", "Pass", "", "1531151_10", "4.13.3"],
            ["Examining the CT indicators", "Not Done – Not Applicable", "N/A", "1531151_10", "4.10.5.3"],
            ["Examining the general conditions of the Vantage Head frame", "Not Done – Not Applicable", "N/A", "1531151_10", "04.11.2004"],
            ["Examining the emergency alarm", "Pass", "", "1531151_10", "04.07.2004"],
            ["Connecting a laptop in the treatment room", "Pass", "", "1531151_10", "04.01.2001"],
            ["Examining the operator console", "Pass", "", "1531151_10", "04.05.2001"],
            ["Examining the office and medical UPS functionality", "Pass", "", "1531151_10", "04.05.2005"],
            ["Recording problems reported by the customer", "Pass", "", "1531151_10", "04.02.2001"],
            ["Examining the initialization sequence", "Pass", "", "1531151_10", "04.02.2003"],
            ["Removing the covers", "Pass", "", "1531151_10", "04.02.2004"],
            ["Examining the couch movement", "Pass", "", "1531151_10", "04.03.2001"],
            ["Visually examining the electronics", "Pass", "", "1531151_10", "04.03.2003"],
            ["Examining the Y drive mechanics", "Pass", "", "1531151_10", "04.03.2006"],
            ["Examining the operation of the Pause and Emergency Stop buttons", "Pass", "", "1531151_10", "04.08.2001"],
            ["Examining the docking device with the Vantage frame adaper", "Not Done – Not Applicable", "N/A", "1531151_10", "04.03.2011"],
            ["Examining the temperature and humidity", "Pass", "", "1531151_10", "04.04.2001"],
            ["Examining the sector motor assembly", "Pass", "", "1531151_10", "04.04.2002"],
            ["Recording the serial number and the age of each Coordinate frame G", "Pass", "", "1531151_10", "4.10.3.1"],
            ["Examining the sector movement", "Pass", "", "1531151_10", "04.04.2003"],
            ["Examining the door travel distances", "Pass", "", "1531151_10", "04.04.2006"],
            ["Examining the monitors", "Pass", "", "1531151_10", "04.05.2002"],
            ["Sending log files to Elekta for CS SW version 10.0 or higher", "Pass", "", "1531151_10", "4.13.2"],
            ["Examining the medical cabinet", "Pass", "", "1531151_10", "04.05.2004"],
            ["Examining the MCU", "Pass", "", "1531151_10", "04.05.2006"],
            ["Examining the covers and the foot squeeze protection", "Pass", "", "1531151_10", "04.07.2001"],
            ["Examining the interlocks, buttons, and indicators", "Pass", "", "1531151_10", "04.07.2002"],
            ["Examining the clearance check tool sensor", "Pass", "", "1531151_10", "04.07.2006"],
            ["Examining LGP system software", "Pass", "", "1531151_10", "4.12.2.1"],
            ["Examining the clearance check tool accuracy with Vantage frame adapter and QA tool Vantage", "Not Done – Not Applicable", "N/A", "1531151_10", "04.07.2008"],
            ["Examining the Emergency Exit sequence", "Pass", "", "1531151_10", "04.08.2002"],
            ["Examining the sector positions", "Pass", "", "1531151_10", "04.07.2010"],
            ["Examining the PPS precision", "Pass", "", "1531151_10", "4.7.11.1"],
            ["Examining the focus precision with the QA tool Vantage", "Not Done – Not Applicable", "N/A", "1531151_10", "4.7.11.3"],
            ["Examining the curved fixation posts", "Pass", "", "1531151_10", "4.10.4.4"],
            ["Examining the MR indicators", "Pass", "", "1531151_10", "4.10.5.1"],
            ["Examining the frame cap", "Pass", "", "1531151_10", "04.10.2006"],
            ["Examining the MRI adapters", "Pass", "", "1531151_10", "4.11.5.2"],
            ["Examining the LGP workstations", "Pass", "", "1531151_10", "4.12.1.1"],
            ["Extra Line in file 1", "Pass", "", "1531151_10", "4.12.1.1"],
            ["Extra Line in file 2", "Pass", "", "153_2", "5.2"],
            ["Extra Line in file 3", "Pass", "comment", "153_3", "5.2"],
            ["Extra Line in file 4", "Pass", "", "153_4", "5.2"],
            ["Extra Line in file 5", "Pass", "", "153_5", "5.2"]
        ];

        return data.map(row => Object.fromEntries(row.map((value, index) => [headers[index], value])));
    }

})();
