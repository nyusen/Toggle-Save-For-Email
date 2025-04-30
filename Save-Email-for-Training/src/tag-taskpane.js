// Global variables for authentication
let accessToken = null;
let idToken = null;

// Store tags globally so we can filter them
let availableTags = [];

// Store the authenticated HTML content
const authenticatedHTML = `
    <div class="tag-container">
        <div class="section">
            <h2 class="section-title">All Tags</h2>
            <div class="search-container">
                <input type="text" 
                       id="tagSearch" 
                       class="search-input" 
                       placeholder="Search tags..."
                       oninput="filterTags(this.value)">
            </div>
            <div id="tagList" class="tag-list">
                <!-- Tags will be populated here -->
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">Add a Tag</h2>
            <div class="add-tag-container">
                <input type="text" 
                       id="customTagInput" 
                       class="custom-tag-input" 
                       placeholder="Enter a custom tag"
                       oninput="toggleAddButton()">
                <button id="addTagButton" 
                        class="add-tag-button" 
                        onclick="addCustomTag()" 
                        disabled>Add Tag</button>
                <div id="tagError" class="tag-error">This tag already exists</div>
            </div>
        </div>

        <div class="section">
            <h2 class="section-title">Selected Tags</h2>
            <div id="selectedTags" class="selected-tags">
                <!-- Selected tags will appear here -->
            </div>
        </div>
    </div>
`;

Office.onReady(() => {
    // Check if we have tokens in session storage
    accessToken = sessionStorage.getItem('accessToken');
    idToken = sessionStorage.getItem('idToken');

    // Set saveForTraining to true when taskpane loads
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const props = result.value;
            props.set("saveForTraining", true);
            props.saveAsync(() => {
                initializeUI();
            });
        }
    });
});

function initializeUI() {
    if (idToken) {
        // If authenticated, show the tag interface and load tags
        document.body.innerHTML = authenticatedHTML;
        loadTags();
    } else {
        // If not signed in, show sign-in dialog
        document.body.innerHTML = '<div class="error-message">Please sign in to manage tags.</div>';
        showSignInDialog();
    }
}

function showSignInDialog() {
    Office.context.ui.displayDialogAsync('https://nyusen.github.io/Toggle-Save-For-Email/signin-dialog.html',
        { height: 40, width: 30, displayInIframe: true },
        function (signInResult) {
            if (signInResult.status === Office.AsyncResultStatus.Failed) {
                console.error(signInResult.error.message);
                showError('Failed to open sign-in dialog');
                return;
            }

            const signInDialog = signInResult.value;
            signInDialog.addEventHandler(Office.EventType.DialogMessageReceived, async function (signInArg) {
                if (signInArg.message === 'signin') {
                    // User wants to sign in
                    signInDialog.close();
                    // Add timeout before proceeding
                    setTimeout(() => {
                        handleSignIn().then(() => {
                            // User is signed in, restore the authenticated UI and load tags
                            setTimeout(() => {
                                initializeUI();
                            }, 1000);
                        }).catch((error) => {
                            showError('Failed to sign in');
                        });
                    }, 1000);
                } else if (signInArg.message === 'cancel') {
                    // User doesn't want to sign in
                    signInDialog.close();
                    setTimeout(() => {
                        showError('Authentication required to load tags');
                    }, 1000);
                }
            });
        }
    );
}

// Function to handle sign-in
function handleSignIn() {
    return new Promise((resolve, reject) => {
        // Generate random state and PKCE verifier
        const state = generateRandomString(32);
        const codeVerifier = generateRandomString(64);
        const nonce = generateRandomString(16);

        // Store state and code verifier in parent window's session storage
        window.sessionStorage.setItem('authState', state);
        window.sessionStorage.setItem('codeVerifier', codeVerifier);

        // Generate code challenge
        generateCodeChallenge(codeVerifier).then(codeChallenge => {
            // Build the URL to our local sign-in start page
            const startUrl = new URL('https://nyusen.github.io/Toggle-Save-For-Email/sign-in-start.html');
            startUrl.searchParams.append('state', state);
            startUrl.searchParams.append('nonce', nonce);
            startUrl.searchParams.append('code_challenge', codeChallenge);
            startUrl.searchParams.append('code_challenge_method', 'S256');
            startUrl.searchParams.append('code_verifier', codeVerifier);

            // Add delay before opening the auth dialog
            setTimeout(() => {
                // Open our local page in a dialog, which will then redirect to Microsoft
                Office.context.ui.displayDialogAsync(startUrl.toString(),
                    { height: 60, width: 30 },
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Failed) {
                            console.error(result.error.message);
                            reject(new Error(result.error.message));
                            return;
                        }

                        const dialog = result.value;
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                            try {
                                const message = JSON.parse(arg.message);
                                if (message.type === 'token') {
                                    // Store both tokens
                                    accessToken = message.token;
                                    idToken = message.idToken;
                                    sessionStorage.setItem('accessToken', message.token);
                                    sessionStorage.setItem('idToken', message.idToken);
                                    sessionStorage.setItem('tokenType', message.tokenType);
                                    sessionStorage.setItem('expiresIn', message.expiresIn);
                                    dialog.close();
                                    resolve(message.token);
                                } else if (message.type === 'error') {
                                    dialog.close();
                                    reject(new Error(message.error));
                                }
                            } catch (e) {
                                dialog.close();
                                reject(e);
                            }
                        });

                        dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
                            dialog.close();
                            reject(new Error('Dialog closed'));
                        });
                    }
                );
            }, 1000); // Wait 1 second to ensure previous dialog is fully closed
        }).catch(error => {
            reject(error);
        });
    });
}

// Function to show error message in the taskpane
function showError(message) {
    const container = document.querySelector('.tag-container');
    container.innerHTML = `
        <div style="color: red; padding: 20px; text-align: center;">
            <p>${message}</p>
            <button onclick="showSignInDialog()">Sign In</button>
        </div>
    `;
}

// Function to fetch tags from the server
async function loadTags() {
    showTagListLoading();

    try {
        // Make authenticated request to get tags
        const response = await makeAuthenticatedRequest('https://ml-inf-svc-prd.eventellect.com/corpus-collector/api/metadata/tags');
        availableTags = await response.json();
        displayFilteredTags(availableTags);
    } catch (error) {
        console.error('Error loading tags:', error);
        const tagList = document.getElementById('tagList');
        tagList.innerHTML = '<div class="error-message">Failed to load tags. Please try again later.</div>';
    }
}

// Function to display filtered tags
function displayFilteredTags(tags) {
    const tagList = document.getElementById('tagList');
    if (tags.length === 0) {
        tagList.innerHTML = '<div class="no-results">No matching tags found</div>';
        return;
    }

    // Sort tags alphabetically by description, ignoring case
    const sortedTags = [...tags].sort((a, b) =>
        a.description.toLowerCase().localeCompare(b.description.toLowerCase())
    );

    // Get currently selected tags from custom properties
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const props = result.value;
            let selectedTags = props.get('selectedTags') || [];
            if (typeof selectedTags === 'string') {
                selectedTags = JSON.parse(selectedTags);
            }

            // Create tag list with checked state preserved
            tagList.innerHTML = sortedTags.map(tag => {
                const isChecked = selectedTags.some(selectedTag => selectedTag.id === tag.id);
                return `
                    <div class="tag-item">
                        <input type="checkbox" id="tag-${tag.id}" value="${tag.id}" 
                            onchange="handleTagSelection(this)" ${isChecked ? 'checked' : ''}>
                        <label for="tag-${tag.id}">${tag.description}</label>
                    </div>
                `;
            }).join('');
        }
    });
}

// Function to filter tags based on search input
function filterTags(searchText) {
    if (!availableTags) return;

    const searchLower = searchText.toLowerCase().trim();
    const filteredTags = searchLower === ''
        ? availableTags
        : availableTags.filter(tag =>
            tag.description.toLowerCase().includes(searchLower));

    displayFilteredTags(filteredTags);

    // Restore checked state for selected tags
    const selectedTags = getSelectedTags();
    selectedTags.forEach(tag => {
        const checkbox = document.getElementById(`tag-${tag.id}`);
        if (checkbox) {
            checkbox.checked = true;
        }
    });
}

// Function to get currently selected tags
function getSelectedTags() {
    const selectedTagsContainer = document.getElementById('selectedTags');
    return Array.from(selectedTagsContainer.getElementsByClassName('selected-tag'))
        .map(tagElement => {
            const tagId = parseInt(tagElement.dataset.tagId);
            return availableTags.find(tag => tag.id === tagId);
        })
        .filter(tag => tag); // Remove any undefined entries
}

// Function to make authenticated requests
async function makeAuthenticatedRequest(url, options = {}) {
    if (!idToken) {
        throw new Error('No id token available');
    }

    const requestOptions = {
        ...options,
        headers: {
            ...options.headers,
            'Authorization': `Bearer ${idToken}`
        }
    };

    const response = await fetch(url, requestOptions);

    if (response.status === 401) {
        // Clear tokens if unauthorized
        sessionStorage.removeItem('accessToken');
        sessionStorage.removeItem('idToken');
        accessToken = null;
        idToken = null;
    }

    return response;
}

// Utility functions for authentication
function generateRandomString(length) {
    const array = new Uint8Array(length);
    crypto.getRandomValues(array);
    return Array.from(array, byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

async function generateCodeChallenge(verifier) {
    const encoder = new TextEncoder();
    const data = encoder.encode(verifier);
    const hash = await crypto.subtle.digest('SHA-256', data);
    return base64URLEncode(hash);
}

function base64URLEncode(buffer) {
    return btoa(String.fromCharCode(...new Uint8Array(buffer)))
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, '');
}

// Function to handle tag selection
function handleTagSelection(checkbox) {
    const tagId = parseInt(checkbox.value);
    const selectedTagsContainer = document.getElementById('selectedTags');

    if (checkbox.checked) {
        // Find the tag object from availableTags
        const selectedTag = availableTags.find(tag => tag.id === tagId);

        if (selectedTag) {
            const tagElement = document.createElement('span');
            tagElement.className = 'selected-tag';
            tagElement.dataset.tagId = selectedTag.id;
            tagElement.dataset.description = selectedTag.description; // Store description for getSelectedTags
            tagElement.innerHTML = `
                ${selectedTag.description}
                <span class="remove-tag" onclick="removeTag('${selectedTag.id}')">&times;</span>
            `;
            selectedTagsContainer.appendChild(tagElement);
        }
    } else {
        // Remove the tag
        const existingTag = selectedTagsContainer.querySelector(`[data-tag-id="${tagId}"]`);
        if (existingTag) {
            existingTag.remove();
        }
    }

    // Save selected tags to custom properties
    saveSelectedTags();
}

// Function to remove a tag
function removeTag(tagId) {
    // Uncheck the checkbox
    const checkbox = document.getElementById(`tag-${tagId}`);
    if (checkbox) {
        checkbox.checked = false;
    }

    // Remove the tag from selected tags
    const selectedTagsContainer = document.getElementById('selectedTags');
    const existingTag = selectedTagsContainer.querySelector(`[data-tag-id="${tagId}"]`);
    if (existingTag) {
        existingTag.remove();
    }

    // Save selected tags to custom properties
    saveSelectedTags();
}

// Function to save selected tags to custom properties
function saveSelectedTags() {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const props = result.value;
            const selectedTags = getSelectedTags();
            props.set("selectedTags", JSON.stringify(selectedTags));
            props.saveAsync();
        }
    });
}

// Function to load selected tags from custom properties
function loadSelectedTags() {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const props = result.value;
            const savedTags = props.get("selectedTags");
            if (savedTags) {
                try {
                    const selectedTagIds = JSON.parse(savedTags);
                    selectedTagIds.forEach(tag => {
                        const checkbox = document.getElementById(`tag-${tag.id}`);
                        if (checkbox) {
                            checkbox.checked = true;
                            // Trigger the selection handler to update the UI
                            handleTagSelection(checkbox);
                        }
                    });
                } catch (e) {
                    console.error('Error parsing saved tags:', e);
                }
            }
        }
    });
}

// Function to normalize tag text for comparison
function normalizeTagText(text) {
    return text.toLowerCase().replace(/\s+/g, '');
}

// Function to show tag error message
function showTagError(message = 'This tag already exists') {
    const errorElement = document.getElementById('tagError');
    errorElement.textContent = message;
    errorElement.classList.add('visible');
    setTimeout(() => {
        errorElement.classList.remove('visible');
    }, 3000);
}

// Function to toggle the Add Tag button state
function toggleAddButton() {
    const input = document.getElementById('customTagInput');
    const button = document.getElementById('addTagButton');
    const errorElement = document.getElementById('tagError');

    if (input.value.trim() !== '') {
        button.classList.add('active');
        button.disabled = false;
    } else {
        button.classList.remove('active');
        button.disabled = true;
    }
}

// Function to set button loading state
function setAddButtonLoading(isLoading) {
    const button = document.getElementById('addTagButton');
    const input = document.getElementById('customTagInput');

    if (isLoading) {
        button.classList.add('loading');
        button.disabled = true;
        input.disabled = true;
        button.textContent = 'Adding';
    } else {
        button.classList.remove('loading');
        button.textContent = 'Add Tag';
        input.disabled = false;
        toggleAddButton(); // This will set the appropriate enabled/disabled state
    }
}

// Function to add a custom tag
async function addCustomTag() {
    const input = document.getElementById('customTagInput');
    const tagName = input.value.trim();

    if (!tagName) {
        return;
    }

    // Check if tag already exists
    const normalizedNewTag = normalizeTagText(tagName);
    const tagExists = availableTags.some(tag =>
        normalizeTagText(tag.description) === normalizedNewTag
    );

    if (tagExists) {
        showTagError();
        return;
    }

    setAddButtonLoading(true);

    try {
        // Make API call to get tag ID
        const response = await makeAuthenticatedRequest('https://ml-inf-svc-prd.eventellect.com/corpus-collector/api/metadata/tags', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ description: tagName })
        });

        const responseData = await response.json();

        if (response.status === 400) {
            showTagError(responseData.detail);
            setAddButtonLoading(false);
            return;
        }

        if (!response.ok) {
            throw new Error('Failed to create tag');
        }

        const tagObject = {
            id: responseData.id,
            description: responseData.description
        };

        // Add the new tag to availableTags
        availableTags.push(tagObject);

        // Store the tag object in custom properties
        Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const props = result.value;
                let selectedTags = props.get('selectedTags') || [];
                if (typeof selectedTags === 'string') {
                    selectedTags = JSON.parse(selectedTags);
                }

                // Add the new tag to selected tags if not already present
                if (!selectedTags.some(tag => tag.id === tagObject.id)) {
                    selectedTags.push(tagObject);
                    props.set('selectedTags', JSON.stringify(selectedTags));
                    props.saveAsync(() => {
                        // After saving, update both displays with the new state
                        updateSelectedTagsDisplay(selectedTags);
                        displayFilteredTags(availableTags);
                        input.value = '';
                        setAddButtonLoading(false);
                    });
                }
            }
        });
    } catch (error) {
        console.error('Error creating tag:', error);
        setAddButtonLoading(false);
        showTagError('Failed to create tag. Please try again.');
    }
}

// Function to update the selected tags display
function updateSelectedTagsDisplay(selectedTags) {
    const selectedTagsContainer = document.getElementById('selectedTags');

    selectedTagsContainer.innerHTML = selectedTags.map(tag => `
        <span class="selected-tag" data-tag-id="${tag.id}">
            ${tag.description}
            <span class="remove-tag" onclick="removeTag('${tag.id}')">&times;</span>
        </span>
    `).join('');
}

// Function to show loading state in tag list
function showTagListLoading() {
    const tagList = document.getElementById('tagList');
    tagList.innerHTML = '<div class="loading-spinner">Loading tags</div>';
}
