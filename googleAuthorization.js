const { authenticate: GoogleAuthorization } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const fs = require('fs').promises;
const path = require('path');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/calendar'];

const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');
const BACKUP_TOKEN_PATH = path.join(process.cwd(), 'token_backup.json');

async function loadSavedCredentialsIfExist() {
    try {
        const content = await fs.readFile(TOKEN_PATH);
        const credentials = JSON.parse(content);
        return google.auth.fromJSON(credentials);
    } catch (err) {
        return null;
    }
}

async function backupToken() {
    try {
        await fs.copyFile(TOKEN_PATH, BACKUP_TOKEN_PATH);
    } catch (error) {
        console.error('Error backing up token:', error);
    }
}

async function deleteToken() {
    try {
        await fs.unlink(TOKEN_PATH);
    } catch (error) {
        console.error('Error deleting token:', error);
    }
}

async function deleteBackupToken() {
    try {
        await fs.unlink(BACKUP_TOKEN_PATH);
    } catch (error) {
        console.error('Error deleting backup token:', error);
    }
}

async function saveCredentials(client) {
    const content = await fs.readFile(CREDENTIALS_PATH);
    const keys = JSON.parse(content);
    const key = keys.installed || keys.web;
    const payload = JSON.stringify({
        type: 'authorized_user',
        client_id: key.client_id,
        client_secret: key.client_secret,
        refresh_token: client.credentials.refresh_token,
    });
    await fs.writeFile(TOKEN_PATH, payload);
}

async function authorize() {
    let client = await loadSavedCredentialsIfExist();
    if (client) {
        try {
            const calendar = google.calendar({version: 'v3', auth: client});
// Replace this with any other API call that requires authentication
            await calendar.calendarList.list();
            return client;
        } catch (error) {
            if (error.code === 401 && error.response && error.response.data && error.response.data.error === 'invalid_grant') {
                await backupToken();
                await deleteToken();
            } else {
                console.error('Error while checking credentials:', error);
            }
        }
    }

    client = await GoogleAuthorization({
        scopes: SCOPES,
        keyfilePath: CREDENTIALS_PATH,
    });

    if (client.credentials) {
        await saveCredentials(client);
        await deleteBackupToken();
    }

    return client;
}

module.exports = authorize;
