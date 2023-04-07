const fs = require('fs').promises;
const path = require('path');

const INSTITUTIONS_PATH = path.join(process.cwd(), 'institutions.json');

async function readInstitutions() {
    const content = await fs.readFile(INSTITUTIONS_PATH, 'utf-8');
    return JSON.parse(content);
}

async function writeInstitutions(institutions) {
    const content = JSON.stringify(institutions, null, 2);
    await fs.writeFile(INSTITUTIONS_PATH, content, 'utf-8');
}

async function getInstitution(serial_id) {
    const institutions = await readInstitutions();
    return institutions.find(inst => inst.serial_id === serial_id);
}

async function getDefaultEmail(serial_id) {
    const institution = await getInstitution(serial_id);
    return institution.emails.find(email => email.is_default);
}

module.exports = {
    readInstitutions,
    writeInstitutions,
    getInstitution,
    getDefaultEmail
};
