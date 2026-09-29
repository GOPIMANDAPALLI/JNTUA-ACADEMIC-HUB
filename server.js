// ==========================================================
// JNTUA ACADEMIC HUB - server.js (Excel + Persistent Disk)
// ==========================================================

const express = require('express');
const path = require('path');
const session = require('express-session');
const FileStore = require('session-file-store')(session);
const ExcelJS = require('exceljs');
const fs = require('fs');
const bcrypt = require('bcrypt');

const app = express();
app.set('trust proxy', 1);

// ==========================================================
// DIRECTORIES / FILE PATHS
// ==========================================================

const PUBLIC_DIR = path.join(__dirname, 'public');
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const EXCEL_FILE = path.join(DATA_DIR, 'users.xlsx');
const SESSION_DIR = path.join(DATA_DIR, 'sessions');

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(SESSION_DIR, { recursive: true });

// ==========================================================
// MIDDLEWARE
// ==========================================================

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(session({
    store: new FileStore({
        path: SESSION_DIR,
        ttl: 60 * 60 * 24 * 30,
        retries: 1,
        logFn: () => {}
    }),
    secret: process.env.SESSION_SECRET || 'jntua-secret',
    resave: false,
    saveUninitialized: false,
    cookie: {
        httpOnly: true,
        secure: process.env.NODE_ENV === 'production',
        sameSite: 'lax',
        maxAge: 1000 * 60 * 60 * 24 * 30
    }
}));

app.use(express.static(PUBLIC_DIR));

// ==========================================================
// EXCEL HELPERS
// ==========================================================

const HEADERS = [
    'Name', 'Email', 'Regulation', 'RollNumber', 'Branch',
    'YearOfStudy', 'CollegeName', 'Phone', 'Password'
];

// Runs one file operation at a time so the Excel file never gets corrupted
let queue = Promise.resolve();
function runExclusive(task) {
    const result = queue.then(task, task);
    queue = result.catch(() => {});
    return result;
}

async function createExcelFile() {
    if (fs.existsSync(EXCEL_FILE)) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Users');
    worksheet.addRow(HEADERS);
    await workbook.xlsx.writeFile(EXCEL_FILE);
    console.log('Created Excel file:', EXCEL_FILE);
}

async function readWorkbook() {
    await createExcelFile();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    let worksheet = workbook.getWorksheet('Users') || workbook.getWorksheet(1);
    if (!worksheet) {
        worksheet = workbook.addWorksheet('Users');
        worksheet.addRow(HEADERS);
    }
    return { workbook, worksheet };
}

// Safe write: write to a temp file first, then replace the real file
async function saveWorkbook(workbook) {
    const tempFile = EXCEL_FILE + '.tmp';
    await workbook.xlsx.writeFile(tempFile);
    fs.renameSync(tempFile, EXCEL_FILE);
}

function addUserToExcel(user) {
    return runExclusive(async () => {
        const { workbook, worksheet } = await readWorkbook();

        const newEmail = String(user.email || '').trim().toLowerCase();

        let emailExists = false;
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;
            const existing = String(row.getCell(2).value || '').trim().toLowerCase();
            if (existing === newEmail) emailExists = true;
        });

        if (emailExists) throw new Error('Email already registered');

        const passwordHash = await bcrypt.hash(String(user.password), 10);

        worksheet.addRow([
            user.name || '',
            newEmail,
            user.regulation || '',
            user.rollNumber || '',
            user.branch || '',
            user.yearStudy || '',
            user.collegeName || '',
            user.phone || '',
            passwordHash
        ]);

        await saveWorkbook(workbook);
        console.log('User registered:', newEmail);
    });
}

async function findStoredPassword(email) {
    if (!fs.existsSync(EXCEL_FILE)) return { fileMissing: true };

    const { worksheet } = await readWorkbook();
    let stored = null;

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const rowEmail = String(row.getCell(2).value || '').trim().toLowerCase();
        if (rowEmail === email) stored = String(row.getCell(9).value || '');
    });

    return { stored };
}

// ==========================================================
// HOME
// ==========================================================

app.get('/', (req, res) => {
    const file = (req.session && req.session.user)
        ? 'main-portal.html'
        : 'login.html';

    const page = path.join(PUBLIC_DIR, file);

    if (!fs.existsSync(page)) {
        return res.status(500).send(`${file} not found inside public folder.`);
    }
    return res.sendFile(page);
});

// ==========================================================
// REGISTER
// ==========================================================

app.post('/register', async (req, res) => {
    try {
        const user = req.body || {};

        if (!user.name || !user.email || !user.password || !user.confirmPassword) {
            return res.status(400).json({ message: 'Fill all required fields' });
        }

        user.email = String(user.email).trim().toLowerCase();

        if (user.password !== user.confirmPassword) {
            return res.status(400).json({ message: 'Passwords do not match' });
        }

        await addUserToExcel(user);

        return res.json({ message: 'Registered successfully' });

    } catch (err) {
        console.error('REGISTER ERROR:', err);

        if (err.message === 'Email already registered') {
            return res.status(400).json({ message: 'Email already registered' });
        }
        return res.status(500).json({ message: 'Server error. Try later.' });
    }
});

// ==========================================================
// LOGIN
// ==========================================================

app.post('/login', async (req, res) => {
    try {
        const email = String(req.body.email || '').trim().toLowerCase();
        const password = String(req.body.password || '');

        if (!email || !password) {
            return res.status(400).json({ message: 'Enter email and password' });
        }

        const { fileMissing, stored } = await findStoredPassword(email);

        if (fileMissing) {
            return res.status(400).json({ message: 'No users registered yet' });
        }

        if (stored === null) {
            return res.status(400).json({ message: 'Email not registered' });
        }

        // Supports hashed passwords, and old plain-text ones already in the file
        const passwordOk = stored.startsWith('$2')
            ? await bcrypt.compare(password, stored)
            : stored === password;

        if (!passwordOk) {
            return res.status(400).json({ message: 'Incorrect password' });
        }

        req.session.user = email;

        req.session.save((sessionError) => {
            if (sessionError) {
                console.error('SESSION ERROR:', sessionError);
                return res.status(500).json({
                    message: 'Session error. Please try again.'
                });
            }
            return res.json({ message: 'Login successful', redirect: '/' });
        });

    } catch (err) {
        console.error('LOGIN ERROR:', err);
        return res.status(500).json({ message: 'Server error. Try later.' });
    }
});

// ==========================================================
// PROTECTED MAIN PORTAL
// ==========================================================

app.get('/main-portal.html', (req, res) => {
    if (!req.session || !req.session.user) {
        return res.redirect('/');
    }
    const page = path.join(PUBLIC_DIR, 'main-portal.html');
    if (!fs.existsSync(page)) {
        return res.status(500).send('main-portal.html not found inside public folder.');
    }
    return res.sendFile(page);
});

app.get('/api/current-user', (req, res) => {
    if (!req.session || !req.session.user) {
        return res.status(401).json({ loggedIn: false });
    }
    return res.json({ loggedIn: true, email: req.session.user });
});

app.get('/logout', (req, res) => {
    req.session.destroy((err) => {
        if (err) {
            console.error('LOGOUT ERROR:', err);
            return res.status(500).send('Logout failed');
        }
        res.clearCookie('connect.sid');
        return res.redirect('/');
    });
});

// ==========================================================
// OWNER / ADMIN PANEL
// ==========================================================

app.get('/owner-login', (req, res) => {
    if (req.session && req.session.isAdmin === true) {
        return res.redirect('/admin/panel');
    }
    res.set('X-Robots-Tag', 'noindex, nofollow');
    return res.sendFile(path.join(__dirname, 'private', 'admin-login.html'));
});

app.post('/admin/login', async (req, res) => {
    try {
        const email = String(req.body.email || '').trim().toLowerCase();
        const password = String(req.body.password || '');

        const adminEmail = String(process.env.ADMIN_EMAIL || '').trim().toLowerCase();
        const adminPasswordHash = String(process.env.ADMIN_PASSWORD_HASH || '');

        if (!email || !password) {
            return res.status(400).json({
                message: 'Enter owner email and password'
            });
        }

        if (email !== adminEmail) {
            return res.status(401).json({ message: 'Invalid owner credentials' });
        }

        if (!adminPasswordHash) {
            console.error('ADMIN_PASSWORD_HASH is missing in Render Environment Variables');
            return res.status(500).json({
                message: 'Owner password is not configured on server'
            });
        }

        const passwordCorrect = await bcrypt.compare(password, adminPasswordHash);

        if (!passwordCorrect) {
            return res.status(401).json({ message: 'Invalid owner credentials' });
        }

        req.session.isAdmin = true;
        req.session.adminEmail = email;

        req.session.save((err) => {
            if (err) {
                console.error('ADMIN SESSION ERROR:', err);
                return res.status(500).json({ message: 'Session error' });
            }
            console.log('OWNER LOGIN SUCCESS:', email);
            return res.json({
                message: 'Admin login successful',
                redirect: '/admin/panel'
            });
        });

    } catch (error) {
        console.error('ADMIN LOGIN ERROR:', error);
        return res.status(500).json({ message: 'Server error' });
    }
});

function requireAdmin(req, res, next) {
    if (!req.session || req.session.isAdmin !== true) {
        return res.status(401).json({ message: 'Admin authentication required' });
    }
    next();
}

app.get('/admin/panel', (req, res) => {
    if (!req.session || req.session.isAdmin !== true) {
        return res.redirect('/owner-login');
    }
    res.set('X-Robots-Tag', 'noindex, nofollow');
    return res.sendFile(path.join(__dirname, 'private', 'admin.html'));
});

// ----------------------------------------------------------
// GET REGISTERED USERS
// ----------------------------------------------------------

app.get('/admin/api/users', requireAdmin, async (req, res) => {
    try {
        if (!fs.existsSync(EXCEL_FILE)) {
            return res.json({ total: 0, users: [] });
        }

        const { worksheet } = await readWorkbook();
        const users = [];

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;

            users.push({
                name: String(row.getCell(1).value || ''),
                email: String(row.getCell(2).value || ''),
                regulation: String(row.getCell(3).value || ''),
                rollNumber: String(row.getCell(4).value || ''),
                branch: String(row.getCell(5).value || ''),
                yearStudy: String(row.getCell(6).value || ''),
                collegeName: String(row.getCell(7).value || ''),
                phone: String(row.getCell(8).value || '')
            });
        });

        return res.json({ total: users.length, users });

    } catch (error) {
        console.error('ADMIN USERS ERROR:', error);
        return res.status(500).json({ message: 'Could not read users' });
    }
});

// ----------------------------------------------------------
// DOWNLOAD USERS EXCEL FILE
// ----------------------------------------------------------

app.get('/admin/download-users', (req, res) => {
    if (!req.session || req.session.isAdmin !== true) {
        return res.status(401).send('Admin authentication required');
    }

    if (!fs.existsSync(EXCEL_FILE)) {
        return res.status(404).send('Users Excel file not found');
    }

    return res.download(
        EXCEL_FILE,
        'JNTUA-Academic-Hub-Users.xlsx',
        (error) => {
            if (error) console.error('EXCEL DOWNLOAD ERROR:', error);
        }
    );
});

app.get('/admin/logout', (req, res) => {
    req.session.isAdmin = false;
    req.session.adminEmail = null;

    req.session.save((saveError) => {
        if (saveError) console.error('ADMIN LOGOUT ERROR:', saveError);
        return res.json({ message: 'Admin logged out' });
    });
});

// ==========================================================
// 404
// ==========================================================

app.use((req, res) => {
    console.log('404:', req.method, req.originalUrl);
    return res.status(404).send('Page not found');
});

// ==========================================================
// START SERVER
// ==========================================================

const PORT = process.env.PORT || 3000;

app.listen(PORT, '0.0.0.0', () => {
    console.log('====================================');
    console.log('JNTUA ACADEMIC HUB SERVER RUNNING');
    console.log(`PORT: ${PORT}`);
    console.log(`EXCEL: ${EXCEL_FILE}`);
    console.log('====================================');
});
