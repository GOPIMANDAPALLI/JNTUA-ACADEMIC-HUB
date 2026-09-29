// ==========================================================
// JNTUA ACADEMIC HUB - server.js (persistent version)
// ==========================================================

const express = require('express');
const path = require('path');
const session = require('express-session');
const pgSession = require('connect-pg-simple')(session);
const { Pool } = require('pg');
const ExcelJS = require('exceljs');
const fs = require('fs');
const bcrypt = require('bcrypt');

const app = express();
app.set('trust proxy', 1);

const PUBLIC_DIR = path.join(__dirname, 'public');

// ==========================================================
// DATABASE (Supabase Postgres)
// ==========================================================

if (!process.env.DATABASE_URL) {
    console.error('DATABASE_URL is missing in environment variables');
}

const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false }
});

pool.on('error', (err) => console.error('DB POOL ERROR:', err));

// ==========================================================
// MIDDLEWARE
// ==========================================================

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ==========================================================
// SESSION (stored in database, survives restarts)
// ==========================================================

app.use(session({
    store: new pgSession({
        pool,
        tableName: 'user_sessions',
        createTableIfMissing: true
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

        if (user.password !== user.confirmPassword) {
            return res.status(400).json({ message: 'Passwords do not match' });
        }

        const email = String(user.email).trim().toLowerCase();
        const passwordHash = await bcrypt.hash(String(user.password), 10);

        await pool.query(
            `INSERT INTO users
             (name, email, regulation, roll_number, branch, year_study,
              college_name, phone, password_hash)
             VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)`,
            [
                user.name || '',
                email,
                user.regulation || '',
                user.rollNumber || '',
                user.branch || '',
                user.yearStudy || '',
                user.collegeName || '',
                user.phone || '',
                passwordHash
            ]
        );

        console.log('User registered:', email);
        return res.json({ message: 'Registered successfully' });

    } catch (err) {
        // 23505 = unique violation (email already exists)
        if (err.code === '23505') {
            return res.status(400).json({ message: 'Email already registered' });
        }
        console.error('REGISTER ERROR:', err);
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

        const result = await pool.query(
            'SELECT password_hash FROM users WHERE email = $1',
            [email]
        );

        if (result.rows.length === 0) {
            return res.status(400).json({ message: 'Email not registered' });
        }

        const ok = await bcrypt.compare(password, result.rows[0].password_hash);

        if (!ok) {
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

// ==========================================================
// CURRENT USER
// ==========================================================

app.get('/api/current-user', (req, res) => {
    if (!req.session || !req.session.user) {
        return res.status(401).json({ loggedIn: false });
    }
    return res.json({ loggedIn: true, email: req.session.user });
});

// ==========================================================
// LOGOUT
// ==========================================================

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
        const result = await pool.query(
            `SELECT name, email, regulation,
                    roll_number  AS "rollNumber",
                    branch,
                    year_study   AS "yearStudy",
                    college_name AS "collegeName",
                    phone
             FROM users
             ORDER BY id`
        );

        return res.json({
            total: result.rows.length,
            users: result.rows
        });

    } catch (error) {
        console.error('ADMIN USERS ERROR:', error);
        return res.status(500).json({ message: 'Could not read users' });
    }
});

// ----------------------------------------------------------
// DOWNLOAD USERS AS EXCEL (generated from database)
// ----------------------------------------------------------

app.get('/admin/download-users', async (req, res) => {
    if (!req.session || req.session.isAdmin !== true) {
        return res.status(401).send('Admin authentication required');
    }

    try {
        const result = await pool.query(
            `SELECT name, email, regulation, roll_number, branch,
                    year_study, college_name, phone
             FROM users ORDER BY id`
        );

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Users');

        sheet.addRow([
            'Name', 'Email', 'Regulation', 'RollNumber',
            'Branch', 'YearOfStudy', 'CollegeName', 'Phone'
        ]);

        result.rows.forEach((u) => {
            sheet.addRow([
                u.name, u.email, u.regulation, u.roll_number,
                u.branch, u.year_study, u.college_name, u.phone
            ]);
        });

        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            'attachment; filename="JNTUA-Academic-Hub-Users.xlsx"'
        );

        await workbook.xlsx.write(res);
        return res.end();

    } catch (error) {
        console.error('EXCEL DOWNLOAD ERROR:', error);
        return res.status(500).send('Could not generate Excel file');
    }
});

// ----------------------------------------------------------
// ADMIN LOGOUT
// ----------------------------------------------------------

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
    console.log('====================================');
});
