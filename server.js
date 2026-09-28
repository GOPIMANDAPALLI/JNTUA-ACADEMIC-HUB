// ==========================================================
// JNTUA ACADEMIC HUB - server.js
// ==========================================================

const express = require('express');
const path = require('path');
const session = require('express-session');
const ExcelJS = require('exceljs');
const bcrypt = require('bcrypt');
const fs = require('fs');

const app = express();

// ==========================================================
// PATHS
// ==========================================================

const PUBLIC_DIR = path.join(__dirname, 'public');

// Local:
//     ./data/users.xlsx
//
// Render:
//     DATA_DIR=/var/data
//
// DATA_DIR should NOT be inside public folder.

const DATA_DIR =
    process.env.DATA_DIR || path.join(__dirname, 'data');

const EXCEL_FILE =
    path.join(DATA_DIR, 'users.xlsx');

// Create data folder automatically
if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

// ==========================================================
// MIDDLEWARE
// ==========================================================

app.use(express.json());

app.use(express.urlencoded({
    extended: true
}));

// Serve website files
app.use(express.static(PUBLIC_DIR));

// ==========================================================
// SESSION
// ==========================================================

app.use(session({

    secret:
        process.env.SESSION_SECRET ||
        'jntua-secret-change-this',

    resave: false,

    saveUninitialized: false,

    cookie: {
        httpOnly: true,

        // Local HTTP
        // For Render HTTPS, set secure: true
        secure: process.env.NODE_ENV === 'production',

        sameSite: 'lax',

        // 30 days
        maxAge: 1000 * 60 * 60 * 24 * 30
    }

}));

// ==========================================================
// EXCEL HEADERS
// ==========================================================

const HEADERS = [
    'Name',
    'Email',
    'Regulation',
    'RollNumber',
    'Branch',
    'YearOfStudy',
    'CollegeName',
    'Phone',
    'PasswordHash',
    'RegisteredAt'
];

// ==========================================================
// CREATE EXCEL FILE
// ==========================================================

async function createExcelFile() {

    if (fs.existsSync(EXCEL_FILE)) {
        return;
    }

    const workbook = new ExcelJS.Workbook();

    const worksheet =
        workbook.addWorksheet('Users');

    worksheet.addRow(HEADERS);

    // Header formatting
    worksheet.getRow(1).font = {
        bold: true
    };

    await workbook.xlsx.writeFile(EXCEL_FILE);

    console.log(
        'Created Excel file:',
        EXCEL_FILE
    );
}

// ==========================================================
// VALIDATE EMAIL
// ==========================================================

function isValidEmail(email) {

    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);

}

// ==========================================================
// FIND USER BY EMAIL
// ==========================================================

async function findUserByEmail(email) {

    if (!fs.existsSync(EXCEL_FILE)) {
        return null;
    }

    const workbook =
        new ExcelJS.Workbook();

    await workbook.xlsx.readFile(
        EXCEL_FILE
    );

    const worksheet =
        workbook.getWorksheet('Users') ||
        workbook.getWorksheet(1);

    if (!worksheet) {
        return null;
    }

    let foundUser = null;

    worksheet.eachRow(
        (row, rowNumber) => {

            if (rowNumber === 1) {
                return;
            }

            const storedEmail =
                String(
                    row.getCell(2).value || ''
                )
                .trim()
                .toLowerCase();

            if (storedEmail === email) {

                foundUser = row;

            }

        }
    );

    return foundUser;
}

// ==========================================================
// ADD USER TO EXCEL
// ==========================================================

async function addUserToExcel(user) {

    await createExcelFile();

    const workbook =
        new ExcelJS.Workbook();

    await workbook.xlsx.readFile(
        EXCEL_FILE
    );

    let worksheet =
        workbook.getWorksheet('Users');

    if (!worksheet) {

        worksheet =
            workbook.addWorksheet('Users');

        worksheet.addRow(HEADERS);

    }

    const newEmail =
        String(user.email || '')
        .trim()
        .toLowerCase();

    // Check duplicate email
    const existingUser =
        await findUserByEmail(newEmail);

    if (existingUser) {

        throw new Error(
            'Email already registered'
        );

    }

    // Hash password
    const passwordHash =
        await bcrypt.hash(
            user.password,
            12
        );

    // Add user
    worksheet.addRow([

        user.name || '',

        newEmail,

        user.regulation || '',

        user.rollNumber || '',

        user.branch || '',

        user.yearStudy || '',

        user.collegeName || '',

        user.phone || '',

        passwordHash,

        new Date().toISOString()

    ]);

    await workbook.xlsx.writeFile(
        EXCEL_FILE
    );

    console.log(
        'USER REGISTERED:',
        newEmail
    );
}

// ==========================================================
// HOME ROUTE
// ==========================================================

app.get('/', (req, res) => {

    console.log('--------------------------');

    console.log('GET /');

    console.log(
        'Session user:',
        req.session
            ? req.session.user || 'none'
            : 'none'
    );

    // Already logged in
    if (
        req.session &&
        req.session.user
    ) {

        const mainPage =
            path.join(
                PUBLIC_DIR,
                'main-portal.html'
            );

        if (!fs.existsSync(mainPage)) {

            return res.status(500).send(
                'main-portal.html not found inside public folder.'
            );

        }

        return res.sendFile(mainPage);

    }

    // Not logged in
    const loginPage =
        path.join(
            PUBLIC_DIR,
            'login.html'
        );

    if (!fs.existsSync(loginPage)) {

        return res.status(500).send(
            'login.html not found inside public folder.'
        );

    }

    return res.sendFile(loginPage);

});

// ==========================================================
// REGISTER
// ==========================================================

app.post('/register', async (req, res) => {

    try {

        const user =
            req.body || {};

        // Required fields
        if (
            !user.name ||
            !user.email ||
            !user.password ||
            !user.confirmPassword
        ) {

            return res.status(400).json({

                message:
                    'Fill all required fields'

            });

        }

        // Clean email
        user.email =
            String(user.email)
            .trim()
            .toLowerCase();

        // Email validation
        if (!isValidEmail(user.email)) {

            return res.status(400).json({

                message:
                    'Invalid email ID'

            });

        }

        // Password confirmation
        if (
            user.password !==
            user.confirmPassword
        ) {

            return res.status(400).json({

                message:
                    'Passwords do not match'

            });

        }

        // Password minimum length
        if (user.password.length < 6) {

            return res.status(400).json({

                message:
                    'Password must be at least 6 characters'

            });

        }

        // Add to Excel
        await addUserToExcel(user);

        return res.json({

            message:
                'Registered successfully'

        });

    }

    catch (err) {

        console.error(
            'REGISTER ERROR:',
            err
        );

        if (
            err.message ===
            'Email already registered'
        ) {

            return res.status(400).json({

                message:
                    'Email already registered'

            });

        }

        return res.status(500).json({

            message:
                'Server error. Try later.'

        });

    }

});

// ==========================================================
// LOGIN
// ==========================================================

app.post('/login', async (req, res) => {

    try {

        const email =
            String(
                req.body.email || ''
            )
            .trim()
            .toLowerCase();

        const password =
            String(
                req.body.password || ''
            );

        // Empty fields
        if (!email || !password) {

            return res.status(400).json({

                message:
                    'Enter email and password'

            });

        }

        // Invalid email format
        if (!isValidEmail(email)) {

            return res.status(400).json({

                message:
                    'Invalid email ID'

            });

        }

        // Excel doesn't exist
        if (!fs.existsSync(EXCEL_FILE)) {

            return res.status(400).json({

                message:
                    'Email not registered'

            });

        }

        // Find email
        const userRow =
            await findUserByEmail(email);

        // Email doesn't exist
        if (!userRow) {

            return res.status(400).json({

                message:
                    'Email not registered'

            });

        }

        // Password hash
        const storedPasswordHash =
            String(
                userRow.getCell(9).value || ''
            );

        // Compare password
        const passwordMatches =
            await bcrypt.compare(
                password,
                storedPasswordHash
            );

        // Wrong password
        if (!passwordMatches) {

            return res.status(400).json({

                message:
                    'Incorrect password'

            });

        }

        // ==================================================
        // LOGIN SUCCESS
        // ==================================================

        req.session.user = email;

        console.log(
            'LOGIN SUCCESS:',
            email
        );

        // Save session before response
        req.session.save(
            (sessionError) => {

                if (sessionError) {

                    console.error(
                        'SESSION ERROR:',
                        sessionError
                    );

                    return res.status(500).json({

                        message:
                            'Session error. Please try again.'

                    });

                }

                console.log(
                    'SESSION SAVED:',
                    email
                );

                return res.json({

                    message:
                        'Login successful',

                    redirect:
                        '/'

                });

            }
        );

    }

    catch (err) {

        console.error(
            'LOGIN ERROR:',
            err
        );

        return res.status(500).json({

            message:
                'Server error. Try later.'

        });

    }

});

// ==========================================================
// PROTECTED MAIN PORTAL
// ==========================================================

app.get(
    '/main-portal.html',
    (req, res) => {

        console.log(
            'MAIN PORTAL REQUEST'
        );

        // Not logged in
        if (
            !req.session ||
            !req.session.user
        ) {

            console.log(
                'NO SESSION -> LOGIN'
            );

            return res.redirect('/');

        }

        const mainPage =
            path.join(
                PUBLIC_DIR,
                'main-portal.html'
            );

        if (!fs.existsSync(mainPage)) {

            return res.status(500).send(
                'main-portal.html not found inside public folder.'
            );

        }

        return res.sendFile(
            mainPage
        );

    }
);

// ==========================================================
// CURRENT USER
// ==========================================================

app.get(
    '/api/current-user',
    (req, res) => {

        if (
            !req.session ||
            !req.session.user
        ) {

            return res.status(401).json({

                loggedIn: false

            });

        }

        return res.json({

            loggedIn: true,

            email:
                req.session.user

        });

    }
);

// ==========================================================
// LOGOUT
// ==========================================================

app.get(
    '/logout',
    (req, res) => {

        req.session.destroy(
            (err) => {

                if (err) {

                    console.error(
                        'LOGOUT ERROR:',
                        err
                    );

                    return res.status(500).send(
                        'Logout failed'
                    );

                }

                res.clearCookie(
                    'connect.sid'
                );

                return res.redirect('/');

            }
        );

    }
);

// ==========================================================
// 404
// ==========================================================

app.use(
    (req, res) => {

        console.log(
            '404:',
            req.method,
            req.originalUrl
        );

        return res.status(404).send(
            'Page not found'
        );

    }
);

// ==========================================================
// START SERVER
// ==========================================================

const PORT =
    process.env.PORT || 3000;

app.listen(
    PORT,
    '0.0.0.0',
    () => {

        console.log(
            '===================================='
        );

        console.log(
            'JNTUA ACADEMIC HUB SERVER RUNNING'
        );

        console.log(
            `PORT: ${PORT}`
        );

        console.log(
            `URL: http://localhost:${PORT}`
        );

        console.log(
            `EXCEL: ${EXCEL_FILE}`
        );

        console.log(
            '===================================='
        );

    }
);
