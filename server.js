// server.js

const express = require('express');
const path = require('path');
const session = require('express-session');
const ExcelJS = require('exceljs');
const fs = require('fs');
const bcrypt = require('bcrypt');

const app = express();

// ==================================================
// MIDDLEWARE
// ==================================================

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve public folder
app.use(express.static(path.join(__dirname, 'public')));

// ==================================================
// SESSION
// ==================================================

app.use(session({
    secret: process.env.SESSION_SECRET || 'jntua-secret',
    resave: false,
    saveUninitialized: false,
    cookie: {
        httpOnly: true,
        secure: false
    }
}));

// ==================================================
// DATA STORAGE
// ==================================================

// Render Persistent Disk:
// /var/data
//
// Local testing:
// ./data

const DATA_DIR =
    process.env.DATA_DIR || path.join(__dirname, 'data');

if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

const EXCEL_FILE =
    path.join(DATA_DIR, 'users.xlsx');

console.log('Excel file:', EXCEL_FILE);

// ==================================================
// HOME PAGE
// ==================================================

app.get('/', (req, res) => {

    // Already logged in
    if (req.session.user) {

        return res.sendFile(
            path.join(
                __dirname,
                'public',
                'main-portal.html'
            )
        );
    }

    // Not logged in
    res.sendFile(
        path.join(
            __dirname,
            'public',
            'login.html'
        )
    );
});

// ==================================================
// CREATE / OPEN EXCEL
// ==================================================

async function getWorksheet() {

    const workbook = new ExcelJS.Workbook();

    let worksheet;

    if (fs.existsSync(EXCEL_FILE)) {

        await workbook.xlsx.readFile(
            EXCEL_FILE
        );

        worksheet =
            workbook.getWorksheet(1);

    } else {

        worksheet =
            workbook.addWorksheet('Users');

        worksheet.addRow([
            'Name',
            'Email',
            'Regulation',
            'RollNumber',
            'Branch',
            'YearOfStudy',
            'CollegeName',
            'Phone',
            'PasswordHash'
        ]);

        await workbook.xlsx.writeFile(
            EXCEL_FILE
        );
    }

    if (!worksheet) {

        worksheet =
            workbook.addWorksheet('Users');

        worksheet.addRow([
            'Name',
            'Email',
            'Regulation',
            'RollNumber',
            'Branch',
            'YearOfStudy',
            'CollegeName',
            'Phone',
            'PasswordHash'
        ]);

        await workbook.xlsx.writeFile(
            EXCEL_FILE
        );
    }

    return {
        workbook,
        worksheet
    };
}

// ==================================================
// REGISTER
// ==================================================

app.post('/register', async (req, res) => {

    try {

        const user = req.body;

        const name =
            user.name?.trim();

        const email =
            user.email?.trim().toLowerCase();

        const password =
            user.password;

        const confirmPassword =
            user.confirmPassword;

        console.log(
            'Registration:',
            email
        );

        // Required fields
        if (
            !name ||
            !email ||
            !password ||
            !confirmPassword
        ) {

            return res.status(400).json({
                message: 'Fill all required fields'
            });
        }

        // Password match
        if (
            password !== confirmPassword
        ) {

            return res.status(400).json({
                message: 'Passwords do not match'
            });
        }

        // ==================================================
        // OPEN EXCEL
        // ==================================================

        const {
            workbook,
            worksheet
        } = await getWorksheet();

        // ==================================================
        // CHECK EMAIL
        // ==================================================

        let emailExists = false;

        worksheet.eachRow(
            (row, rowNumber) => {

                if (rowNumber === 1) {
                    return;
                }

                const existingEmail =
                    String(
                        row.getCell(2).value || ''
                    )
                    .trim()
                    .toLowerCase();

                if (
                    existingEmail === email
                ) {
                    emailExists = true;
                }
            }
        );

        if (emailExists) {

            return res.status(400).json({
                message: 'Email already registered'
            });
        }

        // ==================================================
        // HASH PASSWORD
        // ==================================================

        const passwordHash =
            await bcrypt.hash(
                password,
                12
            );

        // ==================================================
        // ADD USER
        // ==================================================

        worksheet.addRow([
            name,
            email,
            user.regulation || '',
            user.rollNumber || '',
            user.branch || '',
            user.yearStudy || '',
            user.collegeName || '',
            user.phone || '',
            passwordHash
        ]);

        // Save Excel
        await workbook.xlsx.writeFile(
            EXCEL_FILE
        );

        console.log(
            'User saved:',
            email
        );

        res.json({
            message: 'Registered successfully'
        });

    } catch (err) {

        console.error(
            'Registration error:',
            err
        );

        res.status(500).json({
            message: 'Server error. Try later.'
        });
    }
});

// ==================================================
// LOGIN
// ==================================================

app.post('/login', async (req, res) => {

    try {

        const email =
            req.body.email?.trim().toLowerCase();

        const password =
            req.body.password;

        console.log(
            'Login:',
            email
        );

        if (!email || !password) {

            return res.status(400).json({
                message:
                    'Enter email and password'
            });
        }

        // ==================================================
        // CHECK EXCEL
        // ==================================================

        if (!fs.existsSync(EXCEL_FILE)) {

            return res.status(400).json({
                message:
                    'Email not registered'
            });
        }

        const workbook =
            new ExcelJS.Workbook();

        await workbook.xlsx.readFile(
            EXCEL_FILE
        );

        const worksheet =
            workbook.getWorksheet(1);

        let userRow = null;

        worksheet.eachRow(
            (row, rowNumber) => {

                if (rowNumber === 1) {
                    return;
                }

                const existingEmail =
                    String(
                        row.getCell(2).value || ''
                    )
                    .trim()
                    .toLowerCase();

                if (
                    existingEmail === email
                ) {

                    userRow = row;
                }
            }
        );

        // ==================================================
        // EMAIL NOT FOUND
        // ==================================================

        if (!userRow) {

            return res.status(400).json({
                message:
                    'Email not registered'
            });
        }

        // ==================================================
        // CHECK PASSWORD
        // ==================================================

        const passwordHash =
            userRow.getCell(9).value;

        const passwordCorrect =
            await bcrypt.compare(
                password,
                passwordHash
            );

        if (!passwordCorrect) {

            return res.status(400).json({
                message:
                    'Incorrect password'
            });
        }

        // ==================================================
        // LOGIN SUCCESS
        // ==================================================

        req.session.user = email;

        res.json({
            message:
                'Login successful',
            redirect: '/'
        });

    } catch (err) {

        console.error(
            'Login error:',
            err
        );

        res.status(500).json({
            message:
                'Server error. Try later.'
        });
    }
});

// ==================================================
// PROTECTED MAIN PORTAL
// ==================================================

app.get('/main-portal.html', (req, res) => {

    if (!req.session.user) {

        return res.redirect('/');
    }

    res.sendFile(
        path.join(
            __dirname,
            'public',
            'main-portal.html'
        )
    );
});

// ==================================================
// LOGOUT
// ==================================================

app.get('/logout', (req, res) => {

    req.session.destroy((err) => {

        if (err) {
            console.error(
                'Logout error:',
                err
            );
        }

        res.redirect('/');
    });
});

// ==================================================
// START SERVER
// ==================================================

const PORT =
    process.env.PORT || 3000;

app.listen(
    PORT,
    '0.0.0.0',
    () => {

        console.log(
            `Server running on port ${PORT}`
        );

        console.log(
            `Data directory: ${DATA_DIR}`
        );

        console.log(
            `Excel file: ${EXCEL_FILE}`
        );
    }
);
