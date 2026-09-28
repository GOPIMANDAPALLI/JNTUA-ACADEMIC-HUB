// server.js

const express = require('express');
const path = require('path');
const session = require('express-session');
const ExcelJS = require('exceljs');
const fs = require('fs');

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
// DATA DIRECTORY
// ==================================================

// Local: ./data
// Render: /var/data through DATA_DIR environment variable

const DATA_DIR =
    process.env.DATA_DIR || path.join(__dirname, 'data');

if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

const EXCEL_FILE = path.join(DATA_DIR, 'users.xlsx');

console.log('Excel file location:', EXCEL_FILE);

// ==================================================
// HOME PAGE
// ==================================================

app.get('/', (req, res) => {

    // If user is already logged in
    // show main portal without changing URL
    if (req.session.user) {

        return res.sendFile(
            path.join(__dirname, 'public', 'main-portal.html')
        );
    }

    // If user is not logged in
    // show login page
    res.sendFile(
        path.join(__dirname, 'public', 'login.html')
    );
});

// ==================================================
// ADD USER TO EXCEL
// ==================================================

async function addUserToExcel(user) {

    const workbook = new ExcelJS.Workbook();
    let worksheet;

    // Existing Excel file
    if (fs.existsSync(EXCEL_FILE)) {

        await workbook.xlsx.readFile(EXCEL_FILE);

        worksheet = workbook.getWorksheet(1);

        if (!worksheet) {

            worksheet = workbook.addWorksheet('Users');

            worksheet.addRow([
                'Name',
                'Email',
                'Regulation',
                'RollNumber',
                'Branch',
                'YearOfStudy',
                'CollegeName',
                'Phone',
                'Password'
            ]);
        }

    } else {

        // Create new Excel file
        worksheet = workbook.addWorksheet('Users');

        worksheet.addRow([
            'Name',
            'Email',
            'Regulation',
            'RollNumber',
            'Branch',
            'YearOfStudy',
            'CollegeName',
            'Phone',
            'Password'
        ]);
    }

    // ==================================================
    // CHECK DUPLICATE EMAIL
    // ==================================================

    let emailExists = false;

    worksheet.eachRow((row, rowNumber) => {

        if (rowNumber !== 1) {

            const existingEmail =
                row.getCell(2).value;

            if (existingEmail === user.email) {
                emailExists = true;
            }
        }
    });

    if (emailExists) {
        throw new Error('Email already registered');
    }

    // ==================================================
    // ADD USER
    // ==================================================

    worksheet.addRow([
        user.name,
        user.email,
        user.regulation || '',
        user.rollNumber || '',
        user.branch || '',
        user.yearStudy || '',
        user.collegeName || '',
        user.phone || '',
        user.password
    ]);

    // Save Excel
    await workbook.xlsx.writeFile(EXCEL_FILE);

    console.log(
        'User saved successfully:',
        user.email
    );
}

// ==================================================
// REGISTER
// ==================================================

app.post('/register', async (req, res) => {

    try {

        const user = req.body;

        console.log(
            'Registration request:',
            user.email
        );

        // Required fields
        if (
            !user.name ||
            !user.email ||
            !user.password ||
            !user.confirmPassword
        ) {

            return res.status(400).json({
                message: 'Fill all required fields'
            });
        }

        // Password confirmation
        if (
            user.password !==
            user.confirmPassword
        ) {

            return res.status(400).json({
                message: 'Passwords do not match'
            });
        }

        // Save user
        await addUserToExcel(user);

        res.json({
            message: 'Registered successfully'
        });

    } catch (err) {

        console.log(
            'Registration error:',
            err
        );

        if (
            err.message ===
            'Email already registered'
        ) {

            return res.status(400).json({
                message: 'Email already registered'
            });
        }

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

        const {
            email,
            password
        } = req.body;

        console.log(
            'Login attempt:',
            email
        );

        // Check Excel
        if (!fs.existsSync(EXCEL_FILE)) {

            return res.status(400).json({
                message: 'No users registered yet'
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

                if (
                    rowNumber !== 1 &&
                    row.getCell(2).value === email
                ) {

                    userRow = row;
                }
            }
        );

        // Email not found
        if (!userRow) {

            return res.status(400).json({
                message: 'Email not registered'
            });
        }

        // Password is column 9
        if (
            userRow.getCell(9).value !== password
        ) {

            return res.status(400).json({
                message: 'Incorrect password'
            });
        }

        // Create session
        req.session.user = email;

        // IMPORTANT:
        // Do NOT send main-portal.html here.
        // Frontend should redirect to "/".
        res.json({
            message: 'Login successful',
            redirect: '/'
        });

    } catch (err) {

        console.log(
            'Login error:',
            err
        );

        res.status(500).json({
            message: 'Server error. Try later.'
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
            console.log(
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
            `Excel file: ${EXCEL_FILE}`
        );
    }
);
