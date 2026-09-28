// server.js
const express = require('express');
const path = require('path');
const session = require('express-session');
const ExcelJS = require('exceljs');
const fs = require('fs');

const app = express();
app.use(express.json());


// Session setup
app.use(session({
    secret: 'jntua-secret',
    resave: false,
    saveUninitialized: true
}));

const EXCEL_FILE = path.join(__dirname, 'users.xlsx');

// Helper: add user to Excel safely
async function addUserToExcel(user) {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    // If file exists, read it; else create new workbook
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

    // Check duplicate email
    let emailExists = false;
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber !== 1) {
            if (row.getCell(2).value === user.email) {
                emailExists = true;
            }
        }
    });

    if (emailExists) throw new Error('Email already registered');

    // Add new user
    worksheet.addRow([
        user.name,
        user.email,
        user.regulation || '',
        user.rollNumber || '',
        user.branch || '',
        user.yearStudy || '',     // NEW
        user.collegeName || '',   // NEW
        user.phone || '',
        user.password
    ]);

    await workbook.xlsx.writeFile(EXCEL_FILE);
}

// Registration
app.post('/register', async (req, res) => {
    try {
        const user = req.body;

        if (!user.name || !user.email || !user.password || !user.confirmPassword) {
            return res.status(400).json({ message: 'Fill all required fields' });
        }

        if (user.password !== user.confirmPassword) {
            return res.status(400).json({ message: 'Passwords do not match' });
        }

        await addUserToExcel(user);

        res.json({ message: 'Registered successfully' });

    } catch (err) {
        console.log(err.message);

        if (err.message === 'Email already registered') {
            res.status(400).json({ message: 'Email already registered' });
        } else {
            res.status(500).json({ message: 'Server error. Try later.' });
        }
    }
});

// Login
app.post('/login', async (req, res) => {
    try {
        const { email, password } = req.body;

        if (!fs.existsSync(EXCEL_FILE)) {
            return res.status(400).json({ message: 'No users registered yet' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(EXCEL_FILE);
        const worksheet = workbook.getWorksheet(1);

        let userRow = null;

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber !== 1 && row.getCell(2).value === email) {
                userRow = row;
            }
        });

        if (!userRow) return res.status(400).json({ message: 'Email not registered' });

        // Password column changed from 7 → 9
        if (userRow.getCell(9).value !== password) {
            return res.status(400).json({ message: 'Incorrect password' });
        }

        req.session.user = email;

        res.json({ message: 'Login successful' });

    } catch (err) {
        console.log(err);
        res.status(500).json({ message: 'Server error. Try later.' });
    }
});


app.get('/', (req, res) => {
    res.sendFile(
        path.join(__dirname, 'public', 'login.html')
    );
});

// Protected main page
app.get('/main-portal.html', (req, res) => {
    if (!req.session.user) return res.redirect('/');
    res.sendFile(path.join(__dirname, 'public', 'main-portal.html'));
});

// Logout
app.get('/logout', (req, res) => {
    req.session.destroy(() => res.redirect('/'));
});

// Start server
app.listen(3000, () => console.log('Server running on http://localhost:3000'));


