// ==================================================
// JNTUA ACADEMIC HUB - SERVER.JS
// ==================================================

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
app.use(
    express.static(
        path.join(__dirname, 'public')
    )
);

// ==================================================
// SESSION
// ==================================================

app.use(
    session({
        secret:
            process.env.SESSION_SECRET ||
            'jntua-academic-hub-secret-2026',

        resave: false,

        saveUninitialized: false,

        cookie: {
            httpOnly: true,
            secure: false,
            sameSite: 'lax'
        }
    })
);

// ==================================================
// DATA DIRECTORY
// ==================================================

// Local:
// ./data
//
// Render:
// /var/data
//
// Render Environment Variable:
// DATA_DIR=/var/data

const DATA_DIR =
    process.env.DATA_DIR ||
    path.join(__dirname, 'data');

// Create data directory if it doesn't exist
if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, {
        recursive: true
    });
}

// Excel file
const EXCEL_FILE =
    path.join(
        DATA_DIR,
        'users.xlsx'
    );

console.log('====================================');
console.log('DATA DIRECTORY:');
console.log(DATA_DIR);

console.log('EXCEL FILE:');
console.log(EXCEL_FILE);

console.log('====================================');

// ==================================================
// CREATE NEW EXCEL FILE
// ==================================================

async function createUsersExcel() {

    const workbook =
        new ExcelJS.Workbook();

    const worksheet =
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

    console.log(
        'New users.xlsx created.'
    );
}

// ==================================================
// GET USERS WORKSHEET
// ==================================================

async function getUsersWorksheet() {

    // Create Excel if it doesn't exist
    if (!fs.existsSync(EXCEL_FILE)) {

        await createUsersExcel();
    }

    const workbook =
        new ExcelJS.Workbook();

    await workbook.xlsx.readFile(
        EXCEL_FILE
    );

    let worksheet =
        workbook.getWorksheet('Users');

    // If Users sheet doesn't exist,
    // use first sheet
    if (!worksheet) {

        worksheet =
            workbook.worksheets[0];
    }

    // If there is no sheet at all
    if (!worksheet) {

        worksheet =
            workbook.addWorksheet(
                'Users'
            );

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
// HOME PAGE
// ==================================================

app.get('/', (req, res) => {

    try {

        // User already logged in
        if (
            req.session &&
            req.session.user
        ) {

            return res.sendFile(
                path.join(
                    __dirname,
                    'public',
                    'main-portal.html'
                )
            );
        }

        // User not logged in
        return res.sendFile(
            path.join(
                __dirname,
                'public',
                'login.html'
            )
        );

    } catch (error) {

        console.error(
            'HOME PAGE ERROR:',
            error
        );

        return res.status(500).send(
            'Unable to open website.'
        );
    }
});

// ==================================================
// REGISTER
// ==================================================

app.post(
    '/register',
    async (req, res) => {

        try {

            console.log('');
            console.log(
                '===================================='
            );
            console.log(
                'NEW REGISTRATION REQUEST'
            );
            console.log(
                '===================================='
            );

            const body =
                req.body || {};

            // ------------------------------------------
            // READ FORM VALUES
            // ------------------------------------------

            const name =
                String(
                    body.name || ''
                ).trim();

            const email =
                String(
                    body.email || ''
                )
                    .trim()
                    .toLowerCase();

            const password =
                String(
                    body.password || ''
                );

            const confirmPassword =
                String(
                    body.confirmPassword || ''
                );

            const regulation =
                String(
                    body.regulation || ''
                ).trim();

            const rollNumber =
                String(
                    body.rollNumber || ''
                ).trim();

            const branch =
                String(
                    body.branch || ''
                ).trim();

            const yearStudy =
                String(
                    body.yearStudy || ''
                ).trim();

            const collegeName =
                String(
                    body.collegeName || ''
                ).trim();

            const phone =
                String(
                    body.phone || ''
                ).trim();

            console.log(
                'Registration email:',
                email
            );

            // ------------------------------------------
            // REQUIRED FIELDS
            // ------------------------------------------

            if (
                !name ||
                !email ||
                !password ||
                !confirmPassword
            ) {

                return res.status(400).json({
                    message:
                        'Please fill all required fields.'
                });
            }

            // ------------------------------------------
            // BASIC EMAIL CHECK
            // ------------------------------------------

            const emailPattern =
                /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

            if (
                !emailPattern.test(email)
            ) {

                return res.status(400).json({
                    message:
                        'Please enter a valid Gmail/email address.'
                });
            }

            // ------------------------------------------
            // PASSWORD LENGTH
            // ------------------------------------------

            if (
                password.length < 6
            ) {

                return res.status(400).json({
                    message:
                        'Password must contain at least 6 characters.'
                });
            }

            // ------------------------------------------
            // CONFIRM PASSWORD
            // ------------------------------------------

            if (
                password !==
                confirmPassword
            ) {

                return res.status(400).json({
                    message:
                        'Passwords do not match.'
                });
            }

            // ------------------------------------------
            // GET EXCEL
            // ------------------------------------------

            const {
                workbook,
                worksheet
            } =
                await getUsersWorksheet();

            // ------------------------------------------
            // CHECK DUPLICATE EMAIL
            // ------------------------------------------

            let emailExists = false;

            worksheet.eachRow(
                (row, rowNumber) => {

                    if (
                        rowNumber === 1
                    ) {
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
                    message:
                        'Email already registered.'
                });
            }

            // ------------------------------------------
            // HASH PASSWORD
            // ------------------------------------------

            const passwordHash =
                await bcrypt.hash(
                    password,
                    12
                );

            // ------------------------------------------
            // ADD USER
            // ------------------------------------------

            worksheet.addRow([
                name,
                email,
                regulation,
                rollNumber,
                branch,
                yearStudy,
                collegeName,
                phone,
                passwordHash
            ]);

            // ------------------------------------------
            // SAVE EXCEL
            // ------------------------------------------

            await workbook.xlsx.writeFile(
                EXCEL_FILE
            );

            console.log(
                'USER REGISTERED SUCCESSFULLY:',
                email
            );

            console.log(
                'Excel saved at:',
                EXCEL_FILE
            );

            console.log(
                '===================================='
            );

            return res.status(200).json({
                message:
                    'Registered successfully.'
            });

        } catch (error) {

            console.error('');
            console.error(
                '===================================='
            );
            console.error(
                'REGISTRATION ERROR'
            );
            console.error(
                error
            );
            console.error(
                '===================================='
            );

            return res.status(500).json({
                message:
                    'Registration failed: ' +
                    error.message
            });
        }
    }
);

// ==================================================
// LOGIN
// ==================================================

app.post(
    '/login',
    async (req, res) => {

        try {

            console.log('');
            console.log(
                '===================================='
            );
            console.log(
                'LOGIN REQUEST'
            );
            console.log(
                '===================================='
            );

            const email =
                String(
                    req.body?.email || ''
                )
                    .trim()
                    .toLowerCase();

            const password =
                String(
                    req.body?.password || ''
                );

            console.log(
                'Login email:',
                email
            );

            // ------------------------------------------
            // CHECK INPUT
            // ------------------------------------------

            if (
                !email ||
                !password
            ) {

                return res.status(400).json({
                    message:
                        'Enter email and password.'
                });
            }

            // ------------------------------------------
            // CHECK EXCEL
            // ------------------------------------------

            if (
                !fs.existsSync(
                    EXCEL_FILE
                )
            ) {

                return res.status(400).json({
                    message:
                        'Email not registered.'
                });
            }

            // ------------------------------------------
            // OPEN EXCEL
            // ------------------------------------------

            const workbook =
                new ExcelJS.Workbook();

            await workbook.xlsx.readFile(
                EXCEL_FILE
            );

            const worksheet =
                workbook.getWorksheet(
                    'Users'
                ) ||
                workbook.worksheets[0];

            if (!worksheet) {

                return res.status(500).json({
                    message:
                        'Users database sheet not found.'
                });
            }

            // ------------------------------------------
            // FIND USER
            // ------------------------------------------

            let userRow = null;

            worksheet.eachRow(
                (row, rowNumber) => {

                    if (
                        rowNumber === 1
                    ) {
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

            // ------------------------------------------
            // EMAIL NOT FOUND
            // ------------------------------------------

            if (!userRow) {

                return res.status(400).json({
                    message:
                        'Email not registered.'
                });
            }

            // ------------------------------------------
            // PASSWORD HASH
            // ------------------------------------------

            const passwordHash =
                String(
                    userRow.getCell(9).value || ''
                );

            if (!passwordHash) {

                return res.status(500).json({
                    message:
                        'Password data not found. Please register again.'
                });
            }

            // ------------------------------------------
            // CHECK PASSWORD
            // ------------------------------------------

            let passwordCorrect =
                false;

            try {

                passwordCorrect =
                    await bcrypt.compare(
                        password,
                        passwordHash
                    );

            } catch (bcryptError) {

                console.error(
                    'BCRYPT ERROR:',
                    bcryptError
                );

                return res.status(500).json({
                    message:
                        'Password verification failed. Please register again.'
                });
            }

            // ------------------------------------------
            // WRONG PASSWORD
            // ------------------------------------------

            if (
                !passwordCorrect
            ) {

                return res.status(400).json({
                    message:
                        'Incorrect password.'
                });
            }

            // ------------------------------------------
            // CREATE SESSION
            // ------------------------------------------

            req.session.user =
                email;

            console.log(
                'LOGIN SUCCESS:',
                email
            );

            // ------------------------------------------
            // SEND SUCCESS
            // ------------------------------------------

            return res.status(200).json({
                message:
                    'Login successful.',
                redirect:
                    '/'
            });

        } catch (error) {

            console.error('');
            console.error(
                '===================================='
            );
            console.error(
                'LOGIN ERROR'
            );
            console.error(
                error
            );
            console.error(
                '===================================='
            );

            return res.status(500).json({
                message:
                    'Login failed: ' +
                    error.message
            });
        }
    }
);

// ==================================================
// PROTECTED MAIN PORTAL
// ==================================================

app.get(
    '/main-portal.html',
    (req, res) => {

        // Not logged in
        if (
            !req.session ||
            !req.session.user
        ) {

            return res.redirect('/');
        }

        // Logged in
        return res.sendFile(
            path.join(
                __dirname,
                'public',
                'main-portal.html'
            )
        );
    }
);

// ==================================================
// LOGOUT
// ==================================================

app.get(
    '/logout',
    (req, res) => {

        req.session.destroy(
            (error) => {

                if (error) {

                    console.error(
                        'LOGOUT ERROR:',
                        error
                    );

                    return res.status(500).send(
                        'Logout failed.'
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

// ==================================================
// 404
// ==================================================

app.use(
    (req, res) => {

        return res.status(404).send(
            'Page not found.'
        );
    }
);

// ==================================================
// START SERVER
// ==================================================

const PORT =
    process.env.PORT || 3000;

app.listen(
    PORT,
    '0.0.0.0',
    () => {

        console.log('');
        console.log(
            '===================================='
        );

        console.log(
            'JNTUA ACADEMIC HUB SERVER STARTED'
        );

        console.log(
            'PORT:',
            PORT
        );

        console.log(
            'DATA DIRECTORY:',
            DATA_DIR
        );

        console.log(
            'EXCEL FILE:',
            EXCEL_FILE
        );

        console.log(
            '===================================='
        );

        console.log('');
    }
);
