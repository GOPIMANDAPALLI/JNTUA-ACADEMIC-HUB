// ==========================================================
// JNTUA ACADEMIC HUB - server.js
// ==========================================================

const express = require('express');
const path = require('path');
const session = require('express-session');
const ExcelJS = require('exceljs');
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
// If DATA_DIR is not set, local data folder is used.

const DATA_DIR =
    process.env.DATA_DIR || path.join(__dirname, 'data');

const EXCEL_FILE = path.join(DATA_DIR, 'users.xlsx');


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


// Serve CSS, JS, images, etc. from public
app.use(express.static(PUBLIC_DIR));


// ==========================================================
// SESSION
// ==========================================================

app.use(session({

    secret:
        process.env.SESSION_SECRET ||
        'jntua-secret',

    resave: false,

    saveUninitialized: false,

    cookie: {
    httpOnly: true,
    secure: false,
    sameSite: 'lax',
    maxAge: 1000 * 60 * 60 * 24 * 30
}

}));


// ==========================================================
// EXCEL HEADER
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
    'Password'
];


// ==========================================================
// CREATE EXCEL FILE IF NEEDED
// ==========================================================

async function createExcelFile() {

    if (fs.existsSync(EXCEL_FILE)) {
        return;
    }

    const workbook = new ExcelJS.Workbook();

    const worksheet =
        workbook.addWorksheet('Users');

    worksheet.addRow(HEADERS);

    await workbook.xlsx.writeFile(EXCEL_FILE);

    console.log(
        'Created Excel file:',
        EXCEL_FILE
    );
}


// ==========================================================
// ADD USER TO EXCEL
// ==========================================================

async function addUserToExcel(user) {

    await createExcelFile();

    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.readFile(EXCEL_FILE);

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


            if (existingEmail === newEmail) {
                emailExists = true;
            }

        }
    );


    if (emailExists) {
        throw new Error(
            'Email already registered'
        );
    }


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

        user.password || ''

    ]);


    await workbook.xlsx.writeFile(
        EXCEL_FILE
    );


    console.log(
        'User registered:',
        newEmail
    );
}


// ==========================================================
// HOME ROUTE
// ==========================================================
//
// /
//
// If logged in:
//     main-portal.html
//
// If not logged in:
//     login.html
//
// ==========================================================

app.get('/', (req, res) => {

    console.log('--------------------------');
    console.log('GET /');
    console.log(
        'Session:',
        req.session ? req.session.user : 'none'
    );


    // User logged in
    if (
        req.session &&
        req.session.user
    ) {

        console.log(
            'Opening MAIN PORTAL for:',
            req.session.user
        );


        const mainPage =
            path.join(
                PUBLIC_DIR,
                'main-portal.html'
            );


        // Check whether main page exists
        if (!fs.existsSync(mainPage)) {

            console.error(
                'ERROR: main-portal.html NOT FOUND'
            );

            return res.status(500).send(
                'main-portal.html not found inside public folder.'
            );
        }


        return res.sendFile(mainPage);
    }


    // User not logged in
    console.log(
        'Opening LOGIN PAGE'
    );


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


        user.email =
            String(user.email)
                .trim()
                .toLowerCase();


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


        // Excel doesn't exist
        if (!fs.existsSync(EXCEL_FILE)) {

            return res.status(400).json({

                message:
                    'No users registered yet'

            });

        }


        // Read Excel
        const workbook =
            new ExcelJS.Workbook();

        await workbook.xlsx.readFile(
            EXCEL_FILE
        );


        const worksheet =
            workbook.getWorksheet('Users') ||
            workbook.getWorksheet(1);


        if (!worksheet) {

            return res.status(500).json({

                message:
                    'Users sheet not found'

            });

        }


        // Find email
        let userRow = null;


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


                if (
                    storedEmail === email
                ) {

                    userRow = row;

                }

            }
        );


        // Email doesn't exist
        if (!userRow) {

            return res.status(400).json({

                message:
                    'Email not registered'

            });

        }


        // Password is column 9
        const storedPassword =
            String(
                userRow.getCell(9).value || ''
            );


        // Wrong password
        if (
            storedPassword !== password
        ) {

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


        // Save session first
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
                    'SESSION SAVED'
                );


                // Tell frontend to redirect
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


        // No login
        if (
            !req.session ||
            !req.session.user
        ) {

            console.log(
                'No session -> LOGIN'
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

