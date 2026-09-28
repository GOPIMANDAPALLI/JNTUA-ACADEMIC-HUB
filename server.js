// ==========================================================
// JNTUA ACADEMIC HUB - COMPLETE SERVER
// User Login + Registration + Admin Panel
// ==========================================================

const express = require('express');
const path = require('path');
const session = require('express-session');
const ExcelJS = require('exceljs');
const fs = require('fs');
const bcrypt = require('bcrypt');

const app = express();


// ==========================================================
// PATHS
// ==========================================================

const PUBLIC_DIR = path.join(__dirname, 'public');

const PRIVATE_DIR = path.join(__dirname, 'private');


// ==========================================================
// DATA DIRECTORY
// ==========================================================

// Local:
// ./data/users.xlsx
//
// Render:
// Set DATA_DIR=/var/data
//
// If DATA_DIR is not set,
// ./data will be used.

const DATA_DIR =
    process.env.DATA_DIR ||
    path.join(__dirname, 'data');


// Create data directory automatically

if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, {
        recursive: true
    });
}


const EXCEL_FILE =
    path.join(DATA_DIR, 'users.xlsx');


// ==========================================================
// MIDDLEWARE
// ==========================================================

app.use(express.json());

app.use(
    express.urlencoded({
        extended: true
    })
);


// IMPORTANT:
// Only public folder is exposed.
//
// private folder is NOT exposed by express.static().

app.use(
    express.static(PUBLIC_DIR)
);


// ==========================================================
// SESSION
// ==========================================================

app.use(
    session({

        secret:
            process.env.SESSION_SECRET ||
            'jntua-secret-change-this',

        resave: false,

        saveUninitialized: false,

        cookie: {

            httpOnly: true,

            secure:
                process.env.NODE_ENV === 'production',

            sameSite: 'lax',

            // 30 days
            maxAge:
                1000 *
                60 *
                60 *
                24 *
                30

        }

    })
);


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


    const workbook =
        new ExcelJS.Workbook();


    const worksheet =
        workbook.addWorksheet('Users');


    worksheet.addRow(HEADERS);


    // Header formatting

    worksheet.getRow(1).font = {
        bold: true
    };


    await workbook.xlsx.writeFile(
        EXCEL_FILE
    );


    console.log(
        'Created Excel:',
        EXCEL_FILE
    );

}


// ==========================================================
// FIND USER BY EMAIL
// ==========================================================

async function findUserByEmail(email) {

    await createExcelFile();


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


            if (
                storedEmail ===
                email
            ) {

                foundUser = {

                    rowNumber,

                    name:
                        String(
                            row.getCell(1).value || ''
                        ),

                    email:
                        storedEmail,

                    regulation:
                        String(
                            row.getCell(3).value || ''
                        ),

                    rollNumber:
                        String(
                            row.getCell(4).value || ''
                        ),

                    branch:
                        String(
                            row.getCell(5).value || ''
                        ),

                    yearStudy:
                        String(
                            row.getCell(6).value || ''
                        ),

                    collegeName:
                        String(
                            row.getCell(7).value || ''
                        ),

                    phone:
                        String(
                            row.getCell(8).value || ''
                        ),

                    passwordHash:
                        String(
                            row.getCell(9).value || ''
                        ),

                    registeredAt:
                        String(
                            row.getCell(10).value || ''
                        )

                };

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
        String(
            user.email || ''
        )
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


            if (
                existingEmail ===
                newEmail
            ) {

                emailExists = true;

            }

        }
    );


    if (emailExists) {

        throw new Error(
            'Email already registered'
        );

    }


    // Hash password

    const passwordHash =
        await bcrypt.hash(
            String(user.password),
            12
        );


    const registeredAt =
        new Date().toISOString();


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

        registeredAt

    ]);


    // Format columns

    worksheet.getColumn(1).width = 25;
    worksheet.getColumn(2).width = 35;
    worksheet.getColumn(3).width = 15;
    worksheet.getColumn(4).width = 20;
    worksheet.getColumn(5).width = 15;
    worksheet.getColumn(6).width = 15;
    worksheet.getColumn(7).width = 30;
    worksheet.getColumn(8).width = 18;
    worksheet.getColumn(9).width = 70;
    worksheet.getColumn(10).width = 30;


    await workbook.xlsx.writeFile(
        EXCEL_FILE
    );


    console.log(
        'USER REGISTERED:',
        newEmail
    );

}


// ==========================================================
// HOME PAGE
// ==========================================================

app.get('/', (req, res) => {

    console.log(
        'GET /'
    );


    // Already logged in

    if (
        req.session &&
        req.session.user
    ) {

        return res.sendFile(
            path.join(
                PUBLIC_DIR,
                'main-portal.html'
            )
        );

    }


    // Not logged in

    return res.sendFile(
        path.join(
            PUBLIC_DIR,
            'login.html'
        )
    );

});


// ==========================================================
// REGISTER
// ==========================================================

app.post(
    '/register',
    async (req, res) => {

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


            const email =
                String(
                    user.email
                )
                    .trim()
                    .toLowerCase();


            // Email validation

            const emailRegex =
                /^[^\s@]+@[^\s@]+\.[^\s@]+$/;


            if (
                !emailRegex.test(email)
            ) {

                return res.status(400).json({

                    message:
                        'Invalid email ID'

                });

            }


            // Password confirmation

            if (
                String(user.password) !==
                String(user.confirmPassword)
            ) {

                return res.status(400).json({

                    message:
                        'Passwords do not match'

                });

            }


            // Add to Excel

            await addUserToExcel({

                name:
                    String(user.name).trim(),

                email,

                regulation:
                    user.regulation || '',

                rollNumber:
                    user.rollNumber || '',

                branch:
                    user.branch || '',

                yearStudy:
                    user.yearStudy || '',

                collegeName:
                    user.collegeName || '',

                phone:
                    user.phone || '',

                password:
                    String(user.password)

            });


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

    }
);


// ==========================================================
// USER LOGIN
// ==========================================================

app.post(
    '/login',
    async (req, res) => {

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

            if (
                !email ||
                !password
            ) {

                return res.status(400).json({

                    message:
                        'Enter email and password'

                });

            }


            // Email format

            const emailRegex =
                /^[^\s@]+@[^\s@]+\.[^\s@]+$/;


            if (
                !emailRegex.test(email)
            ) {

                return res.status(400).json({

                    message:
                        'Invalid email ID'

                });

            }


            // Find user

            const user =
                await findUserByEmail(
                    email
                );


            if (!user) {

                return res.status(400).json({

                    message:
                        'Email not registered'

                });

            }


            // Compare password

            const passwordMatch =
                await bcrypt.compare(
                    password,
                    user.passwordHash
                );


            if (!passwordMatch) {

                return res.status(400).json({

                    message:
                        'Incorrect password'

                });

            }


            // Login success

            req.session.user =
                user.email;


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
                        'USER LOGIN SUCCESS:',
                        user.email
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

    }
);


// ==========================================================
// PROTECTED MAIN PORTAL
// ==========================================================

app.get(
    '/main-portal.html',
    (req, res) => {

        if (
            !req.session ||
            !req.session.user
        ) {

            return res.redirect('/');

        }


        return res.sendFile(
            path.join(
                PUBLIC_DIR,
                'main-portal.html'
            )
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
// USER LOGOUT
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
// ADMIN AUTHENTICATION
// ==========================================================

function requireAdmin(req, res, next) {

    if (
        req.session &&
        req.session.admin === true
    ) {

        return next();

    }


    return res.status(401).json({

        message:
            'Admin authentication required'

    });

}


// ==========================================================
// SECRET ADMIN ENTRY URL
// ==========================================================

// IMPORTANT:
// This URL is NOT displayed anywhere
// on the main website.
//
// Example:
// http://localhost:3000/owner-panel-x7k9
//
// Production:
// https://yourdomain.com/owner-panel-x7k9

app.get(
    '/owner-panel-x7k9',
    (req, res) => {

        // Already admin logged in

        if (
            req.session &&
            req.session.admin === true
        ) {

            return res.sendFile(
                path.join(
                    PRIVATE_DIR,
                    'admin.html'
                )
            );

        }


        // Admin login page

        return res.sendFile(
            path.join(
                PRIVATE_DIR,
                'admin-login.html'
            )
        );

    }
);


// ==========================================================
// ADMIN LOGIN
// ==========================================================

app.post(
    '/admin/login',
    async (req, res) => {

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


            if (
                !email ||
                !password
            ) {

                return res.status(400).json({

                    message:
                        'Enter admin email and password'

                });

            }


            const adminEmail =
                String(
                    process.env.ADMIN_EMAIL || ''
                )
                    .trim()
                    .toLowerCase();


            const adminPasswordHash =
                process.env.ADMIN_PASSWORD_HASH;


            if (
                !adminEmail ||
                !adminPasswordHash
            ) {

                console.error(
                    'ADMIN_EMAIL or ADMIN_PASSWORD_HASH is missing.'
                );


                return res.status(500).json({

                    message:
                        'Admin configuration is missing'

                });

            }


            // Check email

            if (
                email !==
                adminEmail
            ) {

                return res.status(401).json({

                    message:
                        'Invalid admin credentials'

                });

            }


            // Check password

            const passwordMatch =
                await bcrypt.compare(
                    password,
                    adminPasswordHash
                );


            if (!passwordMatch) {

                return res.status(401).json({

                    message:
                        'Invalid admin credentials'

                });

            }


            // Admin session

            req.session.admin =
                true;

            req.session.adminEmail =
                adminEmail;


            req.session.save(
                (sessionError) => {

                    if (sessionError) {

                        console.error(
                            'ADMIN SESSION ERROR:',
                            sessionError
                        );

                        return res.status(500).json({

                            message:
                                'Admin session error'

                        });

                    }


                    return res.json({

                        message:
                            'Admin login successful',

                        redirect:
                            '/owner-panel-x7k9'

                    });

                }
            );

        }

        catch (err) {

            console.error(
                'ADMIN LOGIN ERROR:',
                err
            );


            return res.status(500).json({

                message:
                    'Server error'

            });

        }

    }
);


// ==========================================================
// ADMIN USERS API
// ==========================================================

app.get(
    '/admin/users',
    requireAdmin,
    async (req, res) => {

        try {

            await createExcelFile();


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


            const users = [];


            worksheet.eachRow(
                (row, rowNumber) => {

                    if (rowNumber === 1) {
                        return;
                    }


                    const email =
                        String(
                            row.getCell(2).value || ''
                        )
                            .trim()
                            .toLowerCase();


                    if (!email) {
                        return;
                    }


                    users.push({

                        name:
                            String(
                                row.getCell(1).value || ''
                            ),

                        email,

                        regulation:
                            String(
                                row.getCell(3).value || ''
                            ),

                        roll_number:
                            String(
                                row.getCell(4).value || ''
                            ),

                        branch:
                            String(
                                row.getCell(5).value || ''
                            ),

                        year_of_study:
                            String(
                                row.getCell(6).value || ''
                            ),

                        college_name:
                            String(
                                row.getCell(7).value || ''
                            ),

                        phone:
                            String(
                                row.getCell(8).value || ''
                            ),

                        registered_at:
                            String(
                                row.getCell(10).value || ''
                            )

                    });

                }
            );


            return res.json({

                users

            });

        }

        catch (err) {

            console.error(
                'ADMIN USERS ERROR:',
                err
            );


            return res.status(500).json({

                message:
                    'Failed to load users'

            });

        }

    }
);


// ==========================================================
// ADMIN EXCEL EXPORT
// ==========================================================

app.get(
    '/admin/export-users',
    requireAdmin,
    async (req, res) => {

        try {

            await createExcelFile();


            const workbook =
                new ExcelJS.Workbook();


            await workbook.xlsx.readFile(
                EXCEL_FILE
            );


            const worksheet =
                workbook.getWorksheet('Users') ||
                workbook.getWorksheet(1);


            if (!worksheet) {

                return res.status(500).send(
                    'Users sheet not found'
                );

            }


            const exportWorkbook =
                new ExcelJS.Workbook();


            const exportSheet =
                exportWorkbook.addWorksheet(
                    'Users'
                );


            exportSheet.columns = [

                {
                    header: 'Name',
                    key: 'name',
                    width: 25
                },

                {
                    header: 'Email',
                    key: 'email',
                    width: 35
                },

                {
                    header: 'Regulation',
                    key: 'regulation',
                    width: 15
                },

                {
                    header: 'Roll Number',
                    key: 'roll_number',
                    width: 20
                },

                {
                    header: 'Branch',
                    key: 'branch',
                    width: 15
                },

                {
                    header: 'Year',
                    key: 'year_of_study',
                    width: 15
                },

                {
                    header: 'College',
                    key: 'college_name',
                    width: 30
                },

                {
                    header: 'Phone',
                    key: 'phone',
                    width: 18
                },

                {
                    header: 'Registered At',
                    key: 'registered_at',
                    width: 30
                }

            ];


            worksheet.eachRow(
                (row, rowNumber) => {

                    if (rowNumber === 1) {
                        return;
                    }


                    const email =
                        String(
                            row.getCell(2).value || ''
                        );


                    if (!email) {
                        return;
                    }


                    exportSheet.addRow({

                        name:
                            String(
                                row.getCell(1).value || ''
                            ),

                        email,

                        regulation:
                            String(
                                row.getCell(3).value || ''
                            ),

                        roll_number:
                            String(
                                row.getCell(4).value || ''
                            ),

                        branch:
                            String(
                                row.getCell(5).value || ''
                            ),

                        year_of_study:
                            String(
                                row.getCell(6).value || ''
                            ),

                        college_name:
                            String(
                                row.getCell(7).value || ''
                            ),

                        phone:
                            String(
                                row.getCell(8).value || ''
                            ),

                        registered_at:
                            String(
                                row.getCell(10).value || ''
                            )

                    });

                }
            );


            exportSheet.getRow(1).font = {
                bold: true
            };


            res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            );


            res.setHeader(
                'Content-Disposition',
                'attachment; filename="JNTUA-users.xlsx"'
            );


            await exportWorkbook.xlsx.write(
                res
            );


            res.end();

        }

        catch (err) {

            console.error(
                'EXCEL EXPORT ERROR:',
                err
            );


            return res.status(500).send(
                'Excel export failed'
            );

        }

    }
);


// ==========================================================
// ADMIN LOGOUT
// ==========================================================

app.post(
    '/admin/logout',
    (req, res) => {

        if (!req.session) {

            return res.json({
                message:
                    'Admin logged out'
            });

        }


        req.session.admin = false;

        req.session.adminEmail = null;


        req.session.save(
            () => {

                return res.json({

                    message:
                        'Admin logged out'

                });

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
            '======================================'
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
            '======================================'
        );

    }
);
