
// ==========================================================
// JNTUA ACADEMIC HUB - server.js
// ==========================================================

require('dotenv').config();
 
const express = require('express');
const path = require('path');
const session = require('express-session');
const ExcelJS = require('exceljs');
const fs = require('fs');
const bcrypt = require('bcrypt');
const { createClient } = require('@supabase/supabase-js');

const app = express();


// ==========================================================
// TRUST RENDER PROXY
// ==========================================================

app.set('trust proxy', 1);


// ==========================================================
// SUPABASE PERMANENT DATABASE
// ==========================================================

const SUPABASE_URL =
    String(process.env.SUPABASE_URL || '').trim();

const SUPABASE_SECRET_KEY =
    String(process.env.SUPABASE_SECRET_KEY || '').trim();

let supabase = null;

if (SUPABASE_URL && SUPABASE_SECRET_KEY) {

    supabase = createClient(
        SUPABASE_URL,
        SUPABASE_SECRET_KEY,
        {
            auth: {
                autoRefreshToken: false,
                persistSession: false,
                detectSessionInUrl: false
            }
        }
    );

    console.log('SUPABASE: Connected');

} else {

    console.error(
        'SUPABASE ERROR: Missing SUPABASE_URL or SUPABASE_SECRET_KEY'
    );

}

console.log(
    'SUPABASE URL:',
    SUPABASE_URL || 'MISSING'
);

console.log(
    'SUPABASE KEY TYPE:',
    SUPABASE_SECRET_KEY.startsWith('sb_secret_')
        ? 'SECRET KEY'
        : SUPABASE_SECRET_KEY.startsWith('sb_publishable_')
            ? 'PUBLISHABLE KEY - WRONG FOR BACKEND'
            : SUPABASE_SECRET_KEY
                ? 'OTHER KEY'
                : 'MISSING'
);


// ==========================================================
// DIRECTORIES / FILE PATHS
// ==========================================================

const PUBLIC_DIR =
    path.join(__dirname, 'public');

const DATA_DIR =
    process.env.DATA_DIR ||
    path.join(__dirname, 'data');

const EXCEL_FILE =
    path.join(DATA_DIR, 'users.xlsx');


// ==========================================================
// CREATE DATA DIRECTORY IF NEEDED
// ==========================================================

if (!fs.existsSync(DATA_DIR)) {

    fs.mkdirSync(DATA_DIR, {
        recursive: true
    });

}


// ==========================================================
// MIDDLEWARE
// ==========================================================

app.use(express.json());

app.use(
    express.urlencoded({
        extended: true
    })
);


// ==========================================================
// SESSION
// ==========================================================

app.use(
    session({

        secret:
            process.env.SESSION_SECRET ||
            'jntua-secret',

        resave: false,

        saveUninitialized: false,

        cookie: {

            httpOnly: true,

            secure:
                process.env.NODE_ENV === 'production',

            sameSite: 'lax',

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
// SERVE PUBLIC FOLDER
// ==========================================================

app.use(
    express.static(PUBLIC_DIR)
);


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

    const workbook =
        new ExcelJS.Workbook();

    const worksheet =
        workbook.addWorksheet('Users');

    worksheet.addRow(HEADERS);

    await workbook.xlsx.writeFile(
        EXCEL_FILE
    );

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


    // ======================================================
    // CHECK DUPLICATE EMAIL
    // ======================================================

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
                existingEmail === newEmail
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


    // ======================================================
    // ADD USER
    // ======================================================

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
        'User registered in Excel:',
        newEmail
    );

}


// ==========================================================
// SAVE USER TO SUPABASE
// ==========================================================

async function saveUserToSupabase(user) {

    if (!supabase) {

        throw new Error(
            'Database configuration error'
        );

    }

    const {
        data: existingUser,
        error: checkError
    } = await supabase
        .from('users')
        .select('id')
        .eq('email', user.email)
        .maybeSingle();

    if (checkError) {

        console.error(
            'SUPABASE USER CHECK ERROR:',
            checkError
        );

        throw new Error(
            'Database error'
        );

    }

    if (existingUser) {

        throw new Error(
            'Email already registered'
        );

    }

    const {
        data,
        error
    } = await supabase
        .from('users')
        .insert([{

            name:
                user.name || '',

            email:
                user.email || '',

            password:
                user.password || '',

            regulation:
                user.regulation || '',

            roll_number:
                user.rollNumber || '',

            branch:
                user.branch || '',

            year:
                user.yearStudy || '',

            college_name:
                user.collegeName || '',

            phone:
                user.phone || ''

        }])
        .select()
        .single();

    if (error) {

        console.error(
            'SUPABASE INSERT ERROR:',
            error
        );

        throw new Error(
            'Database error'
        );

    }

    console.log(
        'USER SAVED PERMANENTLY IN SUPABASE:',
        data.email
    );

    return data;

}


// ==========================================================
// FIND USER FROM SUPABASE
// ==========================================================

async function findUserInSupabase(email) {

    if (!supabase) {

        return null;

    }

    const {
        data,
        error
    } = await supabase
        .from('users')
        .select('*')
        .eq('email', email)
        .maybeSingle();

    if (error) {

        console.error(
            'SUPABASE FIND USER ERROR:',
            error
        );

        return null;

    }

    return data || null;

}


// ==========================================================
// HOME ROUTE
// ==========================================================

app.get('/', (req, res) => {

    console.log(
        '--------------------------'
    );

    console.log(
        'GET /'
    );

    console.log(
        'Session:',
        req.session
            ? req.session.user
            : 'none'
    );


    // ======================================================
    // USER LOGGED IN
    // ======================================================

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

        if (!fs.existsSync(mainPage)) {

            console.error(
                'ERROR: main-portal.html NOT FOUND'
            );

            return res.status(500).send(
                'main-portal.html not found inside public folder.'
            );

        }

        return res.sendFile(
            mainPage
        );

    }


    // ======================================================
    // USER NOT LOGGED IN
    // ======================================================

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

    return res.sendFile(
        loginPage
    );

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
            String(
                user.email
            )
                .trim()
                .toLowerCase();


        // ==================================================
        // PASSWORD CONFIRMATION
        // ==================================================

        if (
            user.password !==
            user.confirmPassword
        ) {

            return res.status(400).json({

                message:
                    'Passwords do not match'

            });

        }


        // ==================================================
        // SUPABASE IS THE PERMANENT DATABASE
        // ==================================================

        if (!supabase) {

            return res.status(500).json({

                message:
                    'Database configuration error'

            });

        }


        // ==================================================
        // SAVE PERMANENTLY TO SUPABASE
        // ==================================================

        await saveUserToSupabase(user);


        // ==================================================
        // EXCEL BACKUP
        // ==================================================

        try {

            await addUserToExcel(user);

        }

        catch (excelError) {

            console.error(
                'EXCEL BACKUP ERROR:',
                excelError
            );

            // IMPORTANT:
            // Supabase already contains the permanent user.
            // Excel backup failure must NOT cancel registration.

        }


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


        if (
            err.message ===
            'Database configuration error'
        ) {

            return res.status(500).json({

                message:
                    'Database configuration error'

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


        // ==================================================
        // EMPTY FIELDS
        // ==================================================

        if (
            !email ||
            !password
        ) {

            return res.status(400).json({

                message:
                    'Enter email and password'

            });

        }


        // ==================================================
        // SUPABASE LOGIN - PRIMARY
        // ==================================================

        if (supabase) {

            const supabaseUser =
                await findUserInSupabase(email);


            if (supabaseUser) {

                const storedPassword =
                    String(
                        supabaseUser.password || ''
                    );


                // ==================================================
                // PASSWORD CHECK
                // ==================================================

                let passwordCorrect =
                    false;


                // Support existing plain-text passwords
                if (
                    storedPassword ===
                    password
                ) {

                    passwordCorrect =
                        true;

                }

                // Support bcrypt passwords if any exist
                else if (
                    storedPassword.startsWith('$2')
                ) {

                    try {

                        passwordCorrect =
                            await bcrypt.compare(
                                password,
                                storedPassword
                            );

                    }

                    catch (bcryptError) {

                        console.error(
                            'BCRYPT ERROR:',
                            bcryptError
                        );

                    }

                }


                if (!passwordCorrect) {

                    return res.status(400).json({

                        message:
                            'Incorrect password'

                    });

                }


                // ==================================================
                // LOGIN SUCCESS
                // ==================================================

                req.session.user =
                    email;


                console.log(
                    'SUPABASE LOGIN SUCCESS:',
                    email
                );


                return req.session.save(
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


                        return res.json({

                            message:
                                'Login successful',

                            redirect:
                                '/'

                        });

                    }
                );

            }

        }


        // ==================================================
        // EXCEL FALLBACK
        // ==================================================

        if (
            !fs.existsSync(EXCEL_FILE)
        ) {

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
            workbook.getWorksheet('Users') ||
            workbook.getWorksheet(1);


        if (!worksheet) {

            return res.status(500).json({

                message:
                    'Users sheet not found'

            });

        }


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

                    userRow =
                        row;

                }

            }
        );


        if (!userRow) {

            return res.status(400).json({

                message:
                    'Email not registered'

            });

        }


        // Password is column 9 in current Excel structure
        const storedPassword =
            String(
                userRow.getCell(9).value || ''
            );


        if (
            storedPassword !==
            password
        ) {

            return res.status(400).json({

                message:
                    'Incorrect password'

            });

        }


        req.session.user =
            email;


        console.log(
            'EXCEL LOGIN SUCCESS:',
            email
        );


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


        // ==================================================
        // NO LOGIN
        // ==================================================

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


        if (
            !fs.existsSync(mainPage)
        ) {

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

                loggedIn:
                    false

            });

        }


        return res.json({

            loggedIn:
                true,

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
// OWNER / ADMIN PANEL
// ==========================================================


// ----------------------------------------------------------
// ADMIN LOGIN PAGE
// ----------------------------------------------------------

app.get(
    '/owner-login',
    (req, res) => {

        // Already admin logged in
        if (
            req.session &&
            req.session.isAdmin === true
        ) {

            return res.redirect(
                '/admin/panel'
            );

        }


        // Prevent search engines from indexing
        res.set(
            'X-Robots-Tag',
            'noindex, nofollow'
        );


        return res.sendFile(
            path.join(
                __dirname,
                'private',
                'admin-login.html'
            )
        );

    }
);


// ----------------------------------------------------------
// ADMIN LOGIN
// ----------------------------------------------------------

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


            const adminEmail =
                String(
                    process.env.ADMIN_EMAIL || ''
                )
                    .trim()
                    .toLowerCase();


            const adminPasswordHash =
                String(
                    process.env.ADMIN_PASSWORD_HASH || ''
                );


            if (
                !email ||
                !password
            ) {

                return res.status(400).json({

                    message:
                        'Enter owner email and password'

                });

            }


            // ==================================================
            // OWNER EMAIL CHECK
            // ==================================================

            if (
                email !== adminEmail
            ) {

                return res.status(401).json({

                    message:
                        'Invalid owner credentials'

                });

            }


            // ==================================================
            // OWNER PASSWORD CHECK
            // ==================================================

            if (!adminPasswordHash) {

                console.error(
                    'ADMIN_PASSWORD_HASH is missing in Render Environment Variables'
                );

                return res.status(500).json({

                    message:
                        'Owner password is not configured on server'

                });

            }


            const passwordCorrect =
                await bcrypt.compare(
                    password,
                    adminPasswordHash
                );


            if (!passwordCorrect) {

                return res.status(401).json({

                    message:
                        'Invalid owner credentials'

                });

            }


            // ==================================================
            // ADMIN SESSION
            // ==================================================

            req.session.isAdmin =
                true;

            req.session.adminEmail =
                email;


            req.session.save(
                (err) => {

                    if (err) {

                        console.error(
                            'ADMIN SESSION ERROR:',
                            err
                        );

                        return res.status(500).json({

                            message:
                                'Session error'

                        });

                    }


                    console.log(
                        'OWNER LOGIN SUCCESS:',
                        email
                    );


                    return res.json({

                        message:
                            'Admin login successful',

                        redirect:
                            '/admin/panel'

                    });

                }
            );

        }

        catch (error) {

            console.error(
                'ADMIN LOGIN ERROR:',
                error
            );


            return res.status(500).json({

                message:
                    'Server error'

            });

        }

    }
);


// ----------------------------------------------------------
// ADMIN AUTHENTICATION MIDDLEWARE
// ----------------------------------------------------------

function requireAdmin(
    req,
    res,
    next
) {

    if (
        !req.session ||
        req.session.isAdmin !== true
    ) {

        return res.status(401).json({

            message:
                'Admin authentication required'

        });

    }


    next();

}


// ----------------------------------------------------------
// ADMIN PANEL PAGE
// ----------------------------------------------------------

app.get(
    '/admin/panel',
    (req, res) => {

        if (
            !req.session ||
            req.session.isAdmin !== true
        ) {

            return res.redirect(
                '/owner-login'
            );

        }


        res.set(
            'X-Robots-Tag',
            'noindex, nofollow'
        );


        return res.sendFile(
            path.join(
                __dirname,
                'private',
                'admin.html'
            )
        );

    }
);


// ----------------------------------------------------------
// GET REGISTERED USERS
// ----------------------------------------------------------

app.get(
    '/admin/api/users',
    requireAdmin,
    async (req, res) => {

        try {

            // ==================================================
            // SUPABASE USERS - PRIMARY SOURCE
            // ==================================================

            if (supabase) {

                const {
                    data: supabaseUsers,
                    error
                } = await supabase
                    .from('users')
                    .select(
                        'id, name, email, regulation, roll_number, branch, year, college_name, phone, created_at'
                    )
                    .order(
                        'created_at',
                        {
                            ascending: false
                        }
                    );


                if (!error) {

                    const users =
                        (supabaseUsers || [])
                            .map(
                                (user) => ({

                                    name:
                                        user.name || '',

                                    email:
                                        user.email || '',

                                    regulation:
                                        user.regulation || '',

                                    rollNumber:
                                        user.roll_number || '',

                                    branch:
                                        user.branch || '',

                                    yearStudy:
                                        user.year || '',

                                    collegeName:
                                        user.college_name || '',

                                    phone:
                                        user.phone || ''

                                })
                            );


                    return res.json({

                        total:
                            users.length,

                        users

                    });

                }


                console.error(
                    'SUPABASE ADMIN USERS ERROR:',
                    error
                );

            }


            // ==================================================
            // EXCEL FALLBACK
            // ==================================================

            if (
                !fs.existsSync(EXCEL_FILE)
            ) {

                return res.json({

                    total:
                        0,

                    users:
                        []

                });

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

                return res.status(500).json({

                    message:
                        'Users sheet not found'

                });

            }


            const users = [];


            worksheet.eachRow(
                (row, rowNumber) => {

                    // Skip header
                    if (
                        rowNumber === 1
                    ) {

                        return;

                    }


                    users.push({

                        name:
                            String(
                                row.getCell(1).value || ''
                            ),

                        email:
                            String(
                                row.getCell(2).value || ''
                            ),

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
                            )

                    });

                }
            );


            return res.json({

                total:
                    users.length,

                users

            });

        }

        catch (error) {

            console.error(
                'ADMIN USERS ERROR:',
                error
            );


            return res.status(500).json({

                message:
                    'Could not read users'

            });

        }

    }
);


// ----------------------------------------------------------
// DOWNLOAD USERS EXCEL FILE
// ----------------------------------------------------------

app.get(
    '/admin/download-users',
    requireAdmin,
    async (req, res) => {

        try {

            // ==================================================
            // DOWNLOAD FROM SUPABASE
            // ==================================================

            if (supabase) {

                const {
                    data: users,
                    error
                } = await supabase
                    .from('users')
                    .select('*')
                    .order(
                        'created_at',
                        {
                            ascending: false
                        }
                    );


                if (!error) {

                    const workbook =
                        new ExcelJS.Workbook();


                    const worksheet =
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

                        'Phone'

                    ]);


                    (users || []).forEach(
                        (user) => {

                            worksheet.addRow([

                                user.name || '',

                                user.email || '',

                                user.regulation || '',

                                user.roll_number || '',

                                user.branch || '',

                                user.year || '',

                                user.college_name || '',

                                user.phone || ''

                            ]);

                        }
                    );


                    worksheet.columns.forEach(
                        (column) => {

                            column.width =
                                20;

                        }
                    );


                    res.setHeader(
                        'Content-Type',
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    );


                    res.setHeader(
                        'Content-Disposition',
                        'attachment; filename="JNTUA-Academic-Hub-Users.xlsx"'
                    );


                    await workbook.xlsx.write(
                        res
                    );


                    return res.end();

                }


                console.error(
                    'SUPABASE EXCEL DOWNLOAD ERROR:',
                    error
                );

            }


            // ==================================================
            // EXCEL FILE FALLBACK
            // ==================================================

            if (
                !fs.existsSync(EXCEL_FILE)
            ) {

                return res.status(404).send(
                    'Users Excel file not found'
                );

            }


            return res.download(
                EXCEL_FILE,
                'JNTUA-Academic-Hub-Users.xlsx',
                (error) => {

                    if (error) {

                        console.error(
                            'EXCEL DOWNLOAD ERROR:',
                            error
                        );

                    }

                }
            );

        }

        catch (error) {

            console.error(
                'EXCEL DOWNLOAD ERROR:',
                error
            );

            return res.status(500).send(
                'Excel download failed'
            );

        }

    }
);


// ----------------------------------------------------------
// ADMIN LOGOUT
// ----------------------------------------------------------

app.get(
    '/admin/logout',
    (req, res) => {

        req.session.isAdmin =
            false;

        req.session.adminEmail =
            null;


        req.session.save(
            (saveError) => {

                if (saveError) {

                    console.error(
                        'ADMIN LOGOUT ERROR:',
                        saveError
                    );

                }


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
