// ==========================================================
// JNTUA ACADEMIC HUB - server.js
// ==========================================================

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
// SUPABASE
// ==========================================================

const SUPABASE_URL =
    process.env.SUPABASE_URL || '';

const SUPABASE_SERVICE_ROLE_KEY =
    process.env.SUPABASE_SERVICE_ROLE_KEY || '';

let supabase = null;

if (
    SUPABASE_URL &&
    SUPABASE_SERVICE_ROLE_KEY
) {

    supabase = createClient(
        SUPABASE_URL,
        SUPABASE_SERVICE_ROLE_KEY
    );

    console.log(
        'SUPABASE: Connected'
    );

} else {

    console.error(
        'SUPABASE ERROR: SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY is missing'
    );

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

app.use(session({

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

}));


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
//
// Excel is kept for compatibility/download.
// PERMANENT user storage is now SUPABASE.
//
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
        'User also added to Excel:',
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

        if (!supabase) {

            return res.status(500).json({

                message:
                    'Database is not configured on server'

            });

        }


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
        // CHECK EMAIL IN SUPABASE
        // ==================================================

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
                'SUPABASE REGISTER CHECK ERROR:',
                checkError
            );


            return res.status(500).json({

                message:
                    'Database error. Try later.'

            });

        }


        // ==================================================
        // EMAIL ALREADY EXISTS
        // ==================================================

        if (existingUser) {

            return res.status(400).json({

                message:
                    'Email already registered. Please login.'

            });

        }


        // ==================================================
        // SAVE USER PERMANENTLY IN SUPABASE
        // ==================================================

        const {
            data: insertedUser,
            error: insertError
        } = await supabase
            .from('users')
            .insert([

                {

                    name:
                        String(
                            user.name || ''
                        ).trim(),

                    email:
                        user.email,

                    password:
                        user.password,

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

                }

            ])
            .select()
            .single();


        if (insertError) {

            console.error(
                'SUPABASE REGISTER ERROR:',
                insertError
            );


            // Duplicate email protection
            if (
                insertError.code === '23505'
            ) {

                return res.status(400).json({

                    message:
                        'Email already registered. Please login.'

                });

            }


            return res.status(500).json({

                message:
                    'Registration failed. Try later.'

            });

        }


        console.log(
            'USER SAVED PERMANENTLY IN SUPABASE:',
            insertedUser.email
        );


        // ==================================================
        // OPTIONAL EXCEL COPY
        // ==================================================
        //
        // If Excel write fails, registration is STILL
        // permanently saved in Supabase.
        //
        // ==================================================

        try {

            await addUserToExcel(user);

        }

        catch (excelError) {

            console.error(
                'EXCEL COPY ERROR:',
                excelError
            );

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

        if (!supabase) {

            return res.status(500).json({

                message:
                    'Database is not configured on server'

            });

        }


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
        // FIND USER IN PERMANENT SUPABASE DATABASE
        // ==================================================

        const {
            data: user,
            error: loginError
        } = await supabase
            .from('users')
            .select('*')
            .eq('email', email)
            .maybeSingle();


        if (loginError) {

            console.error(
                'SUPABASE LOGIN ERROR:',
                loginError
            );


            return res.status(500).json({

                message:
                    'Database error. Try later.'

            });

        }


        // ==================================================
        // EMAIL DOESN'T EXIST
        // ==================================================

        if (!user) {

            return res.status(400).json({

                message:
                    'Email not registered'

            });

        }


        // ==================================================
        // PASSWORD CHECK
        // ==================================================

        if (
            user.password !== password
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


        // ==================================================
        // SAVE SESSION FIRST
        // ==================================================

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


                // ==================================================
                // TELL FRONTEND TO REDIRECT
                // ==================================================

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

            req.session.isAdmin = true;

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
//
// IMPORTANT:
// Users are now loaded from SUPABASE.
// Therefore Render restart/sleep does NOT erase them.
//
// ----------------------------------------------------------

app.get(
    '/admin/api/users',
    requireAdmin,
    async (req, res) => {

        try {

            if (!supabase) {

                return res.status(500).json({

                    message:
                        'Database is not configured'

                });

            }


            const {
                data: users,
                error
            } = await supabase
                .from('users')
                .select(`
                    id,
                    name,
                    email,
                    regulation,
                    roll_number,
                    branch,
                    year,
                    college_name,
                    phone,
                    created_at
                `)
                .order(
                    'created_at',
                    {
                        ascending: false
                    }
                );


            if (error) {

                console.error(
                    'ADMIN SUPABASE USERS ERROR:',
                    error
                );


                return res.status(500).json({

                    message:
                        'Could not read users'

                });

            }


            const formattedUsers =
                (users || []).map(
                    (user) => {

                        return {

                            id:
                                user.id,

                            name:
                                String(
                                    user.name || ''
                                ),

                            email:
                                String(
                                    user.email || ''
                                ),

                            regulation:
                                String(
                                    user.regulation || ''
                                ),

                            rollNumber:
                                String(
                                    user.roll_number || ''
                                ),

                            branch:
                                String(
                                    user.branch || ''
                                ),

                            yearStudy:
                                String(
                                    user.year || ''
                                ),

                            collegeName:
                                String(
                                    user.college_name || ''
                                ),

                            phone:
                                String(
                                    user.phone || ''
                                ),

                            createdAt:
                                user.created_at || ''

                        };

                    }
                );


            return res.json({

                total:
                    formattedUsers.length,

                users:
                    formattedUsers

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
//
// Excel is generated from SUPABASE.
// Therefore downloaded Excel always contains
// the latest registered users.
//
// ----------------------------------------------------------

app.get(
    '/admin/download-users',
    requireAdmin,
    async (req, res) => {

        try {

            if (!supabase) {

                return res.status(500).send(
                    'Database is not configured'
                );

            }


            const {
                data: users,
                error
            } = await supabase
                .from('users')
                .select(`
                    name,
                    email,
                    regulation,
                    roll_number,
                    branch,
                    year,
                    college_name,
                    phone,
                    created_at
                `)
                .order(
                    'created_at',
                    {
                        ascending: true
                    }
                );


            if (error) {

                console.error(
                    'DOWNLOAD USERS ERROR:',
                    error
                );


                return res.status(500).send(
                    'Could not load users'
                );

            }


            const workbook =
                new ExcelJS.Workbook();


            const worksheet =
                workbook.addWorksheet(
                    'Users'
                );


            worksheet.columns = [

                {
                    header: 'Name',
                    key: 'name',
                    width: 25
                },

                {
                    header: 'Email',
                    key: 'email',
                    width: 30
                },

                {
                    header: 'Regulation',
                    key: 'regulation',
                    width: 15
                },

                {
                    header: 'RollNumber',
                    key: 'roll_number',
                    width: 20
                },

                {
                    header: 'Branch',
                    key: 'branch',
                    width: 15
                },

                {
                    header: 'YearOfStudy',
                    key: 'year',
                    width: 15
                },

                {
                    header: 'CollegeName',
                    key: 'college_name',
                    width: 35
                },

                {
                    header: 'Phone',
                    key: 'phone',
                    width: 18
                },

                {
                    header: 'RegisteredAt',
                    key: 'created_at',
                    width: 25
                }

            ];


            (users || []).forEach(
                (user) => {

                    worksheet.addRow({

                        name:
                            user.name || '',

                        email:
                            user.email || '',

                        regulation:
                            user.regulation || '',

                        roll_number:
                            user.roll_number || '',

                        branch:
                            user.branch || '',

                        year:
                            user.year || '',

                        college_name:
                            user.college_name || '',

                        phone:
                            user.phone || '',

                        created_at:
                            user.created_at || ''

                    });

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


            res.end();

        }

        catch (error) {

            console.error(
                'EXCEL DOWNLOAD ERROR:',
                error
            );


            if (!res.headersSent) {

                return res.status(500).send(
                    'Excel generation failed'
                );

            }

        }

    }
);


// ----------------------------------------------------------
// ADMIN LOGOUT
// ----------------------------------------------------------

app.get(
    '/admin/logout',
    (req, res) => {

        req.session.isAdmin = false;

        req.session.adminEmail = null;


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
            `SUPABASE: ${
                supabase
                    ? 'CONNECTED'
                    : 'NOT CONFIGURED'
            }`
        );


        console.log(
            '===================================='
        );

    }
);
