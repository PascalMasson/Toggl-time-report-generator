import { Toggl } from 'toggl-track';
import 'dotenv/config';
import * as XLSX from 'xlsx';
import yargs from 'yargs/yargs';
import nodemailer from 'nodemailer';
import readline from 'readline';
import open, {openApp, apps} from 'open';
import fs from 'fs';
import path from 'path';

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

function getCommandLine() {
    switch (process.platform) { 
      case 'darwin' : return 'open';
      case 'win32' : return 'start';
      case 'win64' : return 'start';
      default : return 'xdg-open';
   }
}

// Function to send email
const sendEmail = (filename, weeknum) => {
    const transporter = nodemailer.createTransport({
        host: process.env.EMAIL_HOST,
        port: process.env.EMAIL_PORT,
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASS,
        },
    });

    const mailOptions = {
        from: process.env.EMAIL_USER,
        to: process.env.EMAIL_TO,
        cc: process.env.EMAIL_CC_SELF ? process.env.EMAIL_TO : undefined,
        subject: 'Hours for week ' + weeknum,
        text: 'I attached the hours for week ' + weeknum,
        attachments: [
            {
                path: filename
            }
        ]
    };

    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            console.log('Error sending email:', error);
        } else {
            console.log('Email sent:', info.response);
        }
    });
};
const getWeekNumber = (date) => {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
};

const formatDate = (date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;

const getMonday = (date) => {
    const today = date ? new Date(date) : new Date();
    const day = today.getDay();
    const diff = today.getDate() - day + (day === 0 ? -6 : 1);
    return new Date(today.setDate(diff));
};

const argv = yargs(process.argv.slice(2))
    .option('last', { describe: 'Generate report for the last week', type: 'boolean' })
    .option('current', { describe: 'Generate report for the current week', type: 'boolean' })
    .option('week', { describe: 'Generate report for a specific week', type: 'string' })
    .help()
    .alias('help', 'h')
    .argv;

(async () => {
    const toggl = new Toggl({ auth: { token: process.env.TOGGL_TRACK_API_TOKEN } });
    let firstDay, lastDay;

    if (argv.current || argv.last) {
        let targetMonday = getMonday();
        if (argv.last) targetMonday.setDate(targetMonday.getDate() - 7);
        const targetSunday = new Date(targetMonday);
        targetSunday.setDate(targetSunday.getDate() + 7);
        [firstDay, lastDay] = [formatDate(targetMonday), formatDate(targetSunday)];
    } else if (argv.week) {
        console.log(argv.week, new Date(argv.week), getMonday(new Date(argv.week)));
        const specifiedMonday = getMonday(new Date(argv.week));
        const specifiedSunday = new Date(specifiedMonday);
        specifiedSunday.setDate(specifiedSunday.getDate() + 7);
        [firstDay, lastDay] = [formatDate(specifiedMonday), formatDate(specifiedSunday)];
    } else {
        console.log("Please provide either --last, --current, or --week option.");
        return;
    }
    console.log(`Getting entries from ${firstDay} to ${lastDay}`);

    const entries = await toggl.timeEntry.list({
        startDate: firstDay,
        endDate: lastDay,
    });

    const workspaceId = process.env.TOGGL_WORKSPACE_ID;

    const projects = await toggl.projects.list(workspaceId);

    const allClients = await toggl.request("https://api.track.toggl.com/api/v9/workspaces/" + workspaceId + "/clients");

    const groupedEntries = entries.reduce((acc, entry) => {
        const date = entry.start.split('T')[0];
        const project = projects.find(project => project.id === entry.pid);
        if (project.cid != null) return acc;
        if (!acc[date]) {
            acc[date] = {};
        }
        if (!acc[date][project.name]) {
            acc[date][project.name] = [];
        }
        acc[date][project.name].push(entry);
        return acc;
    }, {});

    const totalTimes = Object.entries(groupedEntries).reduce((acc, [date, projects]) => {
        Object.entries(projects).forEach(([project, entries]) => {
            if (!project.includes(' - ')) {
                const total = entries.reduce((acc, entry) => {
                    const start = new Date(entry.start);
                    const end = new Date(entry.stop);
                    const duration = (end - start) / 1000 / 60 / 60;
                    return acc + duration;
                }, 0);
                if (!acc[project]) {
                    acc[project] = {};
                }
                acc[project][date] = total.toFixed(2);
            }
        });
        return acc;
    }, {});

    const projectsWithClients = projects.filter(project => project.cid);

    const entriesWithClients = entries.filter(entry => projectsWithClients.some(project => project.id === entry.pid));

    const groupedEntriesWithClients = entriesWithClients.reduce((acc, entry) => {
        const date = entry.start.split('T')[0];
        const client = projectsWithClients.find(project => project.id === entry.pid).cid;
        const description = entry.description;
        if (!acc[date]) {
            acc[date] = {};
        }
        if (!acc[date][client]) {
            acc[date][client] = {};
        }
        if (!acc[date][client][description]) {
            acc[date][client][description] = [];
        }
        acc[date][client][description].push(entry);
        return acc;
    }, {});

    const totalTimesWithClients = Object.entries(groupedEntriesWithClients).reduce((acc, [date, clients]) => {
        Object.entries(clients).forEach(([client, descriptions]) => {
            const clientName = allClients.find(x => x.id.toString() === client).name;
            Object.entries(descriptions).forEach(([description, entries]) => {
                const total = entries.reduce((acc, entry) => {
                    const start = new Date(entry.start);
                    const end = new Date(entry.stop);
                    const duration = (end - start) / 1000 / 60 / 60;
                    return acc + duration;
                }, 0);
                if (!acc[clientName]) {
                    acc[clientName] = {};
                }
                if (!acc[clientName][date]) {
                    acc[clientName][date] = {};
                }
                acc[clientName][date][description] = total.toFixed(2);
            });
        });
        return acc;
    }, {});

    const totalTimesMerged = {
        ...totalTimes,
        ...totalTimesWithClients
    };

    const data = Object.entries(totalTimesMerged).reduce((acc, [project, dates]) => {
        Object.entries(dates).forEach(([date, descriptions]) => {
            if (typeof descriptions === 'string') {
                acc.push({
                    'Day of week': new Date(date).toLocaleDateString('en-US', {
                        weekday: 'long'
                    }),
                    'Date': date,
                    'Hours': descriptions.toString().replace('.', ','),
                    'For': project,
                    'Task': '',
                });
            } else {
                Object.entries(descriptions).forEach(([description, hours]) => {
                    const day = new Date(date).toLocaleDateString('en-US', {
                        weekday: 'long'
                    });
                    acc.push({
                        'Day of week': day,
                        'Date': date,
                        'Hours': hours.toString().replace('.', ','),
                        'For': project,
                        'Task': description,
                    });
                });
            }
        });
        return acc;
    }, []).sort((a, b) => new Date(a.Date) - new Date(b.Date));

    const sheet = XLSX.utils.json_to_sheet(data);
    const totalHours = Object.values(totalTimesMerged).reduce((acc, dates) => {
        return acc + Object.values(dates).reduce((acc, descriptions) => {
            if (typeof descriptions === 'string') {
                return acc + parseFloat(descriptions);
            }
            return acc + Object.values(descriptions).reduce((acc, hours) => acc + parseFloat(hours), 0);
        }, 0);
    }, 0).toFixed(2).replace('.', ',');

    const totalRow = {
        'Day of week': 'Total',
        'Date': '',
        'Hours': totalHours,
        'For': '',
        'Task': '',
    };

    const emptyRow = {
        'Day of week': '',
        'Date': '',
        'Hours': '',
        'For': '',
        'Task': '',
    };

    XLSX.utils.sheet_add_json(sheet, [emptyRow, totalRow], {
        skipHeader: true,
        origin: -1
    });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');

    const worksheet = workbook.Sheets['Sheet1'];
    const columnWidths = [{
            wch: 15
        },
        {
            wch: 15
        },
        {
            wch: 10
        },
        {
            wch: 30
        },
        {
            wch: 30
        },
    ];
    worksheet['!cols'] = columnWidths;

     const weekNumber = getWeekNumber(new Date(firstDay));
    const year = new Date(firstDay).getFullYear().toString().slice(-2);
    const filename = `hours/20${year}/Urenverantwoording P. Masson week ${weekNumber} ${year}-01.xlsx`;
    //create directory if it doesn't exist
    if (!fs.existsSync(path.dirname(filename))) {
        fs.mkdirSync(path.dirname(filename), { recursive: true });
    }
    console.log(`Writing to ${filename}`);
    XLSX.writeFile(workbook, filename);
		
	await open(filename);
	
	// Ask the user if they want to send the email
	rl.question('Do you want to send the generated file by email? (yes/no): ', (answer) => {
		if (answer.toLowerCase() === 'yes') {
			sendEmail(filename, weekNumber);
		}
		rl.close();
	});
})();