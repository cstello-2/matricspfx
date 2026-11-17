import { writeFileSync, mkdirSync, rmSync } from 'node:fs';
import { deflateRawSync, crc32 } from 'node:zlib';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const outDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'fixtures', 'lists');
mkdirSync(outDir, { recursive: true });

const todayArg = process.argv.indexOf('--today');
const TODAY = todayArg > 0 ? new Date(process.argv[todayArg + 1] + 'T00:00:00Z') : new Date(Date.UTC(new Date().getUTCFullYear(), new Date().getUTCMonth(), new Date().getUTCDate()));

let seed = 42;
const rnd = () => ((seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296);
const pick = (arr, weights) => {
  if (!weights) return arr[Math.floor(rnd() * arr.length)];
  const total = weights.reduce((a, b) => a + b, 0);
  let r = rnd() * total;
  for (let i = 0; i < arr.length; i++) { r -= weights[i]; if (r <= 0) return arr[i]; }
  return arr[arr.length - 1];
};
const int = (a, b) => a + Math.floor(rnd() * (b - a + 1));
const daysAgo = n => new Date(TODAY.getTime() - n * 86400000);
const iso = d => d.toISOString().replace(/\.\d{3}Z$/, 'Z');
const guid = () => 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { const r = Math.floor(rnd() * 16); return (c === 'x' ? r : (r & 3) | 8).toString(16); });

const names = ['Avery Hughes', 'Jordan Patel', 'Riley Chen', 'Sam Okafor', 'Taylor Brooks', 'Casey Nguyen', 'Morgan Silva', 'Jamie Novak', 'Drew Larsen', 'Quinn Romero', 'Robin Haddad', 'Skyler Kim'];
const users = names.map((Title, i) => {
  const email = Title.toLowerCase().replace(' ', '.') + '@contoso.dev';
  return { Id: 10 + i, Title, EMail: email, LoginName: `i:0#.f|membership|${email}`, Department: pick(['Engineering', 'Operations', 'Marketing', 'Finance']) };
});
writeFileSync(join(outDir, '..', 'users.json'), JSON.stringify(users, null, 2));

const sys = (id, created) => ({ Id: id, ID: id, Created: iso(created), Modified: iso(new Date(Math.min(TODAY.getTime(), created.getTime() + int(0, 20) * 86400000))), AuthorId: pick(users).Id });

const projectNames = ['Apollo CRM rollout', 'Billing platform v2', 'Customer portal', 'Data warehouse', 'Intranet refresh', 'Mobile app', 'Network upgrade', 'Onboarding automation', 'Payroll migration', 'Security hardening', 'Supplier portal', 'Teams telephony', 'Website relaunch', 'Zero-trust access'];
const departments = ['Engineering', 'Operations', 'Marketing', 'Finance', 'HR'];
const projects = projectNames.map((Title, i) => {
  const start = daysAgo(int(60, 330));
  const due = new Date(start.getTime() + int(90, 300) * 86400000);
  return {
    ...sys(i + 1, start),
    Title,
    Status: pick(['Active', 'Active', 'Active', 'At risk', 'On hold', 'Completed']),
    Department: pick(departments),
    Budget: int(4, 60) * 5000,
    OwnerId: pick(users).Id,
    StartDate: iso(start),
    DueDate: iso(due),
    Priority: pick(['Low', 'Medium', 'High'], [2, 5, 3]),
  };
});

const taskVerbs = ['Design', 'Build', 'Test', 'Review', 'Document', 'Migrate', 'Configure', 'Deploy', 'Plan', 'Validate'];
const taskObjects = ['API', 'data model', 'UI', 'permissions', 'reports', 'integration', 'pipeline', 'training', 'cutover', 'backlog'];
const tasks = [];
for (let i = 0; i < 420; i++) {
  const p = pick(projects);
  const created = new Date(Math.max(new Date(p.StartDate).getTime(), daysAgo(int(0, 200)).getTime()));
  let status = p.Status === 'Completed' ? 'Done' : pick(['Not started', 'In progress', 'Blocked', 'Done'], [2, 3, 1, 5]);
  const est = int(2, 40);
  const finish = new Date(created.getTime() + int(1, 45) * 86400000);
  if (status === 'Done' && finish > TODAY) status = 'In progress';
  const done = status === 'Done';
  const completed = done ? finish : null;
  tasks.push({
    ...sys(i + 1, created),
    Title: `${pick(taskVerbs)} ${pick(taskObjects)}`,
    ProjectId: p.Id,
    AssignedToId: pick(users).Id,
    Status: status,
    Priority: pick(['Low', 'Medium', 'High', 'Critical'], [3, 5, 3, 1]),
    EstimateHours: est,
    ActualHours: done ? Math.round(est * (0.6 + rnd() * 0.9)) : status === 'Not started' ? 0 : Math.round(est * rnd() * 0.8),
    CompletedDate: completed ? iso(completed) : null,
    Tags: [pick(['frontend', 'backend', 'infra', 'docs', 'qa'])].concat(rnd() > 0.6 ? [pick(['urgent', 'customer', 'tech-debt'])] : []),
  });
}

const categories = ['Hardware', 'Software', 'Access', 'Network', 'Other'];
const teamFor = { Hardware: 'Service Desk', Software: 'Applications', Access: 'Security', Network: 'Infrastructure', Other: 'Service Desk' };
const tickets = [];
for (let i = 0; i < 640; i++) {
  const created = daysAgo(int(0, 120));
  const category = pick(categories, [3, 5, 3, 2, 1]);
  const priority = pick(['Low', 'Normal', 'High', 'Urgent'], [3, 6, 2, 1]);
  const age = (TODAY.getTime() - created.getTime()) / 86400000;
  const resolvedAfter = (priority === 'Urgent' ? 0.3 : priority === 'High' ? 1 : 2.5) * (0.3 + rnd() * 3);
  const resolved = age > resolvedAfter && rnd() > 0.08;
  const status = resolved ? pick(['Resolved', 'Closed'], [1, 2]) : pick(['New', 'Assigned', 'In progress', 'Waiting on user'], [2, 2, 3, 1]);
  tickets.push({
    ...sys(i + 1, created),
    Title: `${category} issue #${1000 + i}`,
    Category: category,
    Priority: priority,
    Status: status,
    Channel: pick(['Email', 'Phone', 'Portal', 'Teams'], [3, 2, 4, 3]),
    Team: rnd() > 0.85 ? pick(['Service Desk', 'Infrastructure', 'Applications', 'Security']) : teamFor[category],
    AssignedToId: status === 'New' ? null : pick(users).Id,
    ResolvedDate: resolved ? iso(new Date(created.getTime() + resolvedAfter * 86400000)) : null,
    Satisfaction: resolved && rnd() > 0.4 ? pick([1, 2, 3, 4, 5], [1, 1, 3, 6, 7]) : null,
  });
}

const taskNotes = [
  'Acceptance criteria are agreed with the product owner. Remaining work is tracked in the linked tickets.',
  'Blocked on access to the staging environment. Escalated to the infrastructure team.',
  'Pairing with the vendor on this one. Expect a first draft by the end of the sprint.',
  'Scope reduced after the steering meeting: only the core flow is in this release.',
  'Needs a review from security before it can go live.',
];
const ticketNotes = [
  'User reports the issue started after the latest update. Screenshots attached in the original email.',
  'Affects the whole floor. Workaround: use the guest network for now.',
  'Requested access for a new starter in the finance team, manager approval received.',
  'Intermittent problem, could not reproduce on a second device.',
  'Caller asked for a call back before noon.',
];
const commentTexts = [
  'I had a look at this and it should be a quick fix.',
  'Can we get more details on how to reproduce this?',
  'Moved to this sprint, it is blocking two other items.',
  'Done on my side, ready for review.',
  'Waiting for feedback from the business owner.',
  'Thanks, confirmed this is working now.',
  'Raised with the vendor, ticket number in the description.',
];

for (const t of tasks) t.Description = rnd() > 0.2 ? pick(taskNotes) : null;
for (const t of tickets) t.Description = pick(ticketNotes);

function seedComments(items) {
  const out = {};
  let id = 1;
  for (const item of items) {
    if (rnd() > 0.3) continue;
    const created = new Date(item.Created).getTime();
    const count = int(1, 4);
    const list = [];
    for (let i = 0; i < count; i++) {
      const at = new Date(created + rnd() * Math.max(0, TODAY.getTime() - created));
      list.push({ id: id++, text: pick(commentTexts), createdDate: iso(at), authorId: pick(users).Id });
    }
    out[item.Id] = list.sort((a, b) => a.createdDate.localeCompare(b.createdDate));
  }
  return out;
}

const text = (title, extra = {}) => ({ type: 'Text', title, ...extra });
const choice = (title, choices, extra = {}) => ({ type: 'Choice', title, choices, ...extra });
const lists = [
  {
    title: 'Projects',
    id: guid(),
    fields: {
      Title: text('Title', { required: true }),
      Status: choice('Status', ['Active', 'At risk', 'On hold', 'Completed']),
      Department: choice('Department', departments),
      Budget: { type: 'Currency', title: 'Budget' },
      Owner: { type: 'User', title: 'Owner' },
      StartDate: { type: 'DateTime', title: 'Start date', dateOnly: true },
      DueDate: { type: 'DateTime', title: 'Due date', dateOnly: true },
      Priority: choice('Priority', ['Low', 'Medium', 'High']),
    },
    items: projects,
  },
  {
    title: 'Tasks',
    id: guid(),
    fields: {
      Title: text('Title', { required: true }),
      Description: { type: 'Note', title: 'Description' },
      Status: choice('Status', ['Not started', 'In progress', 'Blocked', 'Done'], { required: true }),
      Priority: choice('Priority', ['Low', 'Medium', 'High', 'Critical']),
      Project: { type: 'Lookup', title: 'Project', list: 'Projects', field: 'Title' },
      AssignedTo: { type: 'User', title: 'Assigned to' },
      EstimateHours: { type: 'Number', title: 'Estimate (h)' },
      ActualHours: { type: 'Number', title: 'Actual (h)' },
      CompletedDate: { type: 'DateTime', title: 'Completed', dateOnly: true },
      Tags: { type: 'MultiChoice', title: 'Tags', choices: ['frontend', 'backend', 'infra', 'docs', 'qa', 'urgent', 'customer', 'tech-debt'] },
    },
    items: tasks,
    comments: seedComments(tasks),
  },
  {
    title: 'Tickets',
    id: guid(),
    fields: {
      Title: text('Title', { required: true }),
      Description: { type: 'Note', title: 'Description' },
      Status: choice('Status', ['New', 'Assigned', 'In progress', 'Waiting on user', 'Resolved', 'Closed'], { required: true }),
      Priority: choice('Priority', ['Low', 'Normal', 'High', 'Urgent']),
      Category: choice('Category', categories),
      Channel: choice('Channel', ['Email', 'Phone', 'Portal', 'Teams']),
      Team: choice('Team', ['Service Desk', 'Infrastructure', 'Applications', 'Security']),
      AssignedTo: { type: 'User', title: 'Assigned to' },
      ResolvedDate: { type: 'DateTime', title: 'Resolved' },
      Satisfaction: { type: 'Number', title: 'Satisfaction' },
    },
    items: tickets,
    comments: seedComments(tickets),
  },
];
const filesDir = join(outDir, '..', 'files', 'Policies');
rmSync(filesDir, { recursive: true, force: true });
mkdirSync(filesDir, { recursive: true });

function zip(entries) {
  const locals = [];
  const central = [];
  let offset = 0;
  for (const [name, content] of entries) {
    const data = Buffer.from(content, 'utf8');
    const compressed = deflateRawSync(data);
    const nameBytes = Buffer.from(name, 'utf8');
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    locals.push(local, nameBytes, compressed);
    const dir = Buffer.alloc(46);
    dir.writeUInt32LE(0x02014b50, 0);
    dir.writeUInt16LE(20, 4);
    dir.writeUInt16LE(20, 6);
    dir.writeUInt16LE(8, 10);
    dir.writeUInt32LE(crc, 16);
    dir.writeUInt32LE(compressed.length, 20);
    dir.writeUInt32LE(data.length, 24);
    dir.writeUInt16LE(nameBytes.length, 28);
    dir.writeUInt32LE(offset, 42);
    central.push(dir, nameBytes);
    offset += 30 + nameBytes.length + compressed.length;
  }
  const centralSize = central.reduce((a, b) => a + b.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, ...central, end]);
}

const xmlEscape = t => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function docx(paragraphs) {
  const body = paragraphs.map(p => `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(p)}</w:t></w:r></w:p>`).join('');
  return zip([
    ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
    ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'],
    ['word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`],
  ]);
}

const documents = [
  { name: 'Remote work policy.docx', dept: 'HR', title: 'Remote work policy', summary: 'Rules for working from home and other locations.', text: ['Employees may work remotely up to three days per week with manager approval.', 'Home office equipment such as a monitor and chair can be expensed once every three years.', 'Use the company VPN on public networks and lock your laptop when you step away.'] },
  { name: 'Travel and expenses.docx', dept: 'Finance', title: 'Travel and expenses', summary: 'Booking travel and claiming expenses.', text: ['Book flights and hotels through the travel portal at least two weeks in advance.', 'Economy class is standard for flights under six hours.', 'Submit receipts for meals, taxis and mileage within 30 days to be reimbursed.'] },
  { name: 'Password and MFA standard.docx', dept: 'IT', title: 'Password and MFA standard', summary: 'Requirements for passwords and multi-factor authentication.', text: ['All accounts must use multi-factor authentication with the authenticator app.', 'Passwords must be at least 14 characters and must not be reused.', 'Report suspected phishing emails to the security team immediately.'] },
  { name: 'Onboarding checklist.md', dept: 'HR', title: 'Onboarding checklist', summary: 'First week tasks for new starters.', text: ['# Onboarding checklist', '- Collect your laptop and badge on day one', '- Complete security awareness training', '- Meet your buddy and book 1:1s with your team', '- Enrol in benefits and pension within 30 days'] },
  { name: 'Data retention schedule.txt', dept: 'Legal', title: 'Data retention schedule', summary: 'How long different records are kept.', text: ['Customer contracts are retained for seven years after expiry.', 'Personal data of former employees is deleted after two years under GDPR.', 'Backups are encrypted and kept for 35 days.'] },
  { name: 'Purchasing guidelines.docx', dept: 'Finance', title: 'Purchasing guidelines', summary: 'Approval limits for buying goods and services.', text: ['Purchases above 5,000 require a purchase order and budget owner approval.', 'Use preferred suppliers where a framework agreement exists.', 'Software purchases must be reviewed by IT for security and licensing.'] },
  { name: 'Incident response plan.md', dept: 'IT', title: 'Incident response plan', summary: 'What to do during a security incident.', text: ['# Incident response', 'Contain the incident, preserve evidence and notify the security team.', 'Major incidents involving personal data must be reported to the data protection officer within 24 hours.'] },
  { name: 'Parental leave policy.docx', dept: 'HR', title: 'Parental leave policy', summary: 'Leave entitlements for new parents.', text: ['All parents receive 16 weeks of paid leave after birth or adoption.', 'Leave can be taken in up to three blocks within the first year.', 'Talk to HR about flexible working when you return.'] },
  { name: 'Company car policy.pdf', dept: 'Finance', title: 'Company car policy', summary: 'Eligibility and rules for company cars and fuel cards.', text: null },
  { name: 'Acceptable use of IT.txt', dept: 'IT', title: 'Acceptable use of IT', summary: 'Using company devices, email and internet responsibly.', text: ['Company laptops and phones are for business use; limited personal use is allowed.', 'Do not install unapproved software or connect unknown USB devices.', 'Email and chat may be monitored for security purposes.'] },
  { name: 'Health and safety at the office.docx', dept: 'Facilities', title: 'Health and safety at the office', summary: 'Keeping the workplace safe.', text: ['Fire exits must be kept clear at all times.', 'Report accidents and near misses to facilities within 24 hours.', 'First aiders are listed on every floor next to the lifts.'] },
  { name: 'Code of conduct.md', dept: 'Legal', title: 'Code of conduct', summary: 'How we behave with colleagues, customers and suppliers.', text: ['# Code of conduct', 'Treat everyone with respect; harassment and discrimination are not tolerated.', 'Declare gifts above 50 and any conflicts of interest to your manager.'] },
];

const policies = documents.map((d, i) => {
  const created = daysAgo(int(5, 300));
  const file = join(filesDir, d.name);
  if (d.name.endsWith('.docx')) writeFileSync(file, docx(d.text));
  else if (d.text) writeFileSync(file, d.text.join('\n\n') + '\n');
  else writeFileSync(file, Buffer.from('%PDF-1.4\n% workbench placeholder\n'));
  return {
    ...sys(i + 1, created),
    Title: d.title,
    FileLeafRef: d.name,
    FileRef: `/sites/dev/Policies/${d.name}`,
    Summary: d.summary,
    Department: d.dept,
    Topics: i === 0 ? ['Remote work', 'HR'] : i === 1 ? ['Finance', 'Travel'] : [],
  };
});

lists.push({
  title: 'Policies',
  id: guid(),
  baseTemplate: 101,
  fields: {
    FileLeafRef: { type: 'File', title: 'Name', readOnly: true },
    Title: text('Title'),
    Summary: { type: 'Note', title: 'Summary' },
    Department: choice('Department', ['HR', 'Finance', 'IT', 'Legal', 'Facilities']),
    Topics: { type: 'MultiChoice', title: 'Topics', choices: ['HR', 'Finance', 'Travel', 'Security', 'Remote work', 'Privacy', 'Health and safety', 'Procurement', 'Conduct'] },
  },
  items: policies,
  comments: {},
});

for (const l of lists) writeFileSync(join(outDir, `${l.title}.json`), JSON.stringify(l, null, 2));
console.log(`Wrote ${lists.map(l => `${l.title} (${l.items.length})`).join(', ')} and ${users.length} users to ${outDir} (today = ${TODAY.toISOString().slice(0, 10)})`);
